VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "StringBuilder with no API/TLB or VB6 only VBA functions."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' ------------------------------------------------------------
' Name:         StringWorker
' Purpose:      An aid to managing string data.
' Author:       Dave Carter.
' Date:         Tuesday 20 July 2010
' -------------------------------------------------------------------
' Update:       Wednesday 08 June 2011
'               Added Name and Tag as Public Strings
'               just to help give something unique to the string worker
'               in case it is in an array...
' ------------------------------------------------------------
' Update:
'               Thursday 07 April 2011
'               Buggered up pDeriveLinesArray before so tried correcting this time around.
'
' -------------------------------------------------------------------
' Update:       Saturday 12 March 2011
'               Attempt to add instructions to read the start positions of each line.
'               + ReadLinesArray
'               - pDeriveLinesArray
'               This has meant adding a new array of longs, mLinesArray.
'               Running this will cause the main string to shrink to size,
'               because of do today, think tomorrow code approach.
'
'               Have also added
'               + FindAllMatches
'               to help search the string for no. of substring occurances.
'
'
' -------------------------------------------------------------------
' Update:       Monday 25 October 2010
'               Attempt to implement my own vb6 only vba methods (no TLB or API).
'
'               FromArray (VBA.Join), went with the best I could manage.
'               pReverseString: reverse the user string characters (byte oriented)
'               StrReverse: String (VBA.StrReverse): like vba, reverses the string passed (use Reverse to reverse user string only).
'               FindBackwards: String (VBA.InstrRev).
'               FindCharBackwards: String (Same as FindBackwards but for single char find).
'
' -------------------------------------------------------------------

Option Explicit

'   Notes:  This began as an exercise following use of vbAccelerator's StringBuilder Class and fixing the memory bug;
'           I wanted to see if I could knock up something similar, performance wise, without
'           using rtlMoveMemory but with the same idea of string pre-allocation, using
'           Mid$ to shift the text around.  Tests showed that sometimes Steve's StringBuilder came out on
'           top and sometimes mine did, overall I still think Steve's is faster and I know VBSpeed can go faster still.
'           Thinking I hadn't made such a bad job of it I carried on adding methods that I thought I would like
'           to have available; this class is the result.
'           It attempts a few optimisations against the vba string functions and where these aren't likely
'           without api/typelib it falls back on vba; the more data the better the optimisations perform.
' -------------------------------------------------------------------
'           The vba JOIN function is called FromArray and the vba SPLIT function is called ToArray.
' -------------------------------------------------------------------
'           Fine tune concatenation using Pre-Allocation Size (defaults to a whopping 8192 characters unless changed: c_def_PreAllocationSize),
'           for instance, the more data the larger the pre-allocation sixe could be and vice-versa.
' -------------------------------------------------------------------
'
' ... Methods.
'   ... Public.
'
' S   Append.            ' ... Append a string value to the user string.
' S   AppendAsLine.      ' ... Append a string value to the user string with optional new line before and after.
' F   BufferLength: Long.  ' ... Returns the current length of the internal string buffer.
' F   CharCount: Long.   ' ... Returns the number of times a character exists in the user text.
' S   Compact.           ' ... Reduce the available buffer to the size of the current user string.
' S   DeleteAll.         ' ... Deletes the current user string from the buffer.
' S   DeletePortion.     ' ... Deletes a portion of text from the user string by start and length to remove.
' S   Filter.            ' ... Filters character/s from the user string.
' F   Find: Long.        ' ... Returns the starting position of the next occurance (from pStart) of the search string within the user string.
' S   FromArray.         ' ... Uses VBA.Join to convert a string array into a single string.
' S   FromFile.          ' ... Reads the contents of a text file and makes it to the user string, overwriting anything else.
' S   Insert.            ' ... Inserts a text string at the position specified.
' S   PrefixWith.        ' ... Adds a string to the beginning of the user string.
' S   Replace.           ' ... Replaces a search string within the current user string with an alternative string.
' S   Reverse.           ' ... Reverses the current user string.
' S   ToArray.           ' ... Convert the current user string into a string array by a given delimiter.
' S   ToFile.            ' ... Writes the contents of the current user string within the buffer to disk.
' S   ToLowerCase.       ' ... Converts the user text string to lower case characters using LCase$.
' S   ToProperCase.      ' ... Converts the user text string to proper case characters using StrConv with vbProperCase.
' S   ToUpperCase.       ' ... Converts the user text string to upper case characters using UCase$.
' S   Trim.              ' ... Trim trailing and leading Space characters from the current user string.
' S   TrimLeft.          ' ... Trim leading Space characters from the current user string.
' S   TrimRight.         ' ... Trim trailing Space characters from the current user string.
' F   WordCount: Long.   ' ... Counts and returns the number of words within the current user string.
'
'   ... Private.
' F   pCheckFileName: Boolean.  ' ... Checks a file name for invalid characters.
' S   pCryptScribe.      ' ... Manages the encryption / decryption of a string.
' S   pDeMystify.        ' ... Converts encrypted text to a readable string via simple reverse encryption algorithm.
' F   pFileExists: Boolean.  ' ... Tests for file existence, returns True if found else False.
' S   pMystify.          ' ... Converts text to an unreadable string via simple encryption algorithm.
' F   pReadTextFile: String.  ' ... Reads text from a file into the current user string.
' S   pResizeBuffer.     ' ... Manages changing the size of the internal buffer.
' S   pSplitB04.         ' ... Optimised Split method to return the user string as an array of string data.
' S   pWriteTextFile.    ' ... Writes the current user string as text to a text file.
'
' ... Properties.
'   ... Public.
'
' P   Length: Long.      ' ... Returns the length of the current user string (no. of characters).
' P   PreAllocationSize: Long.  ' ... Sets / Returns the size (no. of Space Characters) to add to the buffer when it is not large enough to add the string passed.
' P   TheString: String.  ' ... Sets / Returns the user string.
' -------------------------------------------------------------------

Private m_TheString As String ' ... private field for property TheString.
Private m_PreAllocationSize As Long ' ... (previously m_TheBufferChunkSize) private field for property PreAllocationSize.
Private m_TheBufferLength As Long
Private m_TheTextLength As Long ' ... private field for property Length.
Private m_Row As Long                   ' ... private field for looping through array.

Private mLinesArray() As Long
Private mTheDelimiter As String

Private Const c_def_PreAllocationSize As Long = 8192 ' ... default value for property PreAllocationSize.
Private Const c_def_TheString As String = vbNullString ' ... default value for property TheString.

' ... update, removing space char from invalid file name chars, v6 (code browser).
'Private Const c_def_InvalidFileNameChars As String = "`!$%^&*()-=+[]{}'#@~;,.<>/?\| " ' ... needs " as well, see pCheckFileName.
Private Const c_def_InvalidFileNameChars As String = "`!$%^&*()-=+[]{}'#@~;,.<>/?\|" ' ... needs " as well, see pCheckFileName.

Public Name As String
Public Tag As String

' -------------------------------------------------------------------

Public Sub ReadLinesArray(Optional ByVal pTheDelimiter As String = vbCrLf, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
    pDeriveLinesArray pTheDelimiter, pOK, pErrMsg
End Sub

Private Sub pDeriveLinesArray(Optional ByVal pTheDelimiter As String = vbCrLf, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)

Dim lngDelLength As Long
Dim lngTextLength As Long
Dim c As Long
Dim tmp As Long
' -------------------------------------------------------------------
' ... Helper: read a string as lines of text and return an array of line starting positions.
' ... This has got to be thanks to Chris Lucas from his Split04 on VBSpeed, Thank You Chris.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    pResizeBuffer -1 ' ... compact the main string before commencing.
    ' -------------------------------------------------------------------
    pOK = False ' ... set return to false in case it comes in as true to start with.
    mTheDelimiter = pTheDelimiter
    ' -------------------------------------------------------------------
    lngDelLength = Len(mTheDelimiter)
    lngTextLength = Len(m_TheString)
    ' -------------------------------------------------------------------
    If lngTextLength = 0 Or lngDelLength = 0 Then
        Err.Raise vbObjectError + 1002, , "Parse For Lines Array: not a valid string or delimiter length."
    End If
    
    ' -------------------------------------------------------------------
    ' ... begin processing.
    ReDim mLinesArray(0 To lngTextLength + 2)

    mLinesArray(0) = 1 ' ... first pos is 1.
    c = 1

    tmp = InStr(m_TheString, mTheDelimiter)

    Do While tmp
        mLinesArray(c) = tmp + lngDelLength
        tmp = InStr(mLinesArray(c), m_TheString, mTheDelimiter)
        c = c + 1
    Loop

    mLinesArray(c + 1) = lngTextLength ' ... ubound(lines) - 1 is last line in text while ubound(lines) is total length of text.

    ' -------------------------------------------------------------------
    ReDim Preserve mLinesArray(c + 1)
    ' ... end processing
    ' -------------------------------------------------------------------
    pOK = True

ResumeError:

Exit Sub
ErrHan:
    pOK = False
    pErrMsg = Err.Description
    ReDim mLinesArray(0)
    Debug.Print "StringWorker.pDeriveLinesArray: " & Err.Description
    Resume ResumeError:

End Sub      ' ... pDeriveLinesArray.

' Sub:             Append
' Description:     Append a string value to the user string.

Public Sub Append(ByRef pTheString As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Append.VB_Description = "Append a string value to the user string."

'... Parameters.
'    R__ pTheString: String          ' ... The string value to append.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim lTheStringLength As Long
Dim lTheTextLength As Long
' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    Let lTheStringLength = Len(pTheString)
    Let lTheTextLength = m_TheTextLength    ' ... save original user text length, pResizeBuffer will increase this by Len(pTheString).
    
    Call pResizeBuffer(lTheStringLength, pOK, pErrMsg)

    If pOK = True Then
        
        ' -------------------------------------------------------------------
        ' ... if we error here pOK will be set to False and pErrMsg will receive Err.Description within the error trap.
        ' ... using Mid$ rather than Mid as is faster, to insert the new text into the buffer.
        ' -------------------------------------------------------------------
        If lTheTextLength = 0 Then
            
            If lTheStringLength > 0 Then
                Mid$(m_TheString, 1, lTheStringLength) = pTheString
            End If
            
        Else
            
            Mid$(m_TheString, lTheTextLength + 1, lTheStringLength) = pTheString
        
        End If
                
    End If


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.Append", Err.Number, Err.Description

End Sub ' ... Append.

' Sub:             AppendAsLine
' Description:     Append a string value to the user string with optional new line before and after.

Public Sub AppendAsLine(ByRef pTheString As String, Optional ByVal pBefore As Boolean = False, Optional ByVal pAfter As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute AppendAsLine.VB_Description = "Append a string value to the user string with optional new line before and after."

'... Parameters.
'    R__ pTheString: String          ' ... The string value to append.
'    VO_ pBefore: Boolean            ' ... Add a New Line infront of the text when True.
'    VO_ pAfter: Boolean             ' ... Add a New Line after the text when True.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    ' -------------------------------------------------------------------
    ' ... delegate to Append adding a new lines to the user string as required.
    Call Append(IIf(pBefore, vbNewLine, vbNullString) & pTheString & IIf(pAfter, vbNewLine, vbNullString), pOK, pErrMsg)
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.AppendAsLine", Err.Number, Err.Description

End Sub ' ... AppendAsLine.

' Function:        BufferLength
' Returns:         Long.
' Description:     Returns the current length of the internal string buffer.

Public Function BufferLength() As Long
Attribute BufferLength.VB_Description = "Returns the current length of the internal string buffer."

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    Let BufferLength = m_TheBufferLength
    ' -------------------------------------------------------------------
    
Exit Function
ErrHan:

    Debug.Print "StringWorker.BufferLength", Err.Number, Err.Description

End Function ' ... BufferLength: Long.

' Function:        CharCount
' Returns:         Long.
' Description:     Returns the number of times a character exists in the user text.

Public Function CharCount(ByRef pSearchFor As String, Optional ByRef pCompareMethod As VbCompareMethod = vbBinaryCompare, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute CharCount.VB_Description = "Returns the number of times a character exists in the user text."

'... Parameters.
'    R__ pSearchFor: String          ' ... The character/s to search for within the user text.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim lngCount As Long ' ... a return value to this function.
Dim lngStartPos As Long           ' ... .
Dim lngCharLength As Long
Dim sTmp As String
Dim sQuery As String
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... set up locals.
    lngCharLength = Len(pSearchFor)
    ' ... skip if no search char or query string.
    If m_TheTextLength = 0 Or lngCharLength = 0 Then
        Exit Function
    End If
    ' -------------------------------------------------------------------
    sTmp = Me
    sQuery = pSearchFor
    ' -------------------------------------------------------------------
    If pCompareMethod = vbTextCompare Then
        sTmp = LCase$(sTmp)
        sQuery = LCase$(sQuery)
    End If
    ' -------------------------------------------------------------------
    ' ... initialise search, first occurance position.
    lngStartPos = InStr(sTmp, pSearchFor) ' ... lngStartPos = InStrB(sTmp, pSearchFor) ' if sure char codes are byte sized).
    ' -------------------------------------------------------------------
    ' ... if lngStartPos > 0 then increment count and find next occurance.
    Do Until lngStartPos = 0
        lngCount = lngCount + 1
        lngStartPos = InStr(lngStartPos + lngCharLength, sTmp, pSearchFor)
    Loop
    ' -------------------------------------------------------------------
    
ErrResume:
    
    ' -------------------------------------------------------------------
    ' ... return count.
    CharCount = lngCount
    ' -------------------------------------------------------------------

Exit Function
ErrHan:

    lngCount = 0
    pErrMsg = Err.Description
    pOK = False
    Debug.Print "StringWorker.CharCount", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... CharCount: Long.

Private Sub Class_Initialize()

    ' -------------------------------------------------------------------
    ' ... initialise local fields.
    m_PreAllocationSize = c_def_PreAllocationSize
    m_TheString = c_def_TheString
    ' -------------------------------------------------------------------

End Sub ' ... Class_Initialize:

Private Sub Class_Terminate()
    
    ' -------------------------------------------------------------------
    ' ... clear local fields.
    m_PreAllocationSize = 0
    m_TheString = vbNullString
    ' -------------------------------------------------------------------
    Erase mLinesArray
    ' -------------------------------------------------------------------

End Sub ' ... Class_Terminate:

' Sub:             Compact
' Description:     Reduce the available buffer to the size of the current user string.

Public Sub Compact(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Compact.VB_Description = "Reduce the available buffer to the size of the current user string."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    pResizeBuffer -1, pOK, pErrMsg
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Debug.Print "StringWorker.Compact", Err.Number, Err.Description

End Sub ' ... Compact.

' Sub:             DeleteAll
' Description:     Deletes the current user string from the buffer.

Public Sub DeleteAll(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute DeleteAll.VB_Description = "Deletes the current user string from the buffer."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    ' -------------------------------------------------------------------
    m_TheString = vbNullString
    m_TheTextLength = 0
    m_TheBufferLength = 0
    ' -------------------------------------------------------------------
    pOK = True
    ' -------------------------------------------------------------------
    
Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "StringWorker.DeleteAll", Err.Number, Err.Description

End Sub ' ... DeleteAll.

' Sub:             DeletePortion
' Description:     Deletes a portion of text from the user string by start and length to remove.

Public Sub DeletePortion(ByVal pStart As Long, ByVal pLength As Long, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute DeletePortion.VB_Description = "Deletes a portion of text from the user string by start and length to remove."

'... Parameters.
'    V__ pStart: Long                ' ... The position from which to begin the delete.
'    V__ pLength: Long               ' ... The length of the text (in characters) to delete from the user text.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    If m_TheTextLength > 0 And pLength >= 1 Then
        
        ' -------------------------------------------------------------------
        ' ... ensure start position is at least 1.
        If pStart < 1 Then Let pStart = 1
        ' -------------------------------------------------------------------
        If pStart > m_TheTextLength Then
            ' -------------------------------------------------------------------
            ' ... start > length, trim to user string length.
            Compact pOK, pErrMsg
            ' -------------------------------------------------------------------
        Else
            ' -------------------------------------------------------------------
            ' ... if start and delete length > length reduce pLength to fit.
            If pStart + pLength > m_TheTextLength Then
                Let pLength = m_TheTextLength - pStart + 1
            End If
            ' -------------------------------------------------------------------
            ' ... shift remaining text following start and del. len. to start.
            Mid$(m_TheString, pStart, m_TheTextLength - pLength) = Mid$(m_TheString, pStart + pLength, m_TheTextLength - pLength)
            ' -------------------------------------------------------------------
            ' ... cut user string text length to suit; note shifting didn't delete from the length.
            m_TheTextLength = m_TheTextLength - pLength
            ' -------------------------------------------------------------------
            pOK = True
            ' -------------------------------------------------------------------
        End If
    Else
        ' -------------------------------------------------------------------
        pErrMsg = "Either there is nothing to delete from or nothing to delete."
        ' -------------------------------------------------------------------
    End If

ErrResume:


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.DeletePortion", Err.Number, Err.Description
    Resume ErrResume:


End Sub ' ... DeletePortion.

Function BytesSize() As Double
'
End Function

' Sub:             Filter
' Description:     Filters character/s from the user string.

Public Sub Filter(ByRef pTheFilter As String, Optional ByRef pCountFiltered As Long = 0, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Filter.VB_Description = "Filters character/s from the user string."

'... Parameters.
'    R__ pTheFilter: String          ' ... The string to filter from the user string.
'    RO_ pCountFiltered: Long        ' ... The string to filter from the user string.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    Me.Replace pTheFilter, vbNullString, 1, pCountFiltered, pOK, pErrMsg

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.Filter", Err.Number, Err.Description

End Sub ' ... Filter.

' Function:        Find
' Returns:         Long.
' Description:     Returns the starting position of the next occurance (from pStart) of the search string within the user string.

Public Function Find(ByRef pTheString As String, _
            Optional ByRef pStart As Long = 1, _
            Optional ByRef pCompareMethod As VbCompareMethod = vbBinaryCompare, _
            Optional ByVal pWholeWordOnly As Boolean = False, _
            Optional ByRef pOK As Boolean = False, _
            Optional ByRef pErrMsg As String = vbNullString) As Long

'... Parameters.
'    R__ pTheString: String          ' ... The string value to find.
'    RO_ pStart: Long                ' ... The position from which to begin the search.
'    RO_ pCompareMethod: vbCompareMethod ' ... The vbCompareMethod member for the search comparision.
'    VO_ pWholeWordOnly: Boolean     ' ... only count items identified as words.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim lngCharLength As Long
Dim lFoundPos As Long ' ... a return value to this function.
Dim sSource As String
Dim sQuery As String
Dim lngStart As Long
Dim lngCharLeft As Long
Dim lngCharRight As Long
Dim bDoItAgain As Boolean
Dim bIsWord As Boolean
'Dim bControlChar As Boolean

Const c_CharMax As Long = 65
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    lngCharLength = Len(pTheString)
    
    ' -------------------------------------------------------------------
    ' ... test there is something to search in and for and that the search and search + start are not longer than the source.
    If m_TheTextLength > 0 And lngCharLength > 0 Then 'And ((lngCharLength <= m_TheTextLength) And (lngCharLength + pStart <= m_TheTextLength)) Then
        
        ' -------------------------------------------------------------------
        ' ... ensure start position is at least 1.
        If pStart < 1 Then Let pStart = 1
        
        ' -------------------------------------------------------------------
        ' ... check that we are within limits.
        If (lngCharLength <= m_TheTextLength) And (lngCharLength + pStart <= m_TheTextLength) Then
            
            sSource = Me    ' ... user string, ensures we do not need to test find pos in legal length.
            sQuery = pTheString
            lngStart = pStart
            
            If pCompareMethod = vbTextCompare Then
                ' -------------------------------------------------------------------
                ' ... converting to lower case and using binary compare is faster than text compare alone.
                ' -------------------------------------------------------------------
                sSource = LCase$(sSource)
                sQuery = LCase$(sQuery)
            End If
            
            Do
                                
                bDoItAgain = False
                
                lFoundPos = InStr(lngStart, sSource, sQuery, vbBinaryCompare)
                
                If lFoundPos > 0 Then
                    If pWholeWordOnly Then
                        ' -------------------------------------------------------------------
                        ' ... sum chars either side of query string.
                        ' -------------------------------------------------------------------
                        lngCharLeft = 0: lngCharRight = 0
                        If lFoundPos > 1 Then
                            lngCharLeft = Asc(Mid$(sSource, lFoundPos - 1, 1))
                            If pbIsWordBreakChar(lngCharLeft) Then lngCharLeft = 0
'                bControlChar = lngCharLeft < 48 Or lngCharLeft > 57 And lngCharLeft < 65 Or lngCharLeft > 90 And lngCharLeft < 96 Or lngCharLeft > 122 And lngCharLeft < 127
'                If bControlChar Then lngCharLeft = 0
                            
                        End If
                        If m_TheTextLength > lFoundPos + lngCharLength Then
                            lngCharRight = Asc(Mid$(sSource, lFoundPos + lngCharLength, 1))
                            If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
'                bControlChar = lngCharRight < 48 Or lngCharRight > 57 And lngCharRight < 65 Or lngCharRight > 90 And lngCharRight < 96 Or lngCharRight > 122 And lngCharRight < 127
'                If bControlChar Then lngCharRight = 0
                            
                        End If
                        ' -------------------------------------------------------------------
                        ' ... control char word identification, Left Max 32 + Right Max 32 = Max 64
                        ' -------------------------------------------------------------------
                        bIsWord = lngCharLeft + lngCharRight < c_CharMax
                        If bIsWord = False Then
                            lngStart = lngStart + lngCharLength
                            bDoItAgain = True
                        End If
                    End If
                End If
                
            Loop While bDoItAgain = True
                        
        End If
        
    Else
        ' -------------------------------------------------------------------
        pErrMsg = "Either there is nothing to search in or nothing to search for."
        ' -------------------------------------------------------------------
    End If
    
    Let Find = lFoundPos


Exit Function
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.Find", Err.Number, Err.Description

End Function ' ... Find: Long.

'Public Function FindAllMatches(ByVal pFind As String, pPositionArray() As Long, Optional pStart As Long = 0, Optional pCompare As VbCompareMethod = vbBinaryCompare, Optional ByRef pErrMsg As String = vbnullString) As Long
'
'    FindAllMatches = modStrings.FindAllMatches(TheString, pFind, pPositionArray, pStart, pCompare, pErrMsg)
'
'End Function


Public Function FindAllMatches(ByVal pFind As String, _
                               ByRef pPositionArray() As Long, _
                      Optional ByRef pStart As Long = 0, _
                      Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
                      Optional ByVal pWholeWordOnly As Boolean = False, _
                      Optional ByRef pErrMsg As String = vbNullString) As Long

' ... this method is intended to search for all occurrances of a substring within a source string
' ... and return the number of matches found.
' ... pPositionArray is an array of longs that will be populated by char position numbers
' ... relating to the start of each substring within the source.
' ... However, if nothing is found, or there are invalid parameters, pPositionArray will
' ... be entirely unaffected e.g. not dimensioned in here at all.
' ... if an error is thrown or there is something invalid -1 will be returned.

Dim lngFound As Long
Dim lngLen As Long
Dim lngFLen As Long
Dim lngCounter As Long
Dim lngStart As Long
Dim sTheText As String

Dim lngCharLeft As Long
Dim lngCharRight As Long
'Dim bDoItAgain As Boolean
Dim bIsWord As Boolean
'Dim bControlChar As Boolean

'Const c_CharMax As Long = 65

    On Error GoTo ErrHan:
    
    FindAllMatches = -1  ' ... default error return, should be corrected later.
    
    lngLen = Len(m_TheString)
    lngFLen = Len(pFind)
    
    If lngLen = 0 Then Exit Function
    If lngFLen = 0 Then Exit Function
    If lngFLen > lngLen Then Exit Function
    If pStart > lngLen Then Exit Function
    
    sTheText = m_TheString
    
    lngStart = pStart
    If lngStart < 1 Then
        lngStart = 1
    End If
    
    If pCompare = vbTextCompare Then
        sTheText = LCase$(sTheText)
        pFind = LCase$(pFind)
    End If
    
ReTry:
    lngFound = 0
    If lngStart < lngLen Then
        lngFound = InStr(lngStart, sTheText, pFind)
    End If

    If lngFound > 0 Then
        If pWholeWordOnly Then
            lngCharLeft = 0: lngCharRight = 0
            If lngFound > 1 Then
                lngCharLeft = Asc(Mid$(sTheText, lngFound - 1, 1))
                If pbIsWordBreakChar(lngCharLeft) Then lngCharLeft = 0
'                bControlChar = lngCharLeft < 48 Or lngCharLeft > 57 And lngCharLeft < 65 Or lngCharLeft > 90 And lngCharLeft < 96 Or lngCharLeft > 122 And lngCharLeft < 127
'                If bControlChar Then lngCharLeft = 0
                
'                If lngCharLeft < 48 Then lngCharLeft = 0
            End If
            If lngLen > lngFound + lngFLen Then
                lngCharRight = Asc(Mid$(sTheText, lngFound + lngFLen))
                If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
'                bControlChar = lngCharRight < 48 Or lngCharRight > 57 And lngCharRight < 65 Or lngCharRight > 90 And lngCharRight < 96 Or lngCharRight > 122 And lngCharRight < 127
'                If bControlChar Then lngCharRight = 0
                
'                If lngCharRight < 48 Then lngCharRight = 0
            End If
            bIsWord = lngCharLeft + lngCharRight = 0 ' < c_CharMax
            If bIsWord = False Then
                lngStart = lngFound + 1
                GoTo ReTry:
            End If
        End If
        
        lngStart = lngFound + 1                     ' ... note: instr doesn't mind if start is > than source length.
        ReDim pPositionArray(0)
        pPositionArray(0) = lngFound
        lngCounter = 1
    End If
    
    Do While lngFound > 0
    
        lngFound = InStr(lngStart, sTheText, pFind)
        
        If lngFound > 0 Then
        
            If pWholeWordOnly Then
                lngCharLeft = 0: lngCharRight = 0
                If lngFound > 1 Then
                    lngCharLeft = Asc(Mid$(sTheText, lngFound - 1, 1))
                    If pbIsWordBreakChar(lngCharLeft) Then lngCharLeft = 0
'                    bControlChar = lngCharLeft < 48 Or lngCharLeft > 57 And lngCharLeft < 65 Or lngCharLeft > 90 And lngCharLeft < 96 Or lngCharLeft > 122 And lngCharLeft < 127
'                    If bControlChar Then lngCharLeft = 0
                    
'                    If lngCharLeft < 48 Then lngCharLeft = 0
                End If
                If lngLen > lngFound + lngFLen Then
                    lngCharRight = Asc(Mid$(sTheText, lngFound + lngFLen))
                    If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
'                    bControlChar = lngCharRight < 48 Or lngCharRight > 57 And lngCharRight < 65 Or lngCharRight > 90 And lngCharRight < 96 Or lngCharRight > 122 And lngCharRight < 127
'                    If bControlChar Then lngCharRight = 0
                    
'                    If lngCharRight < 48 Then lngCharRight = 0
                End If
                bIsWord = lngCharLeft + lngCharRight = 0 '< c_CharMax
            End If
        
            If pWholeWordOnly = False Or (pWholeWordOnly And bIsWord) Then
                ReDim Preserve pPositionArray(lngCounter)
                pPositionArray(lngCounter) = lngFound
                lngCounter = lngCounter + 1
            End If
            
            lngStart = lngFound + 1
        
        End If
    
    Loop

    FindAllMatches = lngCounter
    
ErrResume:

    lngCounter = 0&
    lngStart = 0&
    lngLen = 0&
    lngFLen = 0&
    sTheText = vbNullString
        
Exit Function

ErrHan:
    pErrMsg = Err.Description
    Err.Clear
    Resume ErrResume:
    
End Function

Private Function pbIsWordBreakChar(ByVal pKeyCode As Integer) As Boolean
Attribute pbIsWordBreakChar.VB_Description = "Tests if a keycode is a non-typeable character or a character that would break a word in VB."
    
    On Error GoTo ErrHan:
    
    pbIsWordBreakChar = pKeyCode < 48 Or pKeyCode > 57 And pKeyCode < 65 Or pKeyCode > 90 And pKeyCode < 96 Or pKeyCode > 122 And pKeyCode < 127

Exit Function

ErrHan:

    Debug.Print "StringWorker.pbIsWordBreakChar.Error: " & Err.Number & "; " & Err.Description

End Function

Function FindBackwards(sCheck As String, sMatch As String, Optional lStart As Long, Optional pCompare As VbCompareMethod = vbBinaryCompare) As Long

' ... my crap attempt to make an InstrRev function
' ... by way of delegating to InstrCharRev. when len(smatch) = 1, this method
' ... is able to provide the functionality required to search for a substring
' ... within a source string, backwards from a given position in the source.

Dim lngLoop As Long
Dim sBytes() As Byte
Dim sMatchBytes() As Byte
Dim lngLen As Long
Dim lngMLen As Long
Dim lngUBnd As Long
Dim lngChar As Long
Dim lngCounter As Long
Dim lngInLoop As Long
Dim lngFirst As Long
Dim lngMUBnd As Long
Dim lngMChar As Long
Dim bFound As Boolean

Dim k As Long
Dim m As Long

    lngMLen = Len(sMatch)
    If lngMLen = 0 Then Exit Function
    
    lngLen = Len(sCheck)
    If lngLen = 0 Then Exit Function
    
    If lngMLen > lngLen Then Exit Function
    
    
    If lngLen = 1 And lngMLen = 1 Then
        ' ... delegate to single match method, FindCharBackwards.
        FindBackwards = FindCharBackwards(sCheck, sMatch, lStart, pCompare)
        
    Else
    
        ' ... run through the bytes of the source backwards
        ' ... checking for the match, if match and match pos > start then
        ' ... carry on the search.
        ' ... first looking for the last char of the match in the source
        ' ... and then looping backwards through the match and the
        ' ... source to see if chars line up.
    
        sBytes = sCheck                         ' ... convert source to bytes.
        
        lngUBnd = UBound(sBytes)                ' ... get the upper limit of the source bytes.
        
        k = lngLen - 1                          ' ... set up current char pos in source.
        
        If lStart < 1 Then
            lStart = lngLen                     ' ... make sure start is ok for pos test later.
        End If
        
        sMatchBytes = sMatch                    ' ... convert the match to bytes.
        lngMUBnd = UBound(sMatchBytes)          ' ... the upper limit of the match bytes to loop.
        lngFirst = sMatchBytes(lngMUBnd - 1)    ' ... last char of match string.
                
        For lngLoop = lngUBnd - 1 To 0 Step -2
            
            lngChar = sBytes(lngLoop)
            
            If lngChar = lngFirst Then          ' ... when chars are same
                                                ' ... loop thru' the match bytes... backwards.
                lngCounter = 1                  ' ... for counting match length processed.
                
                For lngInLoop = lngMUBnd - 3 To 0 Step -2   ' ... start at penultimate match char.
                
                    m = lngLoop - (2 * lngCounter)          ' ... set up for reading next source char.
                                        
                    If m < 0 Then
                        ' ... try and make sure there's
                        ' ... a valid source char to test against.
                        Exit For
                    End If
                    
                    lngChar = sBytes(m)                     ' ... next source char.
                    
                    lngMChar = sMatchBytes(lngInLoop)       ' ... current match char.
                    
                    If lngChar = lngMChar Then
                    
                        If lngCounter + 1 = lngMLen Then    ' ... if match length achieved... final test.
                        
                            If k <= lStart Then             ' ... try to make sure pos is before user start param.
                                bFound = True
                            Else
                                k = k + lngCounter - 1      ' ... restore k (k is the current char pos in the source).
                            End If
                            
                            Exit For
                            
                        End If
                        
                    Else
                    
                        Exit For
                    
                    End If
                    
                    ' ... chars match so far so test next pair.
                    lngCounter = lngCounter + 1
                    
                    k = k - 1
                    
                Next lngInLoop
                
                If bFound = True Then
                    FindBackwards = k
                    Exit For
                End If
                
            End If
            
            k = k - 1
            
        Next lngLoop
    
        Erase sBytes
        Erase sMatchBytes
        
    End If

End Function

Function FindCharBackwards(ByVal StringCheck As String, ByVal StringMatch As String, Optional Start As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long

' ... Note:
' ...       Only tests for a single character for StringMatch.

Dim lngLoop As Long
Dim lngLen As Long
Dim lngDLen As Long
Dim lngLBnd As Long
Dim lngChar As Long
Dim lngCharPos As Long
Dim lngFindChar As Long
Dim lngStart As Long
Dim lngEnd As Long
Dim sBytes() As Byte
Dim lngReturn As Long
Dim bOK As Boolean


    lngReturn = 0

    lngLen = Len(StringCheck)
    lngDLen = Len(StringMatch)
    lngCharPos = lngLen
    
    If lngLen > 0 And lngDLen = 1 Then
        
        ' ... is it worth testing with instr first to see if
        ' ... we done actually got the search char or
        ' ... is this just bolox?
        
        bOK = InStr(1, StringCheck, StringMatch) > 0
        If bOK = False Then Exit Function
        
        ' ... handle text compare.
        If Compare = vbTextCompare Then
            StringCheck = LCase$(StringCheck)
            StringMatch = LCase$(StringMatch)
        End If
        
        ' ... handle single source char, v6 (code browser).
        If lngLen = 1 Then
            If StringCheck = StringMatch Then
                lngReturn = 1
            End If
            GoTo SkipSingleChar:
        End If
        
        ' ... set up the byte array.
        sBytes = StringCheck
        ' ... get the char no. of the char to find.
        lngFindChar = Asc(StringMatch)
        ' ... set up loop boudaries.
        lngLBnd = LBound(sBytes)
        lngEnd = lngLBnd
        
        ' ... set up the start position.
        If Start > -1 Then
            If Start > 0 Then
                If Start <= lngLen Then
                    lngCharPos = Start
                    lngStart = (Start - 1) * 2
                End If
            End If
        Else
            lngStart = UBound(sBytes) - 1
        End If
        
        ' ... run through the byte array backwards
        ' ... looking for the search char.
        If lngStart > 0 Then
            For lngLoop = lngStart To lngEnd Step -2
                lngChar = sBytes(lngLoop)
                If lngChar = lngFindChar Then
                    ' ... escape if this is the search char.
                    Exit For
                End If
                ' ... decrement char pos.
                lngCharPos = lngCharPos - 1
            Next lngLoop
            
            lngReturn = lngCharPos
            
        End If
        
        Erase sBytes
        
    End If

SkipSingleChar:
    
    FindCharBackwards = lngReturn
    
End Function

' Sub:             FromArray
' Description:     Uses VBA.Join to convert a string array into a single string.

Public Sub FromArray(ByRef pTheStringArray() As String, Optional ByRef pTheDelimiter As String = " ", Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute FromArray.VB_Description = "Uses VBA.Join to convert a string array into a single string."

'... Parameters.
'    R_A pTheStringArray: String     ' ... The string array to convert into the internal user string.
'    RO_ pTheDelimiter: String       ' ... The Delimiter to insert between string items.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.



Dim l As Long
Dim x As Long
Dim p As Long

Const c As Long = 2097152 '  1048576      ' ... err, pretty large string allocation length.

    On Error GoTo ErrStupid:

    m_TheString = Space$(c)
    l = 1
    p = LenB(pTheDelimiter)

    For m_Row = 0 To UBound(pTheStringArray)

        x = LenB(pTheStringArray(m_Row)) ' best so far.
        MidB$(m_TheString, l, x) = pTheStringArray(m_Row)
        l = l + x
        MidB$(m_TheString, l, p) = pTheDelimiter
        l = l + p

    Next m_Row

    l = l \ 2

    m_TheTextLength = l
    m_TheBufferLength = l
    m_TheString = Left$(m_TheString, l)
    
'    ' -------------------------------------------------------------------
'    Me = VBA.Join(pTheStringArray, pTheDelimiter)
'    ' -------------------------------------------------------------------
    pOK = True
    
Exit Sub
ErrStupid:
    If Err.Number = 5 Then
        m_TheString = m_TheString & Space$(c)
        Resume  ' ... could cause too much grief to bear.
    Else
        pErrMsg = Err.Description
        pOK = False
        Debug.Print "StringWorker.FromArray.Error: " & Err.Description
    End If

End Sub ' ... FromArray.

' Sub:             FromFile
' Description:     Reads the contents of a text file and makes it to the user string, overwriting anything else.

Public Sub FromFile(ByRef pTheFileName As String, Optional ByVal pDecrypt As Boolean = False, Optional ByVal pThePassword As String = vbNullString, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute FromFile.VB_Description = "Reads the contents of a text file and makes it to the user string, overwriting anything else."

'... Parameters.
'    R__ pTheFileName: String        ' ... The name of the file to read.
'    VO_ pDecrypt: Boolean           ' ... When True forces text to be decrypted.
'    VO_ pThePassword: String        ' ... The password to apply to the decryption.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim strFileText As String
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... read the text file.
    strFileText = pReadTextFile(pTheFileName, pOK, pErrMsg)
    ' -------------------------------------------------------------------
    
    If pOK = True Then
        ' -------------------------------------------------------------------
        If pDecrypt = True Then
            pDeMystify strFileText, pThePassword, pOK, pErrMsg
            If pOK = False Then
                Err.Raise vbObjectError + 1004, , "Did not Decrypt data:" & pErrMsg
            End If
        End If
        ' -------------------------------------------------------------------
        ' ... copy into buffer as user string.
        Me = strFileText
        
    Else
        ' -------------------------------------------------------------------
        ' ... update pErrMSg explaining file not read.
        Let pErrMsg = "FromFile: Did not Read the File: " & pTheFileName & vbCrLf & pErrMsg
    
    End If
    
ResErr:

    ' ... release resources...
    strFileText = vbNullString

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.FromFile", Err.Number, Err.Description
    Resume ResErr:
    
End Sub ' ... FromFile.

' Sub:             Insert
' Description:     Inserts a text string at the position specified.

Public Sub Insert(ByRef pTheString As String, ByVal pStart As Long, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Insert.VB_Description = "Inserts a text string at the position specified."

'... Parameters.
'    R__ pTheString: String          ' ... The string value to insert.
'    V__ pStart: Long                ' ... The position at which to insert the text in the user string.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

'   Notes:  From what I can see there are 3 things to do:
'           1.  Ensure enough length in the buffer and resize the internal user string text length.
'           2.  Shift trailing text to end of insert position within the user text string.
'           3.  Insert the new test at the position given.
'
'           This method allows the insert position to be greater than the current user string text length;
'           if this is the case then the gap inbetween the current end and the new start will be filled with space characters by default.
'
'           A start position < 1 will be interprested as 1.
'
'           No insert string, no processing, pErrMsg returns reason.
'
'           I have tried to optimise this for speed by using the mid$ function to perform the shift and insert; note preference to avoid api, strptr etc.

' -------------------------------------------------------------------
Dim lngShiftLength As Long
Dim lngInsertLength As Long
Dim lngBufferResize As Long
' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    lngInsertLength = Len(pTheString)
    
    If lngInsertLength > 0 Then ' ... only attempt insert if insert string not "".
    
        If pStart < 1 Then pStart = 1         ' ... Mid$ won't process zero start position.
        
        lngShiftLength = m_TheTextLength - pStart + 1  ' ... Shift length is length of text from insert position to end of user string before any processing.
        If lngShiftLength < 0 Then lngShiftLength = 0           ' ... Mid$ won't process a negative length.
            
        ' -------------------------------------------------------------------
        ' ... allow insert position to be beyond current length.
        lngBufferResize = lngInsertLength + IIf(pStart > m_TheTextLength, pStart - m_TheTextLength, 0)
        
        ' -------------------------------------------------------------------
        ' ... resize the buffer and user text length to include new text @ position given.
        pResizeBuffer lngBufferResize, pOK, pErrMsg
        
        ' -------------------------------------------------------------------
        If pOK = True Then
            
            ' -------------------------------------------------------------------
            ' ... reset return to false.
            pOK = False
            ' -------------------------------------------------------------------
            ' ... shift trailing text to end of insert position within the user text string.
            Mid$(m_TheString, pStart + lngInsertLength, lngShiftLength) = Mid$(m_TheString, pStart, lngShiftLength)
            
            ' -------------------------------------------------------------------
            ' ... insert the new test at the position given.
            Mid$(m_TheString, pStart, lngInsertLength) = pTheString
            
            ' -------------------------------------------------------------------
            pOK = True ' ... upate return to true.
            ' -------------------------------------------------------------------
            
        End If
        
    Else
    
        pErrMsg = "Insert String Length needs to be > 0."
    
    End If

Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "StringWorker.Insert", Err.Number, Err.Description

End Sub ' ... Insert.

' Property Get:    Length
' Type:            Long.
' Description:     Returns the length of the current user string (no. of characters).

Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the length of the current user string (no. of characters)."

    On Error GoTo ErrHan:

    Let Length = m_TheTextLength

Exit Property
ErrHan:

    Debug.Print "StringWorker.Length", Err.Number, Err.Description

End Property ' ... Length: Long.

' Function:        pCheckFileName
' Returns:         Boolean.
' Description:     Checks a file name for invalid characters.

Private Function pCheckFileName(ByVal pTheFileName As String, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute pCheckFileName.VB_Description = "Checks a file name for invalid characters."

'... Parameters.
'    R__ pTheFileName: String        ' ... The Full Path Name and Location of the FIle to check.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sFilePath As String
Dim sFileName As String
Dim sFileExt As String
Dim sTheChar As String
Dim sInvalidChars As String
Dim lngLoop As Long
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... Thanks to Fransesco Balena, VB2TheMax
    ' ... Split the full file name into path, name and extension.
    ' -------------------------------------------------------------------
    sFileName = pTheFileName
    ' -------------------------------------------------------------------
    For lngLoop = Len(sFileName) To 1 Step -1
        If Mid$(sFileName, lngLoop, 1) = "." Then
            sFileExt = Mid$(sFileName, lngLoop + 1)
            sFileName = Left$(sFileName, lngLoop - 1)
        ElseIf InStr(":\", Mid$(sFileName, lngLoop, 1)) Then
            sFilePath = Left$(sFileName, lngLoop)
            If Right$(sFilePath, 1) = "\" Then sFilePath = Left$(sFilePath, lngLoop - 1)
            sFileName = Mid$(sFileName, lngLoop + 1)
            Exit For
        End If
    Next lngLoop
    ' -------------------------------------------------------------------
        
    ' -------------------------------------------------------------------
    ' ... having extracted the file name make sure it is free of invalid characters.
    Let sInvalidChars = c_def_InvalidFileNameChars & Chr$(34) ' "`!$%^&*()-=+[]{}'#@~;,.<>/?\|" & Chr$(34)
    ' -------------------------------------------------------------------
    For lngLoop = 1 To Len(sFileName)
        Let sTheChar = Mid$(sFileName, lngLoop, 1)
        If InStr(sInvalidChars, sTheChar) Then
            ' -------------------------------------------------------------------
            ' ... easier to throw an error than anything else if invalid char. found.
            Err.Raise vbObjectError + 1001, , "The File Name [ " & sFileName & " ] has the following invalid character [ " & sTheChar & " ]."
            ' ... an alternative is to generate error msg and exit function here
            ' ... but that would be wasting the error handler.
        End If
    Next lngLoop
    ' -------------------------------------------------------------------
    ' ... could also check valid directory and prompt to create.
    
    ' -------------------------------------------------------------------
    
    Let bOK = True

ErrResume:

    Let pCheckFileName = bOK


Exit Function
ErrHan:

    Let pErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringWorker.pCheckFileName", Err.Number, Err.Description
    Resume ErrResume:

End Function ' ... pCheckFileName: Boolean.

' Sub:             pCryptScribe
' Description:     Manages the encryption / decryption of a string.

Private Sub pCryptScribe(ByRef pTheString As String, Optional ByVal pThePassword As String = vbNullString, Optional ByVal pCryptDirection As Long = 0, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pCryptScribe.VB_Description = "Manages the encryption / decryption of a string."

'... Parameters.
'    R__ pTheString: String          ' ... The string value to decrypt.
'    VO_ pThePassword: String        ' ... The (optional) password to apply to the decryption.
'    VO_ pCryptDirection: Long       ' ... 0 = Encrypt, 1 = Decrypt.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.
' -------------------------------------------------------------------

' Note:     Not the greatest encrytpion method, without a password the output could
'           more easily be deciphered and with a password one only has to try out
'           as many 'passwords' as it takes to crack.
'           Loop thru the bytes of the text, add / subtract index related number to / from
'           the current byte, if password then xor the current byte with the current password char code.
'           This method simply attempts to ensure that a character is ciphered differently according
'           to its position in the crypt string and that the repetition of this pattern proves
'           sufficiently confusing; use c_IndexSplit to effect change upon the length of the character
'           repetition pattern where more c_IndexSplit = more unique variations of a single character.
'
' -------------------------------------------------------------------
Dim bWord() As Byte
Dim bPWord() As Byte
Dim lngLoop As Long
Dim lngIndex As Long
Dim lngInc As Long
Dim lngChar As Long
Dim lngPWChar As Long
Dim lngPWordLen As Long
Dim lngPWordPos As Long
Dim strReturn As String
' -------------------------------------------------------------------
Const c_IncMax As Long = 3
Const c_IncMin As Long = 1
Const c_IndexSplit As Long = 5
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If Len(pTheString) = 0 Then
        Err.Raise vbObjectError + 1003, "Unable to process a zero length source string."
    End If
    
    bWord = pTheString
    bPWord = pThePassword
    
    lngPWordLen = Len(pThePassword) * 2
    lngPWordPos = 0
    
    strReturn = Space$(Len(pTheString))
    
    If pCryptDirection = 0 Then ' ... Encrypt.
    
        For lngLoop = 0 To UBound(bWord) Step 2
            
            lngInc = lngInc + 1
            If lngInc > c_IncMax Then
                lngInc = c_IncMin
            End If
            
            lngIndex = lngLoop / 2 + 1
            lngChar = bWord(lngLoop)
                        
            If lngIndex Mod c_IndexSplit = 0 Then
                lngChar = lngChar - lngInc
            Else
                lngChar = lngChar + lngInc
            End If
            
            If lngPWordLen > 0 Then ' ... include password?

                ' -------------------------------------------------------------------
                If lngPWordPos >= lngPWordLen Then
                    lngPWordPos = 0
                End If
    
                lngPWChar = bPWord(lngPWordPos)
    
                lngPWordPos = lngPWordPos + 2
    
                lngChar = lngChar Xor lngPWChar
                ' -------------------------------------------------------------------
            
            End If
            
            Mid$(strReturn, lngIndex, 1) = Chr$(lngChar)
            
        Next lngLoop
        
    ElseIf pCryptDirection = 1 Then  ' ... Decrypt.
    
        For lngLoop = 0 To UBound(bWord) Step 2
            
            lngInc = lngInc + 1
            If lngInc > c_IncMax Then
                lngInc = c_IncMin
            End If
            
            lngIndex = lngLoop / 2 + 1
            lngChar = bWord(lngLoop)
            
            If lngPWordLen > 0 Then ' ... password included?

                ' -------------------------------------------------------------------
                If lngPWordPos >= lngPWordLen Then
                    lngPWordPos = 0
                End If
    
                lngPWChar = bPWord(lngPWordPos)
    
                lngPWordPos = lngPWordPos + 2
                
                lngChar = lngChar Xor lngPWChar
                ' -------------------------------------------------------------------
            
            End If
            
            If lngIndex Mod c_IndexSplit = 0 Then
                lngChar = lngChar + lngInc
            Else
                lngChar = lngChar - lngInc
            End If
            
            Mid$(strReturn, lngIndex, 1) = Chr$(lngChar)
            
        Next lngLoop
    
    End If
    
    pTheString = strReturn

    pOK = True
    pErrMsg = vbNullString
ResumeErr:
    On Error GoTo 0
    Erase bWord
    Erase bPWord
    
Exit Sub
ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "StringWorker.pCryptScribe", Err.Number, Err.Description
    Resume ResumeErr:
    
End Sub ' ... pCryptScribe.

' Sub:             pDeMystify
' Description:     Converts encrypted text to a readable string via simple reverse encryption algorithm.

Private Sub pDeMystify(ByRef pTheString As String, Optional ByVal pThePassword As String = vbNullString, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pDeMystify.VB_Description = "Converts encrypted text to a readable string via simple reverse encryption algorithm."

'... Parameters.
'    R__ pTheString: String          ' ... The string value to decrypt.
'    VO_ pThePassword: String        ' ... The (optional) password to apply to the decryption.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.
    
    On Error GoTo ErrHan:
    
    pCryptScribe pTheString, pThePassword, 1, pOK, pErrMsg

Exit Sub
ErrHan:

    pOK = False
    pErrMsg = Err.Description
    Debug.Print "StringWorker.pDeMystify", Err.Number, Err.Description

End Sub ' ... pDeMystify.

' Function:        pFileExists
' Returns:         Boolean.
' Description:     Tests for file existence, returns True if found else False.

Private Function pFileExists(ByRef pTheFileName As String, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute pFileExists.VB_Description = "Tests for file existence, returns True if found else False."

'... Parameters.
'    R__ pTheFileName: String        ' ... The name of the file to test.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim bOK As Boolean ' ... Returns success or failure of this method.

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... check file exists.
    bOK = (Dir$(pTheFileName) <> "")
    ' -------------------------------------------------------------------
    If bOK = False Then
        ' -------------------------------------------------------------------
        Let pErrMsg = "The File [ " & pTheFileName & " ] could not be found."
        ' -------------------------------------------------------------------
    End If

ErrResume:

    Let pFileExists = bOK


Exit Function
ErrHan:

    Let pErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringWorker.pFileExists", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pFileExists: Boolean.

' Sub:             pMystify
' Description:     Converts text to an unreadable string via simple encryption algorithm.

Private Sub pMystify(ByRef pTheString As String, Optional ByVal pThePassword As String = vbNullString, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pMystify.VB_Description = "Converts text to an unreadable string via simple encryption algorithm."

'... Parameters.
'    R__ pTheString: String          ' ... The string value to encrypt.
'    VO_ pThePassword: String        ' ... The (optional) password to apply to the encryption.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:
    
    pCryptScribe pTheString, pThePassword, 0, pOK, pErrMsg
    
Exit Sub
ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "StringWorker.pMystify", Err.Number, Err.Description

End Sub ' ... pMystify.

' Function:        pReadTextFile
' Returns:         String.
' Description:     Reads text from a file into the current user string.

Private Function pReadTextFile(ByRef pTheFileName As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As String
Attribute pReadTextFile.VB_Description = "Reads text from a file into the current user string."

'... Parameters.
'    R__ pTheFileName: String        ' ... The name of the file to read.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim strReturn As String              ' ... a return value to this function.
Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
Dim bFileExists As Boolean
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    bFileExists = pFileExists(pTheFileName, pErrMsg)
    
    If bFileExists = True Then
        
        ' -------------------------------------------------------------------
        ' ... get the next free file number.
        iFileNumber = FreeFile()
        ' -------------------------------------------------------------------
        Open pTheFileName For Binary As #iFileNumber
        ' -------------------------------------------------------------------
        ' ... if execution flow got here, the file has been open without error.
        bFileIsOpen = True
        
        strReturn = Space$(LOF(iFileNumber))
        
        ' -------------------------------------------------------------------
        ' ... read the entire contents in one single operation.
        Get #iFileNumber, , strReturn
        ' -------------------------------------------------------------------
        ' ... Print Statement from pWriteTextFile adds vbCRLF to the text written.
        If Len(strReturn) > 2 Then
            If Right$(strReturn, 2) = vbCrLf Then
                strReturn = Left$(strReturn, Len(strReturn) - 2)
            End If
        End If
        ' -------------------------------------------------------------------
        pReadTextFile = strReturn    ' ... return the string from the text file.
        pOK = True                   ' ... return success.
        ' -------------------------------------------------------------------
    
    End If
    
ErrResume:

    On Error Resume Next
    ' -------------------------------------------------------------------
    ' ... try closing file silently to error.
    If bFileIsOpen = True Then
        Close #iFileNumber
        If Err.Number <> 0 Then
            Debug.Print "StringWorker.pReadTextFile (Close File)", Err.Number, Err.Description
            Err.Clear
        End If
    End If
    
Exit Function
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "StringWorker.pReadTextFile", Err.Number, Err.Description
    
    Err.Clear
    
    Resume ErrResume:
    
End Function ' ... pReadTextFile: String.

' Property Get:    PreAllocationSize
' Type:            Long.
' Description:     Sets / Returns the size (no. of Space Characters) to add to the buffer when it is not large enough to add the string passed.

Public Property Get PreAllocationSize() As Long
Attribute PreAllocationSize.VB_Description = "Sets / Returns the size (no. of Space Characters) to add to the buffer when it is not large enough to add the string passed."

    On Error GoTo ErrHan:

    Let PreAllocationSize = m_PreAllocationSize

Exit Property
ErrHan:

    Debug.Print "StringWorker.PreAllocationSize", Err.Number, Err.Description

End Property ' ... PreAllocationSize: Long.

' Property Let:    PreAllocationSize

Public Property Let PreAllocationSize(ByVal pNewValue As Long)

    On Error GoTo ErrHan:

    If pNewValue >= 1 Then
        m_PreAllocationSize = pNewValue
    Else
        Err.Raise vbObjectError + 1000, , "Pre-Allocation size may not be < 1"
    End If

Exit Property
ErrHan:

    Debug.Print "StringWorker.PreAllocationSize", Err.Number, Err.Description

End Property ' ... PreAllocationSize: Long.

' Sub:             PrefixWith
' Description:     Adds a string to the beginning of the user string.

Public Sub PrefixWith(ByRef pTheText As String, Optional ByRef pNewLineBefore As Boolean = False, Optional ByRef pNewLineAfter As Boolean = False, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute PrefixWith.VB_Description = "Adds a string to the beginning of the user string."

'... Parameters.
'    R__ pTheText: String            ' ... The string value to add at the front of the user string.
'    RO_ pNewLineBefore: Boolean     ' ... When True, adds a new line to the beginning of the source string.
'    RO_ pNewLineAfter: Boolean      ' ... When True, adds a new line to the end of the source string.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... just delegate to insert with zero pos for pre-pend.
    Me.Insert IIf(pNewLineBefore, vbNewLine, vbNullString) & pTheText & IIf(pNewLineAfter, vbNewLine, vbNullString), 0, pOK, pErrMsg
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.PrefixWith", Err.Number, Err.Description

End Sub ' ... PrefixWith.

' Sub:             pResizeBuffer
' Description:     Manages changing the size of the internal buffer.

Private Sub pResizeBuffer(Optional ByVal pLength As Long = -1, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pResizeBuffer.VB_Description = "Manages changing the size of the internal buffer."

'... Parameters.
'    VO_ pLength: Long               ' ... The number of characters to add or take away from the user text string.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    Select Case pLength
        
        Case 0, -1 ' ... compact.
            
            m_TheString = Left$(m_TheString, m_TheTextLength)
            
        Case Is > 0 ' ... grow.
            
            If (m_TheTextLength + pLength) > m_TheBufferLength Then
                
                If (pLength > m_PreAllocationSize) Then
                    
                    m_TheString = m_TheString & Space$(pLength)
                
                Else
                    
                    m_TheString = m_TheString & Space$(m_PreAllocationSize)
                
                End If
            
            End If
            ' -------------------------------------------------------------------
            ' ... update the user string length; contentious placement.
            m_TheTextLength = m_TheTextLength + pLength
            ' -------------------------------------------------------------------
        
        Case Else ' ... shrink; user given length.
            
            pLength = pLength * -1
            
            If m_TheTextLength - pLength > 0 Then
                Let m_TheTextLength = m_TheTextLength - pLength
                Let m_TheString = Left$(m_TheString, m_TheTextLength)
            End If
    
    End Select

    ' -------------------------------------------------------------------
    ' ... update the buffer length.
    Let m_TheBufferLength = Len(m_TheString)
    
    Let pOK = True

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.pResizeBuffer", Err.Number, Err.Description

End Sub ' ... pResizeBuffer.

Private Sub pReverseString(ByRef pTheString As String)

' ... Tries to Reverse the string passed.
' ... Using Byte Arrays and the StrConv function to limit use
' ... of strings or characters.
' ... doing this because am in VB5 and StrReverse not available.
' ... hoping this is optimised for faster performance (somehow)
' ... not returning a string is faster.

' NOTE: (fair warning).
'       This reverses the string passed so behaves quite
'       differently to VB6.VBA.StrReverse.

Dim lngLoop As Long
Dim lngLen As Long
Dim lngChar As Long
Dim lngBytePos As Long
Dim sTheString() As Byte
Dim sTheReturn() As Byte
Dim lngLBnd As Long
Dim lngUBnd As Long

    lngLen = Len(pTheString)
    
    If lngLen > 0 Then
        
        ' ... read the bytes of the input string.
        sTheString = pTheString
        ' ... set up the loop boundaries.
        lngLBnd = LBound(sTheString)
        lngUBnd = UBound(sTheString) - 1
        ' ... dimension the return array (half size of original)
        ReDim sTheReturn(lngUBnd \ 2)
        ' ... loop backwards through the byte array.
        For lngLoop = lngUBnd To lngLBnd Step -2
            ' ... read the next char number.
            lngChar = sTheString(lngLoop)
            ' ... write it to the next pos in the return array.
            sTheReturn(lngBytePos) = lngChar
            ' ... increment byte pos for next go.
            lngBytePos = lngBytePos + 1
                        
        Next lngLoop
        ' ... convert the return array back to a string and
        ' ... write it to the input/output string.
        pTheString = StrConv(sTheReturn, vbUnicode)
        
        ' ... memory clean up.
        Erase sTheString
        Erase sTheReturn
        
    End If

End Sub

' Sub:             pSplitB04
' Description:     Optimised Split method to return the user string as an array of string data.

Private Sub pSplitB04(ByRef pDestinationArray() As String, Optional ByRef pTheDelimiter As String = " ", Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pSplitB04.VB_Description = "Optimised Split method to return the user string as an array of string data."

'... Parameters.
'    R_A pDestinationArray: String   ' ... The array which shall be returned with the string data.
'    RO_ pTheDelimiter: String       ' ... The Delimiter upon which to Split the user string..
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' By Chris Lucas, cdl1051@earthlink.net, 20011208
' http://www.xbeat.net/vbspeed/c_Split.htm

    Dim c&, SLen&, DelLen&, tmp&, Results&()


    On Error GoTo ErrHan:

' -------------------------------------------------------------------
Dim Expression As String
    Expression = Me
' -------------------------------------------------------------------

    SLen = LenB(Expression) \ 2
    DelLen = LenB(pTheDelimiter) \ 2

    ' Bail if we were passed an empty delimiter or an empty expression
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve pDestinationArray(0 To 0)
        pDestinationArray(0) = Expression
        Exit Sub
    End If

    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    tmp = InStr(Expression, pTheDelimiter)

    Do While tmp
        Results(c) = tmp
        c = c + 1
        tmp = InStr(Results(c - 1) + 1, Expression, pTheDelimiter)
    Loop

    ' Size our return array
    ReDim Preserve pDestinationArray(0 To c)

    ' Populate the array
    If c = 0 Then
'         lazy man 's call
        pDestinationArray(0) = Expression
    Else
        ' typical call
        pDestinationArray(0) = Left$(Expression, Results(0) - 1)
        For c = 0 To c - 2
            pDestinationArray(c + 1) = Mid$(Expression, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
        Next c
        pDestinationArray(c + 1) = Right$(Expression, SLen - Results(c) - DelLen + 1)
    End If

    Let pErrMsg = vbNullString
    Let pOK = True


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.pSplitB04", Err.Number, Err.Description

End Sub ' ... pSplitB04.

' Sub:             pWriteTextFile
' Description:     Writes the current user string as text to a text file.

Private Sub pWriteTextFile(ByRef pTheText As String, ByRef pTheFileName As String, Optional ByRef pAppendMode As Boolean = False, Optional ByRef pConfirmOverWrite As Boolean = False, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pWriteTextFile.VB_Description = "Writes the current user string as text to a text file."

'... Parameters.
'    R__ pTheText: String            ' ... The string value to write to the file.
'    R__ pTheFileName: String        ' ... The name of the file to write.
'    RO_ pAppendMode: Boolean        ' ... When True, attempts to Append the Text to an existing File.
'    RO_ pConfirmOverWrite: Boolean  ' ... When True, asks for user permission to over write existing file data.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
Dim bFileExists As Boolean
Dim iAnswer As VbMsgBoxResult
Dim bFileNameOK As Boolean
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    Let bFileNameOK = pCheckFileName(pTheFileName, pErrMsg)
    ' -------------------------------------------------------------------
    
    If bFileNameOK = True Then
    
        If pAppendMode = False And pConfirmOverWrite = True Then
            
            Let bFileExists = pFileExists(pTheFileName, pErrMsg)
            
            If bFileExists = True Then
            
                Let iAnswer = MsgBox("The File [ " & pTheFileName & " ] Exists;" & vbCrLf & "Would you like to over-write this File?", vbYesNo + vbQuestion + vbDefaultButton1, "Over-Write File?")
                
                If iAnswer = vbYes Then
                    
                    Kill pTheFileName
                
                Else
                
                    Let pOK = False
                    Let pErrMsg = "User cancelled Save"
                    GoTo ErrResume:
                    
                End If
                
            End If
            
        End If
        
        ' -------------------------------------------------------------------
        ' ... get the next free file number.
        Let iFileNumber = FreeFile()
        
        If pAppendMode Then
            Open pTheFileName For Append As #iFileNumber
        Else
            Open pTheFileName For Output As #iFileNumber
        End If
        
        ' ... if execution flow got here, the file has been open correctly.
        Let bFileIsOpen = True
        
        ' ... print to the file in one single operation.
        Print #iFileNumber, pTheText;
        
        Let pOK = True
        
    Else
        
        Let pErrMsg = "The program is unable to write to file:" & vbNewLine & pErrMsg
        
    End If
    
ErrResume:

    On Error Resume Next
    ' -------------------------------------------------------------------
    ' ... try closing file silently to error.
    If bFileIsOpen = True Then
        Close #iFileNumber
        If Err.Number <> 0 Then
            Debug.Print "StringWorker.pWriteTextFile (Close File)", Err.Number, Err.Description
            Err.Clear
        End If
    End If

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.pWriteTextFile", Err.Number, Err.Description
    
    Err.Clear
    
    Resume ErrResume:
    
End Sub ' ... pWriteTextFile.

' Sub:             Replace
' Description:     Replaces a search string within the current user string with an alternative string.

Public Sub Replace(ByRef pSearchFor As String, ByRef pReplaceWith As String, Optional pStart As Long = 1, Optional ByRef pCountReplaced As Long = 0, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Replace.VB_Description = "Replaces a search string within the current user string with an alternative string."

'... Parameters.
'    R__ pSearchFor: String          ' ... The character/s to search for within the user text.
'    R__ pReplaceWith: String        ' ... The replacement string value.
'    RO_ pStart: Long                ' ... The position from which to begin the replacing.
'    RO_ pCountReplaced: Long        ' ... Returns the number ofsearch string replacements.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.


'   ... With Thanks.
'   ... submitted 18-Dec-2000 by Jost Schwider to VBSpeed.
'   ... http://www.xbeat.net/vbspeed/c_Replace.htm#Replace09

Dim TextLen As Long
Dim OldLen As Long
Dim NewLen As Long
Dim ReadPos As Long
Dim WritePos As Long
Dim CopyLen As Long
Dim Buffer As String
Dim BufferLen As Long
Dim BufferPosNew As Long
Dim BufferPosNext As Long
Dim Search As String

'   Note:   Using Byte versions of string functions so string must be pure AINSI.
'           This code will not migrate to .NET due to the Binary functions used
'           but then I'm sure .NET has got something way better than this to use.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Search = Me
    pCountReplaced = 0
  
    If pStart < 2 Then
        pStart = InStrB(Search, pSearchFor)
    Else
        pStart = InStrB(pStart + pStart - 1, Search, pSearchFor)
    End If
  
    If pStart Then
    
        ReadPos = 1
        WritePos = 1

        Buffer = Space$(Len(Search))

        OldLen = LenB(pSearchFor)
        NewLen = LenB(pReplaceWith)
        
        Select Case NewLen
        
            Case OldLen ' ... query and replace are same size, no effect on string length.
            
                Buffer = Search
                
                Do Until pStart = 0
                    MidB$(Buffer, pStart) = pReplaceWith
                    pStart = InStrB(pStart + OldLen, Buffer, pSearchFor)
                    pCountReplaced = pCountReplaced + 1
                Loop
            
            Case Is < OldLen ' ... replace is smaller than query string, string length will shrink.
            
                TextLen = LenB(Search)
                
                If NewLen Then
                
                    Do Until pStart = 0
                        CopyLen = pStart - ReadPos
                        If CopyLen Then
                            BufferPosNew = WritePos + CopyLen
                            MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                            MidB$(Buffer, BufferPosNew) = pReplaceWith
                            WritePos = BufferPosNew + NewLen
                        Else
                            MidB$(Buffer, WritePos) = pReplaceWith
                            WritePos = WritePos + NewLen
                        End If
                        ReadPos = pStart + OldLen
                        pStart = InStrB(ReadPos, Search, pSearchFor)
                        pCountReplaced = pCountReplaced + 1
                    Loop
                
                Else    ' ... replace is empty.
                
                    Do Until pStart = 0
                        CopyLen = pStart - ReadPos
                        If CopyLen Then
                            MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                            WritePos = WritePos + CopyLen
                        End If
                        ReadPos = pStart + OldLen
                        pStart = InStrB(ReadPos, Search, pSearchFor)
                        pCountReplaced = pCountReplaced + 1
                    Loop
                
                End If
      
                If ReadPos > TextLen Then
                    Buffer = LeftB$(Buffer, WritePos - 1)
                Else
                    MidB$(Buffer, WritePos) = MidB$(Search, ReadPos)
                    Buffer = LeftB$(Buffer, WritePos + TextLen - ReadPos)
                End If
                
            Case Else   ' ... replace is larger than query, string length will grow.
                
                TextLen = LenB(Search)
                
                BufferPosNew = TextLen + NewLen
                
                If BufferPosNew > BufferLen Then
                    Buffer = Space$(BufferPosNew)
                    BufferLen = LenB(Buffer)
                End If
                      
                Do Until pStart = 0
                    
                    CopyLen = pStart - ReadPos
                    
                    If CopyLen Then
                        
                        BufferPosNew = WritePos + CopyLen
                        BufferPosNext = BufferPosNew + NewLen
                        
                        If BufferPosNext > BufferLen Then
                            ' ... grow the buffer.
                            Buffer = Buffer & Space$(BufferPosNext)
                            BufferLen = LenB(Buffer)
                        End If
                        
                        MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                        MidB$(Buffer, BufferPosNew) = pReplaceWith
                    
                    Else
                        
                        BufferPosNext = WritePos + NewLen
                        
                        If BufferPosNext > BufferLen Then
                            ' ... grow the buffer.
                            Buffer = Buffer & Space$(BufferPosNext)
                            BufferLen = LenB(Buffer)
                        End If
                        
                        MidB$(Buffer, WritePos) = pReplaceWith
                    
                    End If
                    
                    WritePos = BufferPosNext
                    ReadPos = pStart + OldLen
                    pStart = InStrB(ReadPos, Search, pSearchFor)
                    pCountReplaced = pCountReplaced + 1
                
                Loop
      
                If ReadPos > TextLen Then
                    Buffer = LeftB$(Buffer, WritePos - 1)
                Else
                    BufferPosNext = WritePos + TextLen - ReadPos
                    If BufferPosNext < BufferLen Then
                        MidB$(Buffer, WritePos) = MidB$(Search, ReadPos)
                        Buffer = LeftB$(Buffer, BufferPosNext)
                    Else
                        Buffer = LeftB$(Buffer, WritePos - 1) & MidB$(Search, ReadPos)
                    End If
                End If
            
        End Select
        
        Let TheString = Buffer
        
    End If
    
    pOK = True ' ... executed, no errors, doesn't mean anything was replaced though.
    
ResumeError:

Exit Sub

ErrHan:
    
    pOK = False
    Let pErrMsg = Err.Description
    
    Debug.Print "StringWorker.Replace.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:

End Sub ' ... Replace.

Public Sub Reverse(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Reverse.VB_Description = "Reverses the current user string."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    If m_TheTextLength > 0 Then
        
'''        ' ... no optimisation; vba lib.
'''        Mid$(m_TheString, 1, m_TheTextLength) = VBA.StrReverse(Mid$(m_TheString, 1, m_TheTextLength))
        
        m_TheString = Left$(m_TheString, m_TheTextLength)
        pReverseString m_TheString
        
    End If

    Let pOK = True


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.Reverse", Err.Number, Err.Description

End Sub ' ... Reverse.

Public Function StrReverse(ByVal pTheString As String) As String

' ... VB5 substitute for VB6.VBA.StrReverse

Dim sTmp As String

    sTmp = pTheString
    pReverseString sTmp
    
    StrReverse = sTmp
    
    sTmp = vbNullString
    
End Function
' Sub:             Reverse
' Description:     Reverses the current user string.

' Property Get:    TheString: Default.
' Type:            String.
' Description:     Sets / Returns the user string.

Public Property Get TheString() As String
Attribute TheString.VB_Description = "Sets / Returns the user string."
Attribute TheString.VB_UserMemId = 0

    On Error GoTo ErrHan:

    TheString = Left$(m_TheString, m_TheTextLength)

Exit Property
ErrHan:

    Debug.Print "StringWorker.TheString", Err.Number, Err.Description

End Property ' ... TheString: String.

' Property Let:    TheString

Public Property Let TheString(ByVal pNewValue As String)

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    Let m_TheString = pNewValue
    ' -------------------------------------------------------------------
    Let m_TheTextLength = Len(m_TheString)
    Let m_TheBufferLength = m_TheTextLength
    ' -------------------------------------------------------------------

Exit Property
ErrHan:

    ' ... error 14 is ' out of string space '.
    Debug.Print "StringWorker.TheString", Err.Number, Err.Description

End Property ' ... TheString: String.

' Sub:             ToArray
' Description:     Convert the current user string into a string array by a given delimiter.

Public Sub ToArray(ByRef pDestinationArray() As String, Optional ByRef pTheDelimiter As String = """", Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ToArray.VB_Description = "Convert the current user string into a string array by a given delimiter."

'... Parameters.
'    R_A pDestinationArray: String   ' ... The array into which the user string will be split.
'    RO_ pTheDelimiter: String       ' ... The Delimiter upon which to Split the user string.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    pSplitB04 pDestinationArray, pTheDelimiter, pOK, pErrMsg
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.ToArray", Err.Number, Err.Description

End Sub ' ... ToArray.

' Sub:             ToFile
' Description:     Writes the contents of the current user string within the buffer to disk.

Public Sub ToFile(ByRef pTheFileName As String, Optional ByVal pEncrypt As Boolean = False, Optional ByVal pThePassword As String = vbNullString, Optional ByVal pConfirmOverWrite As Boolean = False, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ToFile.VB_Description = "Writes the current user string as text to a text file."

'... Parameters.
'    R__ pTheFileName: String        ' ... The name of the file to write.
'    VO_ pEncrypt: Boolean           ' ... When True forces text to be encrypted before saving to file.
'    VO_ pThePassword: String        ' ... The password to apply to the encryption (mystification).
'    VO_ pConfirmOverWrite: Boolean  ' ... When True, asks for user permission to over write existing file data.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim sTheFileText As String
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    sTheFileText = Me
    ' -------------------------------------------------------------------
    If pEncrypt = True Then
        pMystify sTheFileText, pThePassword, pOK, pErrMsg
        If pOK = False Then
            Err.Raise vbObjectError + 1004, , "Did not Encrypt or Save data:" & pErrMsg
        End If
    End If
    ' ... Note: append mode has been hidden from the interface; it is not appropriate
    ' ...       to encrypted text; the False parameter below forces overwrite rather than append.
    ' -------------------------------------------------------------------
    pWriteTextFile sTheFileText, pTheFileName, False, pConfirmOverWrite, pOK, pErrMsg
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.ToFile", Err.Number, Err.Description

End Sub ' ... ToFile.

' Sub:             ToLowerCase
' Description:     Converts the user text string to lower case characters using LCase$.

Public Sub ToLowerCase(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ToLowerCase.VB_Description = "Converts the user text string to lower case characters using LCase$."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    If Len(m_TheString) > 0 And m_TheTextLength > 0 Then
        
        Mid$(m_TheString, 1, m_TheTextLength) = LCase$(Mid$(m_TheString, 1, m_TheTextLength))
    
        Let pOK = True
        
    Else
        
        Let pErrMsg = "There was nothing to convert to lower case."
        
    End If


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.ToLowerCase", Err.Number, Err.Description

End Sub ' ... ToLowerCase.

' Sub:             ToProperCase
' Description:     Converts the user text string to proper case characters using StrConv with vbProperCase.

Public Sub ToProperCase(Optional ByRef pKeepCapitals As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ToProperCase.VB_Description = "Converts the user text string to proper case characters using StrConv with vbProperCase."

'... Parameters.
'    RO_ pKeepCapitals: Boolean      ' ... Restore existing Capitals post StrConv when True.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim sProper As String
Dim sBuffer As String
Dim lngTmpChar As Long
Dim sTmpChar As String
' -------------------------------------------------------------------

    On Error GoTo ErrHan:

' -------------------------------------------------------------------
' Note  :
'           this bit with thanks to Fransesco Balena, VB2TheMax.
'           http://www.devx.com/vb2themax/Tip/18549
' -------------------------------------------------------------------

    If m_TheTextLength > 0 Then
        
        sBuffer = Me
        sProper = StrConv(sBuffer, vbProperCase)
        
        If pKeepCapitals Then
        
            
            For m_Row = 1 To Len(sBuffer)
                
                lngTmpChar = Asc(Mid$(sBuffer, m_Row, 1))
                
                Select Case lngTmpChar
                    Case 65 To 90
                        sTmpChar = Mid$(sBuffer, m_Row, 1)
                        Mid$(sProper, m_Row, 1) = sTmpChar
                End Select
                
            Next m_Row
        
        End If
        
        Mid$(m_TheString, 1, Len(sBuffer)) = sProper
    
    End If
        
    pOK = True


Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "StringWorker.ToProperCase", Err.Number, Err.Description

End Sub ' ... ToProperCase.

' Sub:             ToUpperCase
' Description:     Converts the user text string to upper case characters using UCase$.

Public Sub ToUpperCase(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ToUpperCase.VB_Description = "Converts the user text string to upper case characters using UCase$."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:

    If Len(m_TheString) > 0 And m_TheTextLength > 0 Then
        
        Mid$(m_TheString, 1, m_TheTextLength) = UCase$(Mid$(m_TheString, 1, m_TheTextLength))
    
        Let pOK = True
        
    Else
        
        Let pErrMsg = "There was nothing to convert to upper case."
        
    End If


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.ToUpperCase", Err.Number, Err.Description

End Sub ' ... ToUpperCase.

' Sub:             Trim
' Description:     Trim trailing and leading Space characters from the current user string.

Public Sub Trim(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Trim.VB_Description = "Trim trailing and leading Space characters fgrom the current user string."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim sTmp As String
Dim lngTmp As Long

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    If m_TheTextLength > 0 Then
        ' -------------------------------------------------------------------
        sTmp = Me
        ' -------------------------------------------------------------------
        If Left$(sTmp, 1) = " " Then
            sTmp = LTrim$(sTmp)
        End If
        ' -------------------------------------------------------------------
        If Right$(sTmp, 1) = " " Then
            sTmp = RTrim$(sTmp)
        End If
        ' -------------------------------------------------------------------
        lngTmp = m_TheTextLength - Len(sTmp)
        ' -------------------------------------------------------------------
        If lngTmp <= m_TheTextLength Then
            Mid$(m_TheString, 1, m_TheTextLength - lngTmp) = sTmp
            m_TheTextLength = m_TheTextLength - lngTmp
        End If
        ' -------------------------------------------------------------------
    End If
    ' -------------------------------------------------------------------
    Let pOK = True
    ' -------------------------------------------------------------------
    
Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.Trim", Err.Number, Err.Description

End Sub ' ... Trim.

' Sub:             TrimLeft
' Description:     Trim leading Space characters from the current user string.

Public Sub TrimLeft(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute TrimLeft.VB_Description = "Trim leading Space characters from the current user string."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim sTmp As String

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    If m_TheTextLength > 0 Then
        ' -------------------------------------------------------------------
        sTmp = Me
        ' -------------------------------------------------------------------
        If Left$(sTmp, 1) = " " Then
            sTmp = LTrim$(sTmp)
            m_TheTextLength = Len(sTmp)
            Mid$(m_TheString, 1, m_TheTextLength) = sTmp
        End If
        ' -------------------------------------------------------------------
        
    End If
    ' -------------------------------------------------------------------
    Let pOK = True
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.TrimLeft", Err.Number, Err.Description

End Sub ' ... TrimLeft.

' Sub:             TrimRight
' Description:     Trim trailing Space characters from the current user string.

Public Sub TrimRight(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute TrimRight.VB_Description = "Trim trailing Space characters from the current user string."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim sTmp As String

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    If m_TheTextLength > 0 Then
        ' -------------------------------------------------------------------
        sTmp = Me
        ' -------------------------------------------------------------------
        If Right$(sTmp, 1) = " " Then
            sTmp = RTrim$(sTmp)
            m_TheTextLength = Len(sTmp)
            Mid$(m_TheString, 1, m_TheTextLength) = sTmp
        End If
        ' -------------------------------------------------------------------
    End If
    ' -------------------------------------------------------------------
    Let pOK = True
    ' -------------------------------------------------------------------

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.TrimRight", Err.Number, Err.Description

End Sub ' ... TrimRight.

' Function:        WordCount
' Returns:         Long.
' Description:     Counts and returns the number of words within the current user string.

Public Function WordCount(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute WordCount.VB_Description = "Counts and returns the number of words within the current user string."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim lngReturn As Long ' ... a return value to this function.
Dim Chars() As Byte
Dim iChar As Integer
Dim lngLoop As Long
Dim lngUbound As Long
Dim lngMax As Long
' -------------------------------------------------------------------
Const c_CharLow As Integer = 32
Const c_charHigh As Integer = 33
' -------------------------------------------------------------------

'   code thanks to:
'   http://www.xbeat.net/vbspeed/c_WordCount.htm#WordCount07
'   it appears that api/typelib is essential to getting faster results.

' Note:
'   I was following other examples at the time of writing which led to the < 33 char no. bit
'   this isn't quite the whole story because it ignores many other word breaking chars.
'   see pbIsWordBreakChar for a later attempt at identifying such chars.
'   In this respect, this method is bound to give stray results.

    On Error GoTo ErrHan:
    
    If m_TheTextLength > 0 Then
    
        Chars = Me
        
        lngUbound = UBound(Chars)
        lngMax = lngUbound + 1
        
        For lngLoop = 0 To lngUbound Step 2
            
            iChar = Chars(lngLoop)
            
            If iChar > c_CharLow Then
                
                Let lngReturn = lngReturn + 1
                
                Do
                    lngLoop = lngLoop + 2
                    
                    If lngMax > lngLoop Then
                        iChar = Chars(lngLoop)
                    Else
                        Exit Do
                    End If
                    
                Loop Until iChar < c_charHigh
                
            End If
            
        Next lngLoop
        
    End If
    
    WordCount = lngReturn
    
    Erase Chars
    
    pOK = True

Exit Function
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "StringWorker.WordCount", Err.Number, Err.Description

End Function ' ... WordCount: Long.
