VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "a class to help read text from a file or a string and then extract single lines or a continuous section of lines."
' v: Best One Yet 16 Jan 2012

' textreader
' a class to help read text from a file or a string
' and then extract single lines or a continuous section of lines
' and search for all instances of a substring and return the
' line/s on which the substring is found

' this represents an attempt to provide an alternative solution to splitting a source
' string into an array of lines when wanting to read the source as individual lines

' Note:
' this class handles text with lines delimited by Chr$(13) & Chr$(10) [vbCrLf] or Chr$(10) [vbLf]
' text lines delimited by any other character/s will not process as planned

' Note:
' GetAllMatches returns a string of results, all the lines on which the Search string has been found
' the field delimiter is Chr$(1) and the Line/record delimiter is vbCrLf
' each record is formatted with the following fields:

' Name                                          Type            Index
' -------------------------------------------------------------------
' Find Index (0 based)                          : Long          : 0
' Tag / Source name                             : String        : 1
' File Name (where applicable)                  : String        : 2
' Line Number                                   : Long          : 3
' Character Position                            : Long          : 4
' Length of Search String                       : Long          : 5
' The Line where the Search String exists       : String        : 6
' -------------------------------------------------------------------

' requires:
'   SBuilder to help build search results in GetAllMatches
'            otherwise, not required
' -------------------------------------------------------------------

Option Explicit

Private mFN As String   ' file name
Private mTxt As String  ' string version of text
Private mDel As String  ' text line delimiter
Private mTag As String  ' user defined string used in getallmatches to identify source name

Private mB() As Byte    ' byte version of text

Private miLP() As Long  ' line char positions

Private miTL As Long    ' text length (chars)
Private miDel As Long   ' line delimiter length
Private miLC As Long    ' line count

'Private mdByteSize As Double

Property Get TextByteSize() As Double
    TextByteSize = UBound(mB) - 1
End Property

Property Get FileName() As String
Attribute FileName.VB_Description = "Returns the name of the File used to provide the source string but only if a file was used to begin with."
' return the file name used to read the source string (if any)
    FileName = mFN
End Property

Property Get TheLineDelimiter() As String
Attribute TheLineDelimiter.VB_Description = "Returns the Line Delimiter found in the source string (only reads vbCrLf and vbLf as line separators)."
' return the line delimiter discovered in the text
    TheLineDelimiter = mDel
End Property

Property Get TheText() As String
Attribute TheText.VB_Description = "Returns the source string in its entirety."
' return the whole source string
    TheText = mTxt
End Property

Property Get Tag() As String
    Tag = mTag
End Property

Property Let Tag(pNewTag As String)
    mTag = pNewTag
End Property

Function GetAllMatches(ByVal pFind As String, _
              Optional ByVal pStart As Long = 0, _
              Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
              Optional ByVal pWholeWordOnly As Boolean = False, _
              Optional ByRef pCount As Long = 0, _
              Optional ByRef pUniqueLineCount As Long = 0, _
              Optional ByRef pErrMsg As String = vbNullString) As String
              
' return a string containing found lines and search specific data

' context
' we have a source string and want to find all instances of a search string within it
' we want the results output to a line delimited string describing and showing the lines
' on which the search string has been found.
' we want this to be able to load results in a list view and use this to navigate the
' source text per found line.

' Note:
' Using a SBuilder Class to generate output because it is much faster
' than building the output with a local string or the function value
' Another optimisation was to save last line and line number to avoid
' re-reading the same line.

' -------------------------------------------------------------------
' BUG:  RESULTS LINES DELIMITERS CAN BE ONE SHORT
'       THIS WILL MESS UP SPLITTING THE RESULTS LINES INTO CORRECT NO. OF LINES
' -------------------------------------------------------------------

Dim iPos()      As Long         ' found char positions
Dim iLines()    As Long         ' lines where found
Dim i           As Long         ' loop
Dim iFLen       As Long         ' length of serch string
Dim sDL         As String       ' result field delimiter
Dim xString     As SBuilder
Dim sLastLine   As String
Dim iLastLine   As Long

    iFLen = Len(pFind)
    sDL = Chr$(1)
    
    pCount = FindAllMatches(pFind, iPos, pStart, pCompare, pWholeWordOnly, pErrMsg)
    
    If pCount > 0 Then
        Set xString = New SBuilder
        DeriveLineNumbers iPos, iLines
        iLastLine = -1
        For i = 0 To UBound(iLines)
            If iLastLine < iLines(i) Then
                iLastLine = iLines(i)
                sLastLine = Line(iLastLine)
                pUniqueLineCount = pUniqueLineCount + 1
            End If
            xString.AppendAsLine i & sDL & mTag & sDL & mFN & sDL & iLastLine & sDL & iPos(i) & sDL & iFLen & sDL & sLastLine
        Next i
        GetAllMatches = Left$(xString.TheString, xString.Length - 2)
        Set xString = Nothing
    End If

    Erase iPos
    Erase iLines
    
End Function

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

' return position of substring in main string
' wrapper for VBA.Instr
' note: not as extensive with its search criteria as Find/Get-AllMatches

' context
' we want to find a string in the main string and
' only want the first/next position at a time


' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    lngCharLength = Len(pTheString)
    
    ' -------------------------------------------------------------------
    ' ... test there is something to search in and for and that the search and search + start are not longer than the source.
    If miTL > 0 And lngCharLength > 0 Then 'And ((lngCharLength <= miTL) And (lngCharLength + pStart <= miTL)) Then
        
        ' -------------------------------------------------------------------
        ' ... ensure start position is at least 1.
        If pStart < 1 Then Let pStart = 1
        
        ' -------------------------------------------------------------------
        ' ... check that we are within limits.
        If (lngCharLength <= miTL) And (lngCharLength + pStart <= miTL) Then
            
            sSource = mTxt    ' ... user string, ensures we do not need to test find pos in legal length.
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
                        End If
                        If miTL > lFoundPos + lngCharLength Then
                            lngCharRight = Asc(Mid$(sSource, lFoundPos + lngCharLength, 1))
                            If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
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
    
ResumeErr:
    sSource = vbNullString
    sQuery = vbNullString
    lFoundPos = 0
    lngStart = 0
    
Exit Function
ErrHan:
    Find = -1
    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringWorker.Find", Err.Number, Err.Description
    GoTo ResumeErr:
    
End Function ' ... Find: Long.

Function FindAllMatches(ByVal pFind As String, _
                        ByRef pPosArray() As Long, _
               Optional ByVal pStart As Long = 0, _
               Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
               Optional ByVal pWholeWordOnly As Boolean = False, _
               Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute FindAllMatches.VB_Description = "Returns the number of substring matches found in source from start pos onwards. pPosArray returns an array of found positions."

' return the number of matches of search string within source string
' and provides an array of positions at which the search results were found

' context
' we have a source string and want to find all instances of a search string within it
' we may or may not want to access each line containing a find result afterwards
' if we do want this then we would use the pPosArray with the DeriveLineNumbers method
' to find the line numbers array and then loop thru' its elements using the Line method
' to return each line of text containing the find

Dim sSource         As String
Dim tmpPos()        As Long
Dim iFindLen        As Long
Dim iCharLeft       As Long
Dim iCharRight      As Long
Dim iFind           As Long
Dim iCount          As Long
Dim jCount          As Long
Dim i               As Long

    On Error GoTo ErrHan:
    ' setup a default return --------------------------------------------
    FindAllMatches = -1
    ' -------------------------------------------------------------------
    iFindLen = Len(pFind)
    If pStart < 1 Then pStart = 1
    ' basic validation --------------------------------------------------
    If miTL = 0 Or iFindLen = 0 Then Exit Function
    If iFindLen > miTL Or pStart > miTL Then Exit Function
    ' search for matches ------------------------------------------------
    ReDim tmpPos(miTL)
    If pCompare = vbTextCompare Then    ' attempt optimisation, avoid unnecessary stuff when possible
        sSource = LCase$(mTxt)
        pFind = LCase$(pFind)
        iFind = InStr(pStart, sSource, pFind)
        Do While iFind
            tmpPos(iCount) = iFind
            iCount = iCount + 1
            iFind = InStr(iFind + 1, sSource, pFind)
        Loop
    Else
        iFind = InStr(pStart, mTxt, pFind)
        Do While iFind
            tmpPos(iCount) = iFind
            iCount = iCount + 1
            iFind = InStr(iFind + 1, mTxt, pFind)
        Loop
    End If
    ' original find unhindered by whole word only testing ---------------
    If iCount Then
        ReDim Preserve tmpPos(iCount - 1)
        ReDim pPosArray(iCount - 1)
        ' filter non-wholewords from results --------------------------------
        If pWholeWordOnly Then
            For i = 0 To iCount - 1
                iCharLeft = 0: iCharRight = 0
                iFind = tmpPos(i)
                If iFind > 1 Then
                    iCharLeft = Asc(Mid$(mTxt, iFind - 1, 1))
                    If pbIsWordBreakChar(iCharLeft) Then iCharLeft = 0
                End If
                If miTL > iFind + iFindLen Then
                    iCharRight = Asc(Mid$(mTxt, iFind + iFindLen))
                    If pbIsWordBreakChar(iCharRight) Then iCharRight = 0
                End If
                If iCharLeft + iCharRight = 0 Then
                    pPosArray(jCount) = iFind
                    jCount = jCount + 1
                End If
            Next i
            ReDim Preserve pPosArray(jCount - 1)
            iCount = jCount
        Else
            ' opportunity for API:RtlMoveMemory
            For i = 0 To iCount - 1
                pPosArray(i) = tmpPos(i)
            Next i
        End If
    Else
        ReDim pPosArray(0)
        pPosArray(0) = -1
    End If
    ' -------------------------------------------------------------------
    FindAllMatches = iCount
    
ErrResume:
    
    sSource = vbNullString
    Erase tmpPos
    
Exit Function

ErrHan:
    pErrMsg = Err.Description
    Err.Clear
    Resume ErrResume:
    
End Function

Function LineNoFromChar(pCharPos As Long) As Long
Attribute LineNoFromChar.VB_Description = "Returns the Line Number that a character position belongs to."
' returns a line number for a given character position

' context
' we have a character position and want to know on which line it belongs

Dim l As Long
Dim i As Long

    On Error GoTo ErrHan:
        
    If pCharPos Then
        l = UBound(miLP)
        If pCharPos <= miLP(l) Then
            For i = 1 To l
                If pCharPos < miLP(i) Then
                    LineNoFromChar = i
                    Exit For
                End If
            Next i
        End If
    End If
    
Exit Function
ErrHan:
    Debug.Print "TextReader.LineNoFromChar.Error: " & Err.Number & "; " & Err.Description

End Function

Function LineFromChar(pCharPos As Long) As String
Attribute LineFromChar.VB_Description = "Returns a Line from the source string given a character position."
' returns a line from a given character position

' context
' we have a character position and want to display the line on which it belongs

Dim l As Long
Dim i As Long

    On Error GoTo ErrHan:
            
    If pCharPos Then
        l = UBound(miLP)
        If pCharPos <= miLP(l) Then
            For i = 1 To l
                If pCharPos < miLP(i) Then
                    ' could have used LineNoFromChar to get here
                    LineFromChar = Line(i)
                    Exit For
                End If
            Next i
        End If
    End If
    
Exit Function
ErrHan:
    Debug.Print "TextReader.LineFromChar.Error: " & Err.Number & "; " & Err.Description

End Function

Function CountLines() As Long
Attribute CountLines.VB_Description = "Returns the number of Lines currently read."
' returns the number of lines in the source string
    CountLines = miLC
End Function

Function CountEmptyLines() As Long
' TODO
End Function

Function Length() As Long
Attribute Length.VB_Description = "Returns the number of characters in the source string."
' returns the number of characters in the source string
    Length = miTL
End Function

Function IsMultiByte() As Boolean
Attribute IsMultiByte.VB_Description = "Returns a Boolean describing whether there are Multi-Byte Characters in the source string."
' returns boolean describing whether the source string contains multi-byte characters

' context
' we have a source string and just want to know if it contains any multi-byte characters
' so that we can process the string correctly w/r displaying mbcs/unicode chars

Dim i As Long
Dim c As Long
Dim l As Long

    On Error GoTo ErrHan:
        
    l = UBound(mB)
    For i = 1 To l Step 2   ' read odd byte value, scbs should be 0
        c = mB(i)
        If c <> 0 Then
            IsMultiByte = True
            Exit For
        End If
    Next i

Exit Function
ErrHan:

End Function

Function Line(ByVal pLineNo As Long) As String
Attribute Line.VB_Description = "Returns a single line from the source string."
Attribute Line.VB_UserMemId = 0
' returns a line from a line number

' context
' we have a source string and a line number and want to access this line alone
    
    Line = Lines(pLineNo, pLineNo)
    
End Function

Function Lines(Optional ByVal pLineFrom As Long = 0, _
               Optional ByVal pLineTo As Long = 0) As String
Attribute Lines.VB_Description = "Returns a continuous section of lines from the source string."
' returns a section of lines from a start line upto, and including, an end line

' context
' we have a source string and line from and line to numbers and want to access the lines
' between and including the line numbers given

Dim iStart  As Long
Dim iLen    As Long
Dim iDel    As Long

    On Error GoTo ErrHan:
    
    If miLC Then
        
        If pLineFrom <= 0 Then pLineFrom = 1
        ' -------------------------------------------------------------------
        If pLineTo <= 0 Then pLineTo = miLC
        ' -------------------------------------------------------------------
        If pLineTo > miLC Then pLineTo = miLC
        ' -------------------------------------------------------------------
        If pLineTo < miLC Then iDel = miDel
        ' -------------------------------------------------------------------
        If pLineTo >= pLineFrom Then
            
            iStart = miLP(pLineFrom - 1)
            
            If pLineTo < miLC Then
                iLen = miLP(pLineTo) - iStart - iDel
                Lines = Mid$(mTxt, iStart, iLen)
            Else
                Lines = Mid$(mTxt, iStart)
            End If
            
        End If
        
    End If

Exit Function
ErrHan:
    Debug.Print "TextReader.Lines.Error: " & Err.Number & "; " & Err.Description

End Function

Sub DeriveLineNumbers(ByRef pPositions() As Long, _
                      ByRef pLineNumbers() As Long)

' this tries to derive an array of line numbers from
' an array of character positions into pLineNumbers

' note:
' values in pPositions elements are presumed to increase from first to last

' context
' we've just run a find all matches and we want to
' extract the lines on which the search string was found
' we've grabbed the pPosArray from the find all matches
' and bung it here to get the line numbers involved

Dim lngLUBnd        As Long     ' ubound for internal lines array
Dim lngPUBnd        As Long     ' ubound for positions array [pPositions]
Dim lngTmpLine      As Long     ' start position of next line
Dim lngTmpPos       As Long     ' next position to process
Dim lngPLoop        As Long     ' positions loop - outer loop
Dim lngLastLine     As Long     ' last line no. processed
Dim lngLLoop        As Long     ' lines loop - inner loop
Dim lngTmp          As Long     ' the positions loop pos to test
Dim lngCounter      As Long     ' line no. indexer

    On Error GoTo ErrHan:
    ' test arrays, line nos & char positions ----------------------------
    ' will trhow error if either array is not dimensioned ---------------
    lngLUBnd = UBound(miLP)
    lngPUBnd = UBound(pPositions)
    ' -------------------------------------------------------------------
    ReDim pLineNumbers(lngPUBnd)
    lngLastLine = 1
    ' loop thru' found char positions -----------------------------------
    For lngPLoop = 0 To lngPUBnd
        ' -------------------------------------------------------------------
        lngTmpPos = pPositions(lngPLoop)
        ' loop thru' lines from last line number ----------------------------
        For lngLLoop = lngLastLine To lngLUBnd
            lngTmpLine = miLP(lngLLoop)
            lngTmp = lngTmpLine - miDel
            If lngTmpPos <= lngTmp Then
                pLineNumbers(lngCounter) = lngLLoop
                lngCounter = lngCounter + 1
                Exit For
            End If
        Next lngLLoop
        ' -------------------------------------------------------------------
    Next lngPLoop
    ' -------------------------------------------------------------------
    
Exit Sub
ErrHan:

End Sub

Sub SplitResultLines(ByRef pResult As String, _
                     ByRef pResultsArray() As String, _
            Optional ByRef pItemCount As Long = 0)
' split results into array of lines, pResultsArray
' -------------------------------------------------------------------
' BUG: RETURNS WRONG ITEM COUNT
'      PERHAPS IT IS GENERATED WHEN BUILDING THE RESULTS LINES
' -------------------------------------------------------------------
' context
' we have a results string from GetAllMatches and
' we want to convert it into an itemised array

    pSplitString pResult, pResultsArray, vbCrLf, pItemCount
    
End Sub

Sub SplitResultLine(ByRef pResultLine As String, _
                    ByRef pResultElementsArray() As String, _
           Optional ByRef pItemCount As Long = 0)
' split result line into array of data elements, pResultElementsArray

' context
' we are have a single line from a results string derived from GetAllMatches and
' we want to convert it into an itemised array.

' note
' elements delimited by Chr$(1)

    pSplitString pResultLine, pResultElementsArray, Chr$(1), pItemCount
    
End Sub

Sub ReadFile(pFileName As String)
Attribute ReadFile.VB_Description = "Instruction to load/read a source string from a file."
' reads a text file for source string

' context
' we have a file we wish to read as text

Dim f As Long
Dim l 'As Long
    
    On Error GoTo ErrHan:
    
    pInit
    
    mFN = pFileName
    
    f = FreeFile
    
    Open mFN For Binary As #f
        
    l = LOF(f)
    If l Then
        ReDim mB(0 To l - 1)
        Get #f, , mB
        mTxt = StrConv(mB, vbUnicode)
    End If

ResumeError:
    
    On Error GoTo 0
    
    Close #f
    
    pPostRead
    
Exit Sub

ErrHan:

    Debug.Print "TextReader.ReadFile.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Sub

Sub ReadText(pText As String)
Attribute ReadText.VB_Description = "Instruction to load source string from another string."
' gets a source string to process as text
    
' context
' we have a string we wish to read as text
' rather than reading text from a file

    pInit
    mB = pText
    mTxt = pText
    pPostRead
End Sub

Private Function pbIsWordBreakChar(ByVal pKeyCode As Integer) As Boolean
' tests if character is a word breaker

' context/used by
' FindAllMatches

    On Error GoTo ErrHan:
    
    pbIsWordBreakChar = pKeyCode < 48 Or _
                        pKeyCode > 57 And pKeyCode < 65 Or _
                        pKeyCode > 90 And pKeyCode < 96 Or _
                        pKeyCode > 122 And pKeyCode < 127

Exit Function
ErrHan:
    Debug.Print "TextReader.pbIsWordBreakChar.Error: " & Err.Number & "; " & Err.Description

End Function

Private Sub pPostRead()
' derive basic information about the source string

' context
' a new source string has been provided (either from file or user) and we
' want to get the line delimiter, number of lines and an array of line positions

Dim i10         As Long     ' char 10
Dim i13         As Long     ' char 13
Dim c           As Long     ' count
Dim iFind       As Long     ' last find char position

    mDel = vbCrLf       ' line delimiter
    miDel = 2           ' line delimiter length
    miTL = Len(mTxt)    ' text length
    If miTL Then
        i10 = InStr(1, mTxt, vbLf)
        i13 = InStr(1, mTxt, vbCr)
        If i13 + 1 <> i10 Then
            If i10 Then
                mDel = vbLf
                miDel = 1
            End If
        End If
        ' ignore last line if empty -----------------------------------------
        If Right$(mTxt, miDel) = mDel Then
            miTL = miTL - miDel
            mTxt = Left$(mTxt, miTL)
        End If
        ' -------------------------------------------------------------------
        If miTL Then
            ReDim Preserve miLP(0 To miTL + 2)
            miLP(0) = 1
            iFind = InStr(mTxt, mDel)
            ' -------------------------------------------------------------------
            Do While iFind
                c = c + 1
                miLP(c) = iFind + miDel
                iFind = InStr(miLP(c), mTxt, mDel)
            Loop
            miLP(c + 1) = miTL  ' ubound(lines) - 1 is last line in text while
                                ' ubound(lines) is total length of text.
            ReDim Preserve miLP(c + 1)
            ' -------------------------------------------------------------------
            miLC = c + 1          ' no. of lines
        End If
    End If
    
End Sub

Private Sub pSplitString(ByRef pTheString As String, _
                         ByRef pDestinationArray() As String, _
                         ByVal pTheDelimiter As String, Optional pItemCount As Long = 0)

' After Chris Lucas, cdl1051@earthlink.net, 20011208
' http://www.xbeat.net/vbspeed/c_Split.htm

    Dim c&, SLen&, DelLen&, tmp&, Results&()

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    pItemCount = 0
    ' -------------------------------------------------------------------
    SLen = LenB(pTheString) \ 2
    DelLen = LenB(pTheDelimiter) \ 2
    ' -------------------------------------------------------------------
    ' Bail if we were passed an empty delimiter or an empty pTheString
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve pDestinationArray(0 To 0)
        pDestinationArray(0) = pTheString
        Exit Sub
    End If
    ' -------------------------------------------------------------------
    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    tmp = InStr(pTheString, pTheDelimiter)

    Do While tmp
        Results(c) = tmp
        c = c + 1
        tmp = InStr(Results(c - 1) + 1, pTheString, pTheDelimiter)
    Loop
    ' -------------------------------------------------------------------
    ' Size our return array
    ReDim Preserve pDestinationArray(0 To c)
    ' -------------------------------------------------------------------
    ' Populate the return array
    If c = 0 Then
        pItemCount = 1
        pDestinationArray(0) = pTheString
    Else
        pItemCount = c
        pDestinationArray(0) = Left$(pTheString, Results(0) - 1)
        For c = 0 To c - 2
            pDestinationArray(c + 1) = Mid$(pTheString, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
        Next c
        pDestinationArray(c + 1) = Right$(pTheString, SLen - Results(c) - DelLen + 1)
    End If
    ' -------------------------------------------------------------------
    Erase Results
    ' -------------------------------------------------------------------
Exit Sub
ErrHan:
    Debug.Print "TextReader.pSplitString", Err.Number, Err.Description

End Sub ' ... pSplitString:

Private Sub pInit()
' refresh for re-use
    pRelease
End Sub

Private Sub pRelease()
' release resources

' context
' we want to release the resources used ahead of terminating the class or re-using it

    Erase mB
    Erase miLP
    mTxt = vbNullString
    mFN = vbNullString
    mTag = vbNullString
    miTL = 0
    miLC = 0
'    mdByteSize = 0
End Sub

Private Sub Class_Terminate()
    pRelease
End Sub