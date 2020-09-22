VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "a class to help build large strings"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' sbuilder
' a class to help building string data, with reading from and writing to file,
' inserting at a given position and deleting a continuous length of characters

Option Explicit

Private mTxt As String                          ' ... private field for property TheString.
Private mAddLen As Long                         ' ... private field for property PreAllocationSize.
Private mBufLen As Long
Private mTxtLen As Long                         ' ... private field for property Length.

Private Const cDefAddLen As Long = 8192         ' ... default value for property PreAllocationSize.

Public Name As String
Attribute Name.VB_VarDescription = "A name for the instance of a class."
Public Tag As String
Attribute Tag.VB_VarDescription = "A Tag to give an instance of this class."

Public Sub Append(ByRef pString As String)
Attribute Append.VB_Description = "Appends a string to the end of the current user string."
' Append a string value to the user string.

Dim iStrLen As Long
Dim iTxtLen As Long

    On Error GoTo ErrHan:

    iStrLen = Len(pString)
    iTxtLen = mTxtLen    ' ... save original user text length, pResizeBuffer will increase this by Len(pString).
    
    pResizeBuffer iStrLen

    If iTxtLen = 0 Then
        If iStrLen > 0 Then
            Mid$(mTxt, 1, iStrLen) = pString
        End If
    Else
        Mid$(mTxt, iTxtLen + 1, iStrLen) = pString
    End If
                
Exit Sub
ErrHan:
    Debug.Print "SBuilder.Append", Err.Number, Err.Description

End Sub ' ... Append.

Public Sub AppendAsLine(Optional ByRef pTheString As String = vbNullString, _
                        Optional ByVal pBefore As Boolean = False, _
                        Optional ByVal pAfter As Boolean = True)
Attribute AppendAsLine.VB_Description = "As  Append, but allows for adding a new line before and/or after the string to be added."
' Append a string value to the user string with optional new line before and after.
    Append IIf(pBefore, vbNewLine, vbNullString) & _
            pTheString & _
             IIf(pAfter, vbNewLine, vbNullString)

End Sub ' ... AppendAsLine.

Public Function BufferLength() As Long
Attribute BufferLength.VB_Description = "Returns the current length of the internal string buffer."
' returns the current length of the internal string buffer.
    BufferLength = mBufLen

End Function

Private Sub Class_Initialize()
    mAddLen = cDefAddLen

End Sub

Private Sub Class_Terminate()
    mAddLen = 0
    mTxt = vbNullString

End Sub ' ... Class_Terminate:

Public Sub Compact(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Compact.VB_Description = "Reduces the size of the internal buffer to the length of the current user string."
' Reduce the available buffer to the size of the current user string.
    pResizeBuffer -1, pOK, pErrMsg

End Sub ' ... Compact.

Public Sub DeleteAll()
Attribute DeleteAll.VB_Description = "Deletes the current user string from the buffer."
' deletes the current user string from the buffer.

    mTxt = vbNullString
    mTxtLen = 0
    mBufLen = 0
    
End Sub ' ... DeleteAll.

Public Sub DeletePortion(ByVal pStart As Long, ByVal pLength As Long, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute DeletePortion.VB_Description = "Deletes a section of the string data from a starting character onwards for a given length."
' Deletes a portion of text from the user string by start and length to remove.

    On Error GoTo ErrHan:

    If mTxtLen > 0 And pLength >= 1 Then
        
        ' ... ensure start position is at least 1.
        If pStart < 1 Then Let pStart = 1
        
        If pStart > mTxtLen Then
            ' ... start > length, trim to user string length.
            Compact pOK, pErrMsg
        Else
            ' ... if start and delete length > length reduce pLength to fit.
            If pStart + pLength > mTxtLen Then
                Let pLength = mTxtLen - pStart + 1
            End If
            ' ... shift remaining text following start and del. len. to start.
            Mid$(mTxt, pStart, mTxtLen - pLength) = Mid$(mTxt, pStart + pLength, mTxtLen - pLength)
            ' ... cut user string text length to suit; note shifting didn't delete from the length.
            mTxtLen = mTxtLen - pLength
            pOK = True
        
        End If
    Else
        pErrMsg = "Either there is nothing to delete from or nothing to delete."
    End If

Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.DeletePortion", Err.Number, Err.Description

End Sub ' ... DeletePortion.

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
    If mTxtLen > 0 And lngCharLength > 0 Then 'And ((lngCharLength <= mTxtLen) And (lngCharLength + pStart <= mTxtLen)) Then
        
        ' -------------------------------------------------------------------
        ' ... ensure start position is at least 1.
        If pStart < 1 Then Let pStart = 1
        
        ' -------------------------------------------------------------------
        ' ... check that we are within limits.
        If (lngCharLength <= mTxtLen) And (lngCharLength + pStart <= mTxtLen) Then
            
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
                        If mTxtLen > lFoundPos + lngCharLength Then
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
    
    lngLen = Len(mTxt)
    lngFLen = Len(pFind)
    
    If lngLen = 0 Then Exit Function
    If lngFLen = 0 Then Exit Function
    If lngFLen > lngLen Then Exit Function
    If pStart > lngLen Then Exit Function
    
    sTheText = mTxt
    
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


'Public Function Find(ByVal pFindWhat As String, _
'            Optional ByVal pStart As Long = 0, _
'            Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare) As Long
'' find a substring within the main string
'' returns -1 if error occurred else 0 = no match, 1 = start of match
'
'' StringWorker Find
''Public Function Find(ByRef pTheString As String, _
''            Optional ByRef pStart As Long = 1, _
''            Optional ByRef pCompareMethod As VbCompareMethod = vbBinaryCompare, _
''            Optional ByVal pWholeWordOnly As Boolean = False, _
''            Optional ByRef pOK As Boolean = False, _
''            Optional ByRef pErrMsg As String = vbNullString) As Long
'
'
'Dim iRet As Long
'
'    On Error GoTo ErrHan:
'    ' -------------------------------------------------------------------
'
'    If pStart < 1 Then pStart = 1
'    iRet = InStr(pStart, Me, pFindWhat, pCompare)
'
'ResumeError:
'    ' -------------------------------------------------------------------
'    Find = iRet
'    ' -------------------------------------------------------------------
'
'Exit Function
'
'ErrHan:
'    iRet = -1   ' return error status
'    Debug.Print "SBuilder.Find.Error: " & Err.Number & "; " & Err.Description
'    Resume ResumeError:
'
'End Function ' ... Find: Long

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

Public Sub Insert(ByRef pTheString As String, _
                  ByVal pStart As Long, _
         Optional ByRef pOK As Boolean = False, _
         Optional ByRef pErrMsg As String = vbNullString)
Attribute Insert.VB_Description = "Inserts a string into the current user string at the position given."
' inserts a text string at the position specified.

Dim lngShiftLength As Long
Dim lngInsertLength As Long
Dim lngBufferResize As Long

    On Error GoTo ErrHan:

    lngInsertLength = Len(pTheString)
    
    If lngInsertLength > 0 Then ' ... only attempt insert if insert string not "".
    
        If pStart < 1 Then pStart = 1         ' ... Mid$ won't process zero start position.
        
        lngShiftLength = mTxtLen - pStart + 1  ' ... Shift length is length of text from insert position to end of user string before any processing.
        If lngShiftLength < 0 Then lngShiftLength = 0           ' ... Mid$ won't process a negative length.
            
        ' ... allow insert position to be beyond current length.
        lngBufferResize = lngInsertLength + IIf(pStart > mTxtLen, pStart - mTxtLen, 0)
        
        ' ... resize the buffer and user text length to include new text @ position given.
        pResizeBuffer lngBufferResize, pOK, pErrMsg
        
        If pOK = True Then
            
            
            pOK = False ' ... reset return to false.
            
            ' ... shift trailing text to end of insert position within the user text string.
            Mid$(mTxt, pStart + lngInsertLength, lngShiftLength) = Mid$(mTxt, pStart, lngShiftLength)
            
            ' ... insert the new test at the position given.
            Mid$(mTxt, pStart, lngInsertLength) = pTheString
            
            pOK = True ' ... upate return to true.
            
        End If
        
    Else
    
        pErrMsg = "Insert String Length needs to be > 0."
    
    End If

Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.Insert", Err.Number, Err.Description

End Sub ' ... Insert.

Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the length of the current user string (no. of characters)."
' returns the length of the current user string (no. of characters).
    Length = mTxtLen

End Property ' ... Length: Long.

Private Function pFileExists(ByRef pTheFileName As String, _
                    Optional ByRef pErrMsg As String = vbNullString) As Boolean
' tests for file existence, returns True if found else False.
Dim i As Long

    On Error GoTo ErrHan:
    
    i = GetAttr(pTheFileName)
    pFileExists = True

Exit Function
ErrHan:

    pFileExists = False
    pErrMsg = "Error Reading: " & pTheFileName & vbNewLine & Err.Description

End Function ' ... pFileExists: Boolean.

Private Sub pReadTextFile(ByRef pTheFileName As String, _
                 Optional ByRef pOK As Boolean = False, _
                 Optional ByRef pErrMsg As String = vbNullString)
                      
' Reads text from a file into the current user string.

Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
Dim bFileExists As Boolean

    On Error GoTo ErrHan:
    
    bFileExists = pFileExists(pTheFileName, pErrMsg)
    
    If bFileExists = True Then
        
        iFileNumber = FreeFile()
        Open pTheFileName For Binary As #iFileNumber
        bFileIsOpen = True
        
        mTxt = String$(LOF(iFileNumber), 0)
        Get #iFileNumber, , mTxt
        If Len(mTxt) > 2 Then
            If Right$(mTxt, 2) = vbCrLf Then
                mTxt = Left$(mTxt, Len(mTxt) - 2)
            End If
        End If
        
        mTxtLen = Len(mTxt)
        mBufLen = mTxtLen
        
        pOK = True                   ' ... return success.
    
    End If
    
ErrResume:

    On Error Resume Next
    
    If bFileIsOpen = True Then
        Close #iFileNumber
        If Err.Number <> 0 Then
            Debug.Print "SBuilder.pReadTextFile (Close File)", Err.Number, Err.Description
            Err.Clear
        End If
    End If
    
Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.pReadTextFile", Err.Number, Err.Description
    
    Err.Clear
    
    Resume ErrResume:
    
End Sub      ' ... pReadTextFile: String.

Public Property Get PreAllocationSize() As Long
Attribute PreAllocationSize.VB_Description = "Sets / Returns the size (no. of Space Characters) to add to the buffer when it is not large enough to add the string passed."
' Sets / Returns the size (no. of Space Characters) to add to the buffer when it is not large enough to add the string passed.

    PreAllocationSize = mAddLen

End Property

Public Property Let PreAllocationSize(ByVal pNewValue As Long)

    If pNewValue >= 1 Then
        mAddLen = pNewValue
    End If

End Property

Public Sub PrefixWith(ByRef pTheText As String, _
             Optional ByVal pNewLineBefore As Boolean = False, _
             Optional ByVal pNewLineAfter As Boolean = False, _
             Optional ByRef pOK As Boolean = False, _
             Optional ByRef pErrMsg As String = vbNullString)
Attribute PrefixWith.VB_Description = "Adds string data to the beginning of the user string."
' adds a string to the beginning of the user string.

    ' ... just delegate to insert with zero pos for pre-pend.
    Insert IIf(pNewLineBefore, vbNewLine, vbNullString) & _
            pTheText & _
             IIf(pNewLineAfter, vbNewLine, vbNullString), 0, pOK, pErrMsg

End Sub ' ... PrefixWith.

Private Sub pResizeBuffer(Optional ByVal pLength As Long = -1, _
                          Optional ByRef pOK As Boolean = False, _
                          Optional ByRef pErrMsg As String = vbNullString)
                          
' manages changing the size of the internal buffer.

    On Error GoTo ErrHan:

    Select Case pLength
        
        Case 0, -1 ' ... compact.
            mTxt = Left$(mTxt, mTxtLen)
            
        Case Is > 0 ' ... grow.
            If (mTxtLen + pLength) > mBufLen Then
                If (pLength > mAddLen) Then
                    mTxt = mTxt & Space$(pLength)
                Else
                    mTxt = mTxt & Space$(mAddLen)
                End If
            End If
            mTxtLen = mTxtLen + pLength
        
        Case Else ' ... shrink; user given length.
            pLength = pLength * -1
            
            If mTxtLen - pLength > 0 Then
                mTxtLen = mTxtLen - pLength
                mTxt = Left$(mTxt, mTxtLen)
            End If
    
    End Select

    mBufLen = Len(mTxt)
    
    pOK = True

Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.pResizeBuffer", Err.Number, Err.Description

End Sub ' ... pResizeBuffer.

Private Sub pWriteTextFile(ByRef pTheText As String, _
                           ByRef pTheFileName As String, _
                  Optional ByVal pConfirmOverWrite As Boolean = False, _
                  Optional ByRef pOK As Boolean = False, _
                  Optional ByRef pErrMsg As String = vbNullString)
' writes the current user string as text to a text file.

Dim iFNum As Integer
Dim bFOpen As Boolean
Dim bFExists As Boolean
Dim iAnswer As VbMsgBoxResult

    On Error GoTo ErrHan:
        
    If pConfirmOverWrite Then
    
        bFExists = pFileExists(pTheFileName, pErrMsg)
        
        If bFExists = True Then
        
            iAnswer = MsgBox("The File [ " & pTheFileName & " ] Exists;" & vbCrLf & "Would you like to over-write this File?", vbYesNo + vbQuestion + vbDefaultButton1, "Over-Write File?")
            
            If iAnswer = vbYes Then
                Kill pTheFileName
            Else
                pOK = False
                pErrMsg = "User cancelled Save"
                GoTo ErrResume:
            End If
            
        End If
        
    End If
    
    iFNum = FreeFile()
    
    Open pTheFileName For Output As #iFNum
    
    bFOpen = True
    
    Print #iFNum, pTheText;
    
    pOK = True
        
ErrResume:

    On Error Resume Next
    
    If bFOpen = True Then
        Close #iFNum
        If Err.Number <> 0 Then
            Debug.Print "SBuilder.pWriteTextFile (Close File)", Err.Number, Err.Description
            Err.Clear
        End If
    End If

Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.pWriteTextFile", Err.Number, Err.Description
    
    Resume ErrResume:
    
End Sub ' ... pWriteTextFile.

Public Sub ReadFromFile(ByRef pTheFileName As String, _
               Optional ByRef pOK As Boolean = False, _
               Optional ByRef pErrMsg As String = vbNullString)
' reads the contents of a text file and makes it to the user string, overwriting anything else.

    On Error GoTo ErrHan:
    
    pReadTextFile pTheFileName, pOK, pErrMsg
    
    If pOK = False Then
        
        pErrMsg = "ReadFromFile: Did not Read the File: " & pTheFileName & vbCrLf & pErrMsg
    
    End If
    
Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.ReadFromFile", Err.Number, Err.Description
    
End Sub ' ... ReadFromFile.

Public Property Get TheString() As String
Attribute TheString.VB_Description = "Sets / Returns the user string."
Attribute TheString.VB_UserMemId = 0
' Returns the user string.
    If mTxtLen > 0 Then
        TheString = Left$(mTxt, mTxtLen)
    End If

End Property ' ... TheString: String.

Public Property Let TheString(ByVal pNewValue As String)
' Sets the user string.
    On Error GoTo ErrHan:
    
    mTxt = pNewValue
    mTxtLen = Len(mTxt)
    mBufLen = mTxtLen

Exit Property
ErrHan:
    ' ... error 14 is ' out of string space '.
    Debug.Print "SBuilder.TheString", Err.Number, Err.Description

End Property

Public Sub WriteToFile(ByRef pTheFileName As String, _
         Optional ByVal pConfirmOverWrite As Boolean = False, _
         Optional ByRef pOK As Boolean = False, _
         Optional ByRef pErrMsg As String = vbNullString)
Attribute WriteToFile.VB_Description = "Writes the current user string to file."
' writes the contents of the current user string within the buffer to disk.

Dim sTheFileText As String

    On Error GoTo ErrHan:
    
    sTheFileText = Me
    pWriteTextFile sTheFileText, pTheFileName, pConfirmOverWrite, pOK, pErrMsg

Exit Sub
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "SBuilder.WriteToFile", Err.Number, Err.Description

End Sub ' ... WriteToFile.

Public Sub Replace(ByRef pSearchFor As String, _
                   ByRef pReplaceWith As String, _
          Optional ByVal pStart As Long = 1, _
          Optional ByRef pCountReplaced As Long = 0, _
          Optional ByRef pOK As Boolean = False, _
          Optional ByRef pErrMsg As String = vbNullString)

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
        
        TextLen = LenB(Search)
        
        Select Case NewLen
        
            Case OldLen ' ... query and replace are same size, no effect on string length.
            
                Buffer = Search
                
                Do Until pStart = 0
                    MidB$(Buffer, pStart) = pReplaceWith
                    pStart = InStrB(pStart + OldLen, Buffer, pSearchFor)
                    pCountReplaced = pCountReplaced + 1
                Loop
            
            Case Is < OldLen ' ... replace is smaller than query string, string length will shrink.
                            
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
                            Buffer = Buffer & Space$(BufferPosNext) ' ... grow the buffer.
                            BufferLen = LenB(Buffer)
                        End If
                        
                        MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                        MidB$(Buffer, BufferPosNew) = pReplaceWith
                    
                    Else
                        
                        BufferPosNext = WritePos + NewLen
                        
                        If BufferPosNext > BufferLen Then
                            Buffer = Buffer & Space$(BufferPosNext) ' ... grow the buffer.
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
        
        TheString = Buffer
        
    End If
    
    pOK = True ' ... executed, no errors, doesn't mean anything was replaced though.
    
ResumeError:

Exit Sub

ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    
    Debug.Print "StringWorker.Replace.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:

End Sub ' ... Replace.

