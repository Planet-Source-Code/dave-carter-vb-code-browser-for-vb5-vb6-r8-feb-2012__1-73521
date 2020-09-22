Attribute VB_Name = "modStrings"
Attribute VB_Description = "A module to help manage string data."

' ... Thanks to Jost Schwider (Replace).

Option Explicit

Public Function Find(ByRef pTheString As String, ByRef pFind As String, Optional ByRef pStart As Long = 1, Optional ByRef pCompareMethod As VbCompareMethod = vbBinaryCompare, Optional pWholeWordOnly As Boolean = False, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute Find.VB_Description = "Returns the starting position of the next occurance (from pStart) of the search string within the user string."
'... Parameters.
'    R__ pFind: String          ' ... The string value to find.
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
Dim lngTheTextLength As Long
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    lngCharLength = Len(pFind)
    lngTheTextLength = Len(pTheString)
    
    ' -------------------------------------------------------------------
    ' ... test there is something to search in and for and that the search and search + start are not longer than the source.
    If lngTheTextLength > 0 And lngCharLength > 0 Then
        
        ' -------------------------------------------------------------------
        ' ... ensure start position is at least 1.
        If pStart < 1 Then Let pStart = 1
        
        ' -------------------------------------------------------------------
        ' ... check that we are within limits.
        If (lngCharLength <= lngTheTextLength) And (lngCharLength + pStart <= lngTheTextLength) Then
            
            sSource = pTheString
            sQuery = pFind
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
                        If lngTheTextLength > lFoundPos + lngCharLength Then
                            lngCharRight = Asc(Mid$(sSource, lFoundPos + lngCharLength, 1))
                            If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
                        End If
                        ' -------------------------------------------------------------------
                        ' ... control char word identification, lngCharLeft + lngCharRight = 0
                        ' -------------------------------------------------------------------
                        bIsWord = lngCharLeft + lngCharRight = 0
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
    Debug.Print "ModStrings.Find", Err.Number, Err.Description

End Function

Public Function FindAllMatches(ByVal pTheText As String, _
                               ByVal pFind As String, _
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
Dim lngCharLeft As Long
Dim lngCharRight As Long
Dim bIsWord As Boolean

    On Error GoTo ErrHan:
    
    FindAllMatches = -1  ' ... default error return, should be corrected later.
    
    lngLen = Len(pTheText)
    lngFLen = Len(pFind)
    
    If lngLen = 0 Then Exit Function
    If lngFLen = 0 Then Exit Function
    If lngFLen > lngLen Then Exit Function
    If pStart > lngLen Then Exit Function
    
    lngStart = pStart
    If lngStart < 1 Then
        lngStart = 1
    End If
    
    If pCompare = vbTextCompare Then
        pTheText = LCase$(pTheText)
        pFind = LCase$(pFind)
    End If
    
ReTry:
    
    lngFound = 0
    If lngStart < lngLen Then
        lngFound = InStr(lngStart, pTheText, pFind)
    End If
    lngFound = InStr(lngStart, pTheText, pFind)
    
    If lngFound > 0 Then
        If pWholeWordOnly Then
            lngCharLeft = 0: lngCharRight = 0
            If lngFound > 1 Then
                lngCharLeft = Asc(Mid$(pTheText, lngFound - 1, 1))
                If pbIsWordBreakChar(lngCharLeft) Then lngCharLeft = 0
            End If
            If lngLen > lngFound + lngFLen Then
                lngCharRight = Asc(Mid$(pTheText, lngFound + lngFLen))
                If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
            End If
            bIsWord = lngCharLeft + lngCharRight = 0
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
        
        lngFound = InStr(lngStart, pTheText, pFind)
        If lngFound > 0 Then
            If pWholeWordOnly Then
                lngCharLeft = 0: lngCharRight = 0
                If lngFound > 1 Then
                    lngCharLeft = Asc(Mid$(pTheText, lngFound - 1, 1))
                    If pbIsWordBreakChar(lngCharLeft) Then lngCharLeft = 0
                End If
                If lngLen > lngFound + lngFLen Then
                    lngCharRight = Asc(Mid$(pTheText, lngFound + lngFLen))
                    If pbIsWordBreakChar(lngCharRight) Then lngCharRight = 0
                End If
                bIsWord = lngCharLeft + lngCharRight = 0
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
        
Exit Function

ErrHan:
    pErrMsg = Err.Description
    Err.Clear
    Resume ErrResume:
    
End Function

Public Function LeftOfComment(ByVal pTheText As String, Optional ByVal pCommentChar As String = "'", Optional pTrimRight As Boolean = True) As String
Attribute LeftOfComment.VB_Description = "Returns the left side of a line of text from the first comment char and optionally right trims the result."

' ... get the left side of a line of text from the first comment char.
' ... and optionally Right Trim the return value.
' ... expects a single char comment character although any character will do, not just the apostrophe.
' ... attempts to ignore comment chars within quotes.
' ... reads bytes rather than characters.

Dim lngPos As Long
Dim lngLoop As Long
Dim lngText As Long
Dim lngComment As Long
Dim bInQuotes As Boolean

Dim lngCommentChar As Long

Dim lngCurrentChar As Long
Dim bText() As Byte

Const cQuoteChar As Long = 34

    LeftOfComment = pTheText
    
    ' ... validation.
    
    lngText = Len(pTheText)
    lngComment = Len(pCommentChar)
    
    If lngComment <> 1 Then Exit Function
    If lngText = 0 Then Exit Function

    lngPos = InStr(1, pTheText, pCommentChar)
    
    If lngPos = 0 Then Exit Function
    
    ' ... end validation.
    
    lngPos = 0
    lngCommentChar = Asc(pCommentChar)
    
    bText = pTheText
    ' ... loop the bytes looking for the comment char.
    For lngLoop = 0 To UBound(bText) Step 2
        
        lngPos = lngPos + 1
        lngCurrentChar = bText(lngLoop)
          
        Select Case lngCurrentChar
            ' ... check if quote char and update bInQuotes accordingly.
            Case cQuoteChar: bInQuotes = Not bInQuotes
            Case lngCommentChar:
                ' ... if not in quotes and is match then found pos (-1).
                If Not bInQuotes Then Exit For
        
        End Select
        
    Next lngLoop
    
    ' ... extract left side.
    LeftOfComment = Left$(pTheText, lngPos - 1)
    If pTrimRight Then
        ' ... trim right side if required.
        LeftOfComment = RTrim$(LeftOfComment)
    End If
    
    ' ... clean up.
    Erase bText
    lngPos = 0&
    lngText = 0&
    lngComment = 0&
    lngCommentChar = 0&

End Function

Public Sub SplitStringPair(ByVal TheString As String, ByVal TheDelimiter As String, ByRef LeftSide As String, ByRef RightSide As String, Optional ByVal TrimLeft As Boolean = False, Optional ByVal TrimRight As Boolean = False)
Attribute SplitStringPair.VB_Description = "Splits a string in two via a delimiter, returning left and right sides."
Dim lngFound As Long
' -------------------------------------------------------------------
' ... helper to split a string in two via a delimiter returning left and right.
' ... e.g.  SplitStringPair("Visual Basic", " ", sLeft, sRight)
' ...       sLeft = "Visual"
' ...       sRight = "Basic"
' ... Primitive, not sensitive to delimiter being within quotes.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    LeftSide = TheString
    RightSide = vbNullString    ' ... clear right side in case ref value is reused in loop or something.
    If Len(TheString) Then
        If Len(TheDelimiter) Then
            lngFound = InStr(1, TheString, TheDelimiter)
            If lngFound > 0 Then
                LeftSide = Left$(TheString, lngFound - 1)
                RightSide = Mid$(TheString, lngFound + Len(TheDelimiter))
            End If
        End If
        If TrimRight Then RightSide = Trim$(RightSide)
    End If
    If TrimLeft Then LeftSide = Trim$(LeftSide)
Exit Sub
ErrHan:
    Debug.Print "modStrings.SplitStringPair.Error: " & Err.Number & "; " & Err.Description
End Sub ' ... pSplitString:

Public Function PadStrings(pString1 As String, pString2 As String, pDistance As Long, Optional pPadding As Long = 0, Optional pNewLinePadding As Long = -1) As String
Attribute PadStrings.VB_Description = "Concatenates two strings formatting them into two columns."

Dim lS1Len As Long
Dim lS2Len As Long

Dim lTmpLen As Long
Dim sTmp As String
    
    lS1Len = Len(pString1)
    lS2Len = Len(pString2)
    
    If pDistance < 1 Then pDistance = 3
    If pPadding < 1 Then pPadding = 0
    
    lTmpLen = pDistance + pPadding
    lTmpLen = lTmpLen + lS2Len
    
    sTmp = Space$(lTmpLen)
    
    If lS1Len <= pDistance - 1 Then
        Mid$(sTmp, 1, lS1Len) = pString1
    Else
        Mid$(sTmp, 1, pDistance) = Left$(pString1, pDistance - 2) & ".."
    End If
    
    If lS2Len > 0 Then
        Mid$(sTmp, pDistance + pPadding + 1, lS2Len) = pString2
    End If
    
    PadStrings = sTmp
    sTmp = vbNullString
    
End Function

Public Sub FilterChar(ByRef pTheString As String, pFilterChar As String, Optional pStart As Long = 1, Optional ByRef pNoOfReplacements As Long = -1, Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare)
Attribute FilterChar.VB_Description = "Filters a single character from a string."
        
    ' ... wrapper to ReplaceChar simply for removing a single character from a source string.
    
    ReplaceChar pTheString, pFilterChar, vbNullString, pStart, pNoOfReplacements, pCompare
    
End Sub

Function InstrRev(sCheck As String, sMatch As String, Optional lStart As Long, Optional pCompare As VbCompareMethod = vbBinaryCompare) As Long
Attribute InstrRev.VB_Description = "Returns the start position of a substring in a string right to left with optional start pos."

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
        ' ... delegate to single match method, InStrRevChar.
        InstrRev = InStrRevChar(sCheck, sMatch, lStart, pCompare)
        
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
                    InstrRev = k
                    Exit For
                End If
                
            End If
            
            k = k - 1
            
        Next lngLoop
    
        Erase sBytes
        Erase sMatchBytes
        
    End If

End Function

Public Function InStrRevChar(ByVal StringCheck As String, ByVal StringMatch As String, Optional Start As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
Attribute InStrRevChar.VB_Description = "Returns the position of a character in a string right to left with optional start pos."

' ... Note:
' ...       Only tests for a single character for StringMatch.
' ... Updated in (code browser) v6 when single char match was not found.

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
        If bOK = False Then Exit Function ' ... return 0.
        
        ' ... handle text compare.
        If Compare = vbTextCompare Then
            StringCheck = LCase$(StringCheck)
            StringMatch = LCase$(StringMatch)
        End If
        
        ' ... handle single source char, v6.
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
        If lngStart > 0 Then  ' v6 updated, failed to read single char
                                                                ' create case above on single char matching single find char.
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

    InStrRevChar = lngReturn
    
End Function

Private Function pbIsWordBreakChar(ByVal pKeyCode As Integer) As Boolean
Attribute pbIsWordBreakChar.VB_Description = "Tests if a keycode is a non-typeable character or a character that would break a word in VB."
    
    On Error GoTo ErrHan:
    
    pbIsWordBreakChar = pKeyCode < 48 Or pKeyCode > 57 And pKeyCode < 65 Or pKeyCode > 90 And pKeyCode < 96 Or pKeyCode > 122 And pKeyCode < 127

Exit Function

ErrHan:

    Debug.Print "modStrings.pbIsWordBreakChar.Error: " & Err.Number & "; " & Err.Description

End Function

Sub RemoveQuotes(ByRef pTheString As String)
Attribute RemoveQuotes.VB_Description = "Removes Quote Characters from  a string."

Const cQuoteChar As Long = 34
    
    ' ... test data.
    ' ... pTheString = "This String has " & chr(34) & chr(34) & " quotes chars. " & chr(34) & " are they still here?" & chr(34)
    ReplaceChar pTheString, Chr$(cQuoteChar), vbNullString
    
End Sub

Function Replace(pExpression As String, pFind As String, pReplace As String, Optional pStart As Long = 1, Optional pCount As Long = -1, Optional pCompare As VbCompareMethod = vbBinaryCompare) As String
Attribute Replace.VB_Description = "Replaces the occurance of a substring within a source string with a different string."

Dim sTmp As String
    
    sTmp = pExpression
    
    If pCompare = vbTextCompare Then
        sTmp = LCase$(sTmp)
        pFind = LCase$(pFind)
    End If
    
    ReplaceChars sTmp, pFind, pReplace, pStart, pCount
    
    Replace = sTmp
    
    sTmp = vbNullString
    
End Function




' Sub:             Replace
' Description:     Replace a string within another string with a different string :).

Public Sub ReplaceChars(pTheString As String, ByRef pSearchFor As String, ByRef pReplaceWith As String, Optional pStart As Long = 1, Optional ByRef pCountReplaced As Long = -1, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ReplaceChars.VB_Description = "Replaces a set of characters found in a string with a different string."

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
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Search = pTheString
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
        
        Let pTheString = Buffer
        
    End If
    
    pOK = True ' ... executed, no errors, doesn't mean anything was replaced though.
    
ResumeError:

Exit Sub

ErrHan:
    
    pOK = False
    Let pErrMsg = Err.Description
    
    Debug.Print "Error.modStrings.Replace.Error: " & Err.Description

    Resume ResumeError:

End Sub ' ... Replace.

Public Sub ReplaceChar(ByRef pTheString As String, ByVal pFind As String, ByVal pReplace As String, Optional ByVal pStart As Long = 1, Optional ByRef pNoOfReplacements As Long = -1, Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare)
Attribute ReplaceChar.VB_Description = "Replaces the occurance of a single character in a string with another character."

' ... Replaces a single character in a string with another character
' ... if the replace is empty char is filtered from the string.
' ... Does this in a loop stepping thru' the bytes of the input string.
' ... When [p]Start is > 1, replacement begins after current Char. at Index = [p]Start.

' ... sBytes is the byte array of the input string and is used to fill output buffer.
' ... sTestBytes is the byte array of the input string and is used to test for char. match.
' ... because the test bytes array may have become lower cased we won't add the test char
' ... to the output byte buffer.

Dim sBytes() As Byte
Dim sTestBytes() As Byte
Dim sFindBytes() As Byte
Dim sReplaceBytes() As Byte

Dim sTheString As String

Dim lngFindChar As Long
Dim sReplaceChar As Long

Dim sReturn() As Byte

Dim lngChar As Long

Dim lngLen As Long
Dim lngLoop As Long
Dim lngLBnd As Long
Dim lngUBnd As Long
Dim lngMaxChars As Long

Dim bFindExists As Boolean
Dim bReplace As Boolean
Dim lngCharCounter As Long

    lngLen = Len(pTheString)
    
    pNoOfReplacements = 0
    
    ' ... bit of validation.
    If lngLen = 0 Then Exit Sub                 ' ... nothing to search in.
    If pStart > lngLen Then Exit Sub            ' ... start beyond scope.
    If Len(pFind) <> 1 Then Exit Sub        ' ... find length > 1: invalid in the method.
    If Len(pReplace) > 1 Then Exit Sub      ' ... replace length > 1: invalid in the method.
    
    If pStart < 1 Then pStart = 1
    bFindExists = InStr(pStart, pTheString, pFind, pCompare) > 0
    
    If bFindExists = False Then Exit Sub        ' ... find isn't in the source.
        
        
    ' ... convert the input to bytes for looping.
    ' ... if text compare the convert text to lower case
    
    sBytes = pTheString
    
    If pCompare = vbTextCompare Then
        pFind = LCase$(pFind)
        sTheString = LCase$(pTheString)
        sTestBytes = sTheString
    Else
        sTestBytes = pTheString
    End If
    
    sFindBytes = pFind
    lngFindChar = sFindBytes(0)                 ' ... read the find char byte value.
    
    
    lngLBnd = LBound(sTestBytes)                ' ... set up loop boundaries
    lngUBnd = UBound(sTestBytes)
    
    
    ' ... grab max chars in input string.
    lngMaxChars = lngUBnd \ 2
    
    ' ... dimension the output array
    ReDim sReturn(lngMaxChars)
    
    If pStart > 1 Then
        ' ... move to start position writing bytes to the output buffer.
        For lngLoop = 0 To lngUBnd Step 2
            lngChar = sBytes(lngLoop)           ' ... byte value of current char.
            sReturn(lngCharCounter) = lngChar
            lngCharCounter = lngCharCounter + 1 ' ... increment position.
            If lngCharCounter = pStart Then
                lngLBnd = lngLoop + 2           ' ... move start byte 2 places forward.
                Exit For                        ' ... and escape.
            End If
        Next lngLoop
    End If
        
    ' ... rather than test this within each iteration
    ' ... write the two possibilities in separate loops.
    If Len(pReplace) = 0 Then
        ' ... just removing the search char, acts like a filter.
        ' ... loop thru' the bytes and build the
        ' ... output bytes without the char to find.
        For lngLoop = lngLBnd To lngUBnd Step 2
            lngChar = sTestBytes(lngLoop)                       ' ... byte value of test char.
            bReplace = lngChar = lngFindChar
            If bReplace = True Then
                pNoOfReplacements = pNoOfReplacements + 1       ' ... increment replacement count.
            Else
                lngChar = sBytes(lngLoop)                       ' ... byte value of current char.
                sReturn(lngCharCounter) = lngChar               ' ... write current char to the output buffer.
                lngCharCounter = lngCharCounter + 1             ' ... next char pos.
            End If
        Next lngLoop
    
    Else
        ' ... replace find with replace.
        ' ... validation above says replace length must be 1.
        sReplaceBytes = pReplace
        sReplaceChar = sReplaceBytes(0)
    
        ' ... loop thru' the bytes and build the
        ' ... output byte array replacing search char.
        For lngLoop = lngLBnd To lngUBnd Step 2
            lngChar = sTestBytes(lngLoop)                       ' ... byte value of test char.
            bReplace = lngChar = lngFindChar                    ' ... test for match.
            If bReplace Then
                ' ... change char to replace char.
                lngChar = sReplaceChar                          ' ... swap find with replace.
                pNoOfReplacements = pNoOfReplacements + 1       ' ... increment replacement count.
            Else
                lngChar = sBytes(lngLoop)                       ' ... byte value of current char.
            End If
            sReturn(lngCharCounter) = lngChar                   ' ... write char to the output buffer.
            lngCharCounter = lngCharCounter + 1                 ' ... next char pos.
        Next lngLoop
    
    End If
                    
    If lngCharCounter > 0 Then
       
        ReDim Preserve sReturn(lngCharCounter - 1)              ' ... resize the output byte array downwards.
        ' ... convert the output array back to a string
        ' ... and use it to overwrite the input string.
        pTheString = StrConv(sReturn, vbUnicode)
'        Debug.Print pTheString
    End If
            
    
End Sub

Sub ReverseString(ByRef pTheString As String)
Attribute ReverseString.VB_Description = "Returns a string whose characters are in reverse order to the input string."

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
        
    End If

End Sub

Function StrReverse(ByVal pTheString As String) As String
Attribute StrReverse.VB_Description = "Reverses the input string."

' ... VB5 substitute for VB6.VBA.StrReverse

Dim sTmp As String

    sTmp = pTheString
    ReverseString sTmp
    
    StrReverse = sTmp
    
    sTmp = vbNullString
    
End Function

Function WrapInQuoteChars(pvSource As String) As String
    
'    WrapInQuoteChars = " Chr$(34) & " & Chr$(34) & pvSource & Chr$(34) & " & Chr$(34)"
    WrapInQuoteChars = Chr$(34) & pvSource & Chr$(34)
    
End Function
