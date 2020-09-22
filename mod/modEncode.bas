Attribute VB_Name = "modEncode"
Attribute VB_Description = "A module to help Encode Plain Text to Rich Text and Hyper-Text Markup."
' what?
'  a module with functions to encode plain text to syntax coloured RTF and HTML.
' why?
'  to help show plain text as code and save it in either format.
' when?
'  when you want to show syntax coloured code in a Rich Text Box or Web Page.
' how?
'  call BuildRTFString with the text to encode as the first parameter.
'  or call BuildHTMLString likewise.
'  The BuildRTFString has a RTF Colour Table parameter which allows you to
'  specify the colours that will be used to draw the text.
'  use BuildRTFColourTable with the desired colours to help do this easily.
' who?
'  d.c.

Option Explicit

' ... this constant provides the list of keywords that will be tested for colouring,
' ... mKeyWords is the actual variable tested to leave scope for a user defined list.
Private Const ckeyWords As String = " Access AddressOf Alias And Array As Binary Boolean ByRef ByVal Byte Call Case CBool CByte CCur CDate CDbl CInt Close CLng Const CSng CStr Currency CVar Date Debug Declare Dim Do Double Each Else ElseIf Empty End Enum Erase Error Event Exit Explicit False For Friend Function Get Global GoTo IIf If Implements Integer Is LBound Let Lib Like Loop Long Mid MidB Mod MsgBox New Next Not Nothing Object On Open Option Optional Or Preserve Print Private Property Public RaiseEvent Read ReDim Resume Select Set Single Static Step String Sub Then To True Type UBound Until Variant Wend While With WithEvents Xor "
Private mkeyWords As String

' ... this constant is the default colour table given to an RTF Encoding.
Private Const cColorTbl As String = "{\colortbl;\red0\green0\blue0;\red0\green0\blue128;\red0\green128\blue0;\red96\green96\blue96;\red0\green128\blue128;}"


' ... Thanks to Steve McMahon, vbAccelerator, for his TranslateColor function.
' ... It basically guards against reading a windows system colour when reading any colour.
Private Declare Function OleTranslateColor Lib "olepro32.dll" _
    (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, _
    pccolorref As Long) As Long

Private Const CLR_INVALID = -1

' ... structure to help define the rgb values of a colour.
Private Type RGBColours
    r As Long
    G As Long
    b As Long
End Type

' ... structure to help define the rtf colour table.
Private Type SyntaxColours
    Text As String
    KeyWord As String
    Comment As String
    Attribute As String
    LineNo As String
End Type

' ... this is a test line _
extended onto the next line

' -------------------------------------------------------------------
' v6, new Function for all black font colours
Function AllBlackFontColours() As String
Attribute AllBlackFontColours.VB_Description = "Returns a RTF Colour Table for the encoder here-in providing all Black Font Colours."

Dim sColourTable As String

    sColourTable = BuildRTFColourTable(0, 0, 0, 0, 0)

    AllBlackFontColours = sColourTable
    
    sColourTable = vbNullString
    
End Function

Function BuildHTMLString(pText As String, _
                Optional pFontName As String = "Courier New", _
                Optional pFontFamily As String = "fnil", _
                Optional pFontSize As String = "8", _
                Optional ByVal pTitle As String = vbNullString) As String

' ... convert plain text into syntax colour coded HTML.
' ... no points for noticing this is (or was) a direct copy of BuildRTFString.
' ... actually, it needs updating to include colour options.

' -------------------------------------------------------------------
Dim lngWordLength As Long
Dim lngWordStart As Long
Dim lngLineCount As Long
Dim lngWordCount As Long
Dim lngCharCount As Long
Dim lngChar As Long
Dim stext As String
Dim bText() As Byte
Dim lngLoop As Long
Dim sRTF As String
Dim sWord As String
Dim bControlChar As Boolean
Dim bNewLine As Boolean
Dim lngPos As Long
Dim bCommenting As Boolean
Dim bAddedCommentColour As Boolean
Dim bInAWord As Boolean
' -------------------------------------------------------------------

Dim bKillInQuotes As Boolean
Dim bInQuotes As Boolean
Dim lngQuoteCount As Long
Dim bApplyColour As Boolean
Dim lngLineWordNo As Long

Dim sHeader As String

Const cNumBufferMultiplier As Long = 12
Const cWordAttribute As String = "Attribute"
Const cWordRem As String = "Rem"

' -------------------------------------------------------------------
    
    If Len(pText) = 0 Then
        Exit Function
    End If
    
    stext = pText

    bText = stext
        
    If Len(stext) Then
        lngLineCount = 1
        lngWordStart = 1
        sRTF = Space$(Len(stext) * cNumBufferMultiplier)    ' ... set up a string buffer twelve times the length of the incoming string.
    End If
    
    For lngLoop = LBound(bText) To UBound(bText) Step 2
        lngChar = bText(lngLoop)
        ' -------------------------------------------------------------------
        ' ... attempt to capture all control chars in an effort to trap all words.
        bControlChar = lngChar < 47 Or lngChar > 57 And lngChar < 64 Or lngChar > 90 And lngChar < 96 Or lngChar > 122 And lngChar < 127
        bNewLine = lngChar = 13
        lngCharCount = lngCharCount + 1
        If bControlChar Then
            If lngChar = 34 Then                ' ... quote ".
                If bCommenting = False Then     ' v2: was absent but required.
                    If bInQuotes = False Then
                        bInQuotes = True
                    Else
                        bKillInQuotes = True        ' ... save killing in quotes until we've read the word.
                    End If
                End If
            ElseIf lngChar = 95 Then ' ... underscore _.
                bInAWord = True
            End If
            sWord = vbNullString
            If lngWordLength > 0 Then                                   ' ... have a word.
                lngLineWordNo = lngLineWordNo + 1
                sWord = Mid$(stext, lngWordStart, lngWordLength)        ' ... get the word (backwards).
                If lngLineWordNo = 1 Then
                    If bCommenting = False Then
                        If Len(sWord) = 9 Then
                            If sWord = cWordAttribute Then
                                bCommenting = True
                                sWord = "<Span Class=" & Chr$(34) & "attrib" & Chr$(34) & ">" & sWord
                                bAddedCommentColour = True
                            End If
                        ElseIf Len(sWord) = 3 Then
                            If sWord = cWordRem Then
                                bCommenting = True
                                sWord = "<Span Class=" & Chr$(34) & "comment" & Chr$(34) & ">"
                                bAddedCommentColour = True
                            End If
                        End If
                    End If
                End If
                bApplyColour = bCommenting = False And bInQuotes = False And bInAWord = False
                If bApplyColour = True Then
                    'If InStr(1, ckeyWords, " " & sWord & " ") Then
                    If InStrB(1, ckeyWords, " " & sWord & " ") Then
                        sWord = "<Span Class=" & Chr$(34) & "keyword" & Chr$(34) & ">" & sWord & "</Span>"
                    End If
                End If
                lngWordCount = lngWordCount + 1
            Else
                If lngChar = 39 Then            ' ... apostrophe '.
                    If bInQuotes = False And bCommenting = False Then
                        If bCommenting = False Then
                            sWord = "<Span Class=" & Chr$(34) & "comment" & Chr$(34) & ">"
                        End If
                        bCommenting = True
                        bAddedCommentColour = True
                    End If
                End If
            End If
            If bKillInQuotes = True Then
                bInQuotes = False
                bKillInQuotes = False   ' v3/4
            End If
            If lngChar = 32 Then ' ... space (kill word joined by underscore _.
                If bInAWord = True Then bInAWord = False
            End If
            If bNewLine = False Then
                ' -------------------------------------------------------------------
                Select Case lngChar
                    Case 34, 38, 39, 60, 62 ' ( " & ' < > ) HTML Reserved Chars.
                        sWord = sWord & "&#" & CStr(lngChar) & ";"
                    Case Else
                        sWord = sWord & Chr$(lngChar)
                End Select
                ' -------------------------------------------------------------------
                lngWordStart = lngCharCount + 1
            Else
                If lngLoop + 2 <= UBound(bText) Then
                    If bText(lngLoop + 2) = 10 Then
                        lngCharCount = lngCharCount + 1
                        lngLoop = lngLoop + 2
                        lngWordStart = lngCharCount + 1 ' ... start at next char in source string.
                        lngLineCount = lngLineCount + 1
                    End If
                    If bAddedCommentColour = True Or InStr(sWord, "<Span>") > 0 Then
                        sWord = sWord & "</Span>"
                    End If
                    sWord = sWord & vbNewLine
                    bInQuotes = False
                    bKillInQuotes = False
                    lngQuoteCount = 0
                    bAddedCommentColour = False
                    bCommenting = False
                    bInAWord = False
                    lngLineWordNo = 0
                End If
            End If
            If Len(sWord) Then
                If lngPos = 0 Then
                    lngPos = 1
                    Mid$(sRTF, lngPos, Len(sWord)) = sWord
                    lngPos = Len(sWord) + 1
                Else
                    Mid$(sRTF, lngPos, Len(sWord)) = sWord
                    lngPos = lngPos + Len(sWord)
                End If
            End If
            lngWordLength = 0
        Else
            lngWordLength = lngWordLength + 1
        End If
    Next lngLoop
    ' -------------------------------------------------------------------
    ' ... check for trailing text.
    If lngWordStart < lngCharCount Then
        lngWordLength = lngCharCount - lngWordStart + 1
        sWord = Mid$(stext, lngWordStart, lngWordLength)
        If InStr(1, mkeyWords, sWord) Then
            sWord = "<Span Class=" & Chr$(34) & "keyword" & Chr$(34) & ">" & sWord & "</Span>"
        End If
        Mid$(sRTF, lngPos, Len(sWord)) = sWord
        lngPos = lngPos + Len(sWord)
        lngWordCount = lngWordCount + 1
    End If
    ' -------------------------------------------------------------------
    ' ... trim the buffer to actual size.
    sRTF = Left$(sRTF, lngPos - 1)
    ' -------------------------------------------------------------------
    sHeader = "<html><head><style>"
    sHeader = sHeader & ".keyword {color:#000099;}"
    sHeader = sHeader & ".comment {color:#008000;}"
    sHeader = sHeader & ".attrib {color:#999999;}"
    sHeader = sHeader & "</style><title>" & pTitle & "</title><body>"
    sHeader = sHeader & "<pre>"
    sHeader = sHeader & sRTF
    sHeader = sHeader & "</pre></body></html>"
    
    BuildHTMLString = sHeader ' GetRTFHeader(pFontName, pFontFamily, pFontSize) & sRTF & "}"
    ' -------------------------------------------------------------------
    ' ... clean up used resources. Multiple Statements per line.
    sRTF = vbNullString: sWord = vbNullString
    sHeader = vbNullString
    lngWordCount = 0: lngWordLength = 0
    lngWordStart = 0
    
    lngChar = 0: lngCharCount = 0
    lngLineCount = 0: lngLoop = 0
    
End Function ' ... BuildHTMLString: String

Function BuildRTFColourTable(Optional pTextClr As Long = 0, _
                             Optional pKeyWordClr As Long = &HFF0000, _
                             Optional pCommentClr As Long = &H8000, _
                             Optional pAttributeClr As Long = &H80808, _
                             Optional pLineNoClr As Long = &H80808) As String
Attribute BuildRTFColourTable.VB_Description = "This tries to build an RTF Colour Table definition to be inserted into an rtf document using the colours provided."

' ... try and return a rtf colour table def for the main colours to use
' ... when building the rtf string.

Dim tSC As SyntaxColours
Dim sTmp As String
Dim sReturn As String

' ... Example Output: {\colortbl;\red0\green0\blue0;\red0\green0\blue128;\red128\green0\blue64;\red0\green128\blue0;\red0\green128\blue128;\red128\green128\blue128;}

    tSC.Text = pGetRTFColour(pTextClr)
    tSC.KeyWord = pGetRTFColour(pKeyWordClr)
    tSC.Comment = pGetRTFColour(pCommentClr)
    tSC.Attribute = pGetRTFColour(pAttributeClr)
    tSC.LineNo = pGetRTFColour(pLineNoClr)
    
           sTmp = tSC.Text
    sTmp = sTmp & tSC.KeyWord
    sTmp = sTmp & tSC.Comment
    sTmp = sTmp & tSC.Attribute
    sTmp = sTmp & tSC.LineNo
    
    sReturn = "{\colortbl;" & sTmp & "}"
    
    BuildRTFColourTable = sReturn
    
    sTmp = vbNullString
    sReturn = vbNullString

End Function

Function BuildRTFString(pText As String, _
               Optional pFontName As String = "Courier New", _
               Optional pFontFamily As String = "fnil", _
               Optional pFontSize As String = "8", _
               Optional pColorTbl As String = vbNullString, _
               Optional pIncludeLineNos As Boolean = False, _
               Optional pShowAttributes As Boolean = False, _
               Optional pFontBold As Boolean = False) As String
Attribute BuildRTFString.VB_Description = "This tries to convert plain text into syntax coloured coded Rich Text."
               
' ... convert plain text to vb syntax coloured rich text.
               
' -------------------------------------------------------------------
Dim lngWordLength As Long
Dim lngWordStart As Long
Dim lngLineCount As Long
Dim lngWordCount As Long
Dim lngCharCount As Long
Dim lngChar As Long
Dim stext As String
Dim bText() As Byte
Dim lngLoop As Long
Dim sRTF As String
Dim sWord As String
Dim bControlChar As Boolean
Dim bNewLine As Boolean
Dim lngPos As Long
Dim bCommenting As Boolean
Dim bAddedCommentColour As Boolean
Dim bInAWord As Boolean
' -------------------------------------------------------------------

Dim bKillInQuotes As Boolean
Dim bInQuotes As Boolean
Dim lngQuoteCount As Long
Dim bApplyColour As Boolean
Dim lngLineWordNo As Long
Dim bAddedAttribute As Boolean
' -------------------------------------------------------------------
' ... late discovery, cf0 is always black, the first colour added
' ... is read as cf1, the next cf2 and so on...

Const cF1 As String = "\cf1 "   ' ... plain text colour (black).
Const cF2 As String = "\cf2 "   ' ... keyword colour.
Const cf3 As String = "\cf3 "   ' ... comment colour.
Const cF4 As String = "\cf4 "   ' ... attribute colour.
Const cF5 As String = "\cf5 "   ' ... line no. colour

Const cFEndPar As String = "\par " & cF1
Const cNumBufferMultiplier As Long = 12
Const cWordAttribute As String = "Attribute"
Const cWordRem As String = "Rem"

Dim bCommentLineExtension As Boolean ' v7
Dim lngLastChar As Long ' v7

' -------------------------------------------------------------------
    
    On Error GoTo ErrHan:
    If Len(pText) = 0 Then
        Exit Function
    End If
    
    stext = pText

    bText = stext
        
    If Len(stext) Then
        lngLineCount = 1
        lngWordStart = 1
        sRTF = Space$(Len(stext) * cNumBufferMultiplier)    ' ... set up a string buffer twelve times the length of the incoming string.
    End If
    
    For lngLoop = LBound(bText) To UBound(bText) Step 2
        lngChar = bText(lngLoop)
        ' -------------------------------------------------------------------
        ' ... attempt to capture all control chars in an effort to trap all words.
        
        bControlChar = lngChar < 47 Or lngChar > 57 And lngChar < 65 Or lngChar > 90 And lngChar < 96 Or lngChar > 122 And lngChar < 127
        bNewLine = lngChar = 13
        lngCharCount = lngCharCount + 1
        If bControlChar Then
            If lngChar = 34 Then                ' ... quote ".
                If bCommenting = False Then     ' v2: was absent but required.
                    If bInQuotes = False Then
                        bInQuotes = True
                    Else
                        bKillInQuotes = True        ' ... save killing in quotes until we've read the word.
                    End If
                End If
            ElseIf lngChar = 95 Then ' ... underscore _.
                bInAWord = True
            End If
            sWord = vbNullString
            If lngWordLength > 0 Then                                   ' ... have a word.
                lngLineWordNo = lngLineWordNo + 1
                sWord = Mid$(stext, lngWordStart, lngWordLength)        ' ... get the word (backwards).
                If lngLineWordNo = 1 Then
                    If bCommenting = False Then
                        If Len(sWord) = 9 Then
                            If sWord = cWordAttribute Then
                                bCommenting = True
                                sWord = cF4 & sWord
                                If pShowAttributes = False Then
                                    sWord = "\v " & sWord
                                End If
                                bAddedCommentColour = True
                                bAddedAttribute = True
                            End If
                        ElseIf Len(sWord) = 3 Then
                            If sWord = cWordRem Then
                                bCommenting = True
                                sWord = cf3 & sWord
                                bAddedCommentColour = True
                            End If
                        End If
                    End If
                End If
                bApplyColour = bCommenting = False And bInQuotes = False And bInAWord = False
                If bApplyColour = True Then
                    'If InStr(1, ckeyWords, " " & sWord & " ") Then
                    If InStrB(1, ckeyWords, " " & sWord & " ") Then
                        sWord = cF2 & sWord & cF1
                    End If
                End If
                lngWordCount = lngWordCount + 1
            Else
                If lngChar = 39 Then            ' ... apostrophe '.
                    If bInQuotes = False And bCommenting = False Then
                        If bCommenting = False Then
                            sWord = cf3
                        End If
                        bCommenting = True
                        bAddedCommentColour = True
                    End If
                End If
            End If
            If bKillInQuotes = True Then
                bInQuotes = False
                bKillInQuotes = False
            End If
'            If lngChar = 32 Then ' ... space (kill word joined by underscore _.
            If lngChar <> 95 Then ' v5 fix, actually any char but underscore will kill word.
                If bNewLine = False Then ' ... v7.
                    If bInAWord = True Then bInAWord = False
                End If
            End If
            If bNewLine = False Then
                ' -------------------------------------------------------------------
                If lngChar = 92 Or lngChar = 123 Or lngChar = 125 Then
                    ' -------------------------------------------------------------------
                    ' ... rtf control chars, prefix with \ [ 92 ].
                    sWord = sWord & Chr$(92) & Chr$(lngChar)
                Else
                    ' -------------------------------------------------------------------
                    ' ... so this is where i might test for a unicode char
                    ' ... if so then try to output unicode character rather than
                    ' ... the default ? not found replacement.
                    '
                    ' cases
                    '   < 128 Normal AINSI Character
                    '    TStr = TStr & Chr$(AByte)
                    '
                    '   > 224 3 byte utf-8 group
                    '    TStr = TStr & ChrW$((ThreeBytes(0) And &HF) * &H1000 + (ThreeBytes(1) And &H3F) * &H40 + (ThreeBytes(2) And &H3F))
                    '    ChrW$((lngChar And 31) * 4096 + (NextChar And 63) * 64 + (NextNextChar And 63))
                    '    ChrW$((lngChar And &HF) * &H1000 + (NextChar And &H3F) * &H40 + (NextNextChar And &H3F))
                    '
                    '   >= 194 AND <= 219 2 byte utf-8 group
                    '    TStr = TStr & ChrW$((TwoBytes(0) And &H1F) * &H40 + (TwoBytes(1) And &H3F))
                    '    ChrW$((lngChar And 31) * 64 + (NextChar And 63))
                    '    ChrW$((lngChar And &H1F) * &H40 + (NextChar And &H3F))
                    '
                    '   else Normal AINSI Character
                    '    TStr = TStr & Chr$(AByte)
                    
                    sWord = sWord & Chr$(lngChar)
                End If
                
                ' -------------------------------------------------------------------
                lngWordStart = lngCharCount + 1
            
            Else
            
                If lngLoop + 2 <= UBound(bText) Then
                    
                    If bText(lngLoop + 2) = 10 Then
                        lngCharCount = lngCharCount + 1
                        lngLoop = lngLoop + 2
                        lngWordStart = lngCharCount + 1 ' ... start at next char in source string.
                        lngLineCount = lngLineCount + 1
                    
                    End If
                    
                    ' -------------------------------------------------------------------
                    ' ... v7, commented and extended line...
                    bCommentLineExtension = bInAWord And bCommenting And lngLastChar = 95
                    
                    If bAddedCommentColour = True Then
                        If bCommentLineExtension = False Then ' ... v7
                            ' ... restore default text colour.
                            sWord = sWord & cF1
                        End If
                    End If
                    
                    If bAddedAttribute = True Then
                        If pShowAttributes = False Then
                            ' ... restore visible text.
                            sWord = sWord & "\v0 "
                        End If
                    End If
                    
                    sWord = sWord & IIf(bCommentLineExtension, "\par ", cFEndPar) ' ... v7, updated.
                    
                    If pIncludeLineNos Then
                        'sWord = sWord & cF5 & Format$(lngLineCount, "00000:") & cF1 & vbTab & vbTab
                        sWord = sWord & cF5 & Format$(lngLineCount, "00000:") & cF1 & vbTab
'                        sWord = sWord & cF5 & "\cb3 " & Format$(lngLineCount, "00000:") & "\cb0 " & cF1 & vbTab
                    End If
                    
                    
                    bInQuotes = False
                    bKillInQuotes = False
                    lngQuoteCount = 0
                    
                    ' -------------------------------------------------------------------
                    ' ... NOTE: if bCommenting and last char pair was space underscore " _"
                    ' ...       then line extention under comment so
                    ' ...       need to carry bCommenting
                    ' -------------------------------------------------------------------
                    If bCommentLineExtension = False Then
                        bAddedCommentColour = False
                        bCommenting = False
                    End If
                    
                    bInAWord = False
                                
                    lngLineWordNo = 0
                    bAddedAttribute = False
                    
                End If
                
            End If
            
            If Len(sWord) Then
                If lngPos = 0 Then
                    lngPos = 1
                    Mid$(sRTF, lngPos, Len(sWord)) = sWord
                    lngPos = Len(sWord) + 1
                Else
                    Mid$(sRTF, lngPos, Len(sWord)) = sWord
                    lngPos = lngPos + Len(sWord)
                End If
            End If
            
            lngWordLength = 0
            lngLastChar = 0 ' v7
            
        Else
            lngWordLength = lngWordLength + 1
        End If
        
        lngLastChar = lngChar ' v7
        
    Next lngLoop
    ' -------------------------------------------------------------------
    ' ... check for trailing text.
    If lngWordStart < lngCharCount Then
        lngWordLength = lngCharCount - lngWordStart + 1
        sWord = Mid$(stext, lngWordStart, lngWordLength)
        If InStr(1, ckeyWords, sWord) Then
            sWord = cF2 & sWord & cF1
        End If
        Mid$(sRTF, lngPos, Len(sWord)) = sWord
        lngPos = lngPos + Len(sWord)
        lngWordCount = lngWordCount + 1
    End If
    ' -------------------------------------------------------------------
    ' ... trim the buffer to actual size.
    sRTF = Left$(sRTF, lngPos - 1)
    If pIncludeLineNos Then
        ' ... remove last line number added as not valid.
        ' ... 12 seems to be a magic number.
        sRTF = Left$(sRTF, (Len(sRTF) - 12)) 'Len("00000:      ")))
        ' ... add first line number.
        'sRTF = cF5 & Format$(1, "00000:") & cF1 & vbTab & vbTab & sRTF
        sRTF = cF5 & Format$(1, "00000:") & cF1 & vbTab & sRTF
    End If
    ' -------------------------------------------------------------------
    BuildRTFString = GetRTFHeader(pFontName, pFontFamily, pFontSize, pColorTbl) & IIf(pFontBold = True, "\b ", "") & sRTF & "}"
    ' -------------------------------------------------------------------

ResumeError:

    ' ... clean up used resources. Multiple Statements per line.
    sRTF = vbNullString: sWord = vbNullString
    lngWordCount = 0: lngWordLength = 0
    lngWordStart = 0
    
    lngChar = 0: lngCharCount = 0
    lngLineCount = 0: lngLoop = 0

Exit Function

ErrHan:

    Debug.Print "modEncode.BuildRTFString.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function ' ... BuildRTFString: String

Private Function pGetRTFColour(pColour As Long) As String
Attribute pGetRTFColour.VB_Description = "try to convert a colour to rtf colour table component."

' ... try to convert a colour to rtf colour table component.

Dim tClr As RGBColours
Dim lngColour As Long
Dim strColour As String
    
    ' ... desired result string in following format...
    '     "\red0\green0\blue0;"     ... black.
    '     "\red0\green0\blue128;"   ... blue.
    '     "\red0\green128\blue0;"   ... green.
    '     "\red128\green0\blue0;"   ... red.

    lngColour = TranslateColor(pColour)
    
    If lngColour <> CLR_INVALID Then
    
        ' ... convert colour to rgb values.
        tClr.r = lngColour And &HFF&
        tClr.G = (lngColour And &HFF00&) \ &H100&
        tClr.b = (lngColour And &HFF0000) \ &H10000
        
        ' ... derive return colour string.
                    strColour = "\red" & tClr.r
        strColour = strColour & "\green" & tClr.G
        strColour = strColour & "\blue" & tClr.b
        strColour = strColour & ";"
    
    Else
        ' ... return black as default, no colour identified above.
        strColour = "\red0\green0\blue0;"
    End If
    
    pGetRTFColour = strColour
    
    strColour = vbNullString
    lngColour = 0&
    
End Function

Public Function GetRTFHeader(Optional pFontName As String = "Courier New", _
                             Optional pFontFamily As String = "fnil", _
                             Optional pFontSize As String = "16", _
                             Optional pColorTbl As String = vbNullString) As String
Attribute GetRTFHeader.VB_Description = "Returns a Header for an RTF document given font name, size & family and colour table."

' ... Returns a Header for an RTF document given font name, size & family and colour table.

Dim sHeader As String
Dim sFFamily As String
Dim dFSize As Double
Dim sFSize As String
Dim sFName As String
        
    ' -------------------------------------------------------------------
    sFName = "Courier New"
    sFFamily = "fmodern"
    sFSize = "16"
    dFSize = 16
    ' -------------------------------------------------------------------
    If Len(pFontName) Then sFName = pFontName
    If Len(pFontFamily) Then sFFamily = pFontFamily
    If Len(pFontSize) Then sFSize = pFontSize
    ' -------------------------------------------------------------------
    dFSize = CDbl(sFSize)
    ' -------------------------------------------------------------------
    If dFSize > 0 Then dFSize = dFSize * 2
    ' -------------------------------------------------------------------
    sFSize = "fs" & CStr(dFSize)
    ' -------------------------------------------------------------------
    sHeader = "{\rtf1\ainsi\deff0{\fonttbl{\f0\" & sFFamily & " " & sFName & ";}}"
    
    If Len(Trim$(pColorTbl)) > 0 Then
        sHeader = sHeader & pColorTbl
    Else
        sHeader = sHeader & cColorTbl
    End If
    ' ... line numbering, can rtb do this itself?
'    sHeader = sHeader & "\sectd \linemod0\linex0\endnhere \pard \cf0 \f0 \" & sFSize & " "
'    sHeader = sHeader & "\sectd \linemod0\linex0\endnhere\pard \cf0 \f0 \" & sFSize & " "
'    sHeader = sHeader & "\sect \linemod1\linex0\pard \cf0 \f0 \" & sFSize & " "
'    sHeader = sHeader & "{\sect \linestarts1 \linemod1 \linex0 }\pard \cf0 \f0 \" & sFSize & " "

    sHeader = sHeader & "\pard \cf0 \f0 \" & sFSize & " "
    ' ... rtldoc should right align the output
'    sHeader = sHeader & "\rtldoc \pard \cf0 \f0 \" & sFSize & " "
'    sHeader = sHeader & "\rtlmark \pard \cf0 \f0 \" & sFSize & " "
    ' -------------------------------------------------------------------
    GetRTFHeader = sHeader
    ' -------------------------------------------------------------------
    sHeader = vbNullString: sFName = vbNullString
    sFFamily = vbNullString: sFSize = vbNullString
    dFSize = 0
    ' -------------------------------------------------------------------
    
End Function

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                       Optional ByRef hPal As Long = 0) As Long
Attribute TranslateColor.VB_Description = "This tries to ensure a true colour definition is used and not a windows system colour."
                             
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
    
End Function

