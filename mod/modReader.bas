Attribute VB_Name = "modReader"
Attribute VB_Description = "A module to provide functions to help reading the contents of text files."

' what?
'  a module dedicated to reading text lines from file or string.
' why?
'  looking into different ways to read text lines.
' when?
'  processing text written in lines.
' how?
'  under development
' who?
'  d.c.

Option Explicit

' ... Note:
' ... My plan was to be able to read / extract lines of text from a string
' ... without splitting the string into an array.
' ... I wanted to search the text and know which lines matches were found on
' ... based upon the char. pos. of the first character in the search string.

' ... a few module fields to avoiding using Statics when
' ... testing re-entrancy (seeing as this is a module).

Private mReadingFile As Boolean ' ... reading a file.
Private mReadingText As Boolean ' ... resding some text.
Private mMaking As Boolean      ' ... generating the line array.


Private Const cMsgInUse As String = "The Program is busy, please wait and try again."

Public Sub DeriveLineLengths(ByRef pTheLines() As Long, ByRef pLineNumbers() As Long, ByRef pLineLengths() As Long, Optional pDelimiterLength As Long = 2, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute DeriveLineLengths.VB_Description = "Return a subset of text line lengths from a list of lines and line numbers."

' ... given an array of line start positions and an array of line numbers
' ... return an array of line lengths for the lines in the array of line numbers.
    
Dim k As Long
Dim max As Long
Dim lIndex As Long
Dim lLen As Long

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... quit if either lines array or line numbers array is undimensioned.
    max = UBound(pTheLines)
    max = UBound(pLineNumbers)
    
    ReDim pLineLengths(max)
    
    For k = 0 To max
        
        lIndex = pLineNumbers(k)
        
        If lIndex < 1 Then
            
            If k = 0 Then
                ' ... first line.   ' v6
                lLen = pTheLines(1) - pDelimiterLength
            Else
                ' ... last line.    ' v6
                lIndex = UBound(pTheLines) - 1
                lLen = pTheLines(lIndex + 1) - pTheLines(lIndex) + 1
            End If
            
        Else
            ' ... all lines in-between first and last.
            lLen = pTheLines(lIndex + 1) - pTheLines(lIndex) - pDelimiterLength
        End If
        
        pLineLengths(k) = lLen
        
    Next k
    
    pOK = True: pErrMsg = vbNullString
    
ResErr:

    ' ... clean up.
    k = 0&
    max = 0&
    lIndex = 0&
    lLen = 0&
    
Exit Sub
ErrHan:
    pOK = False
    pErrMsg = "modReader.DeriveLineLengths.Error: " & Err.Description
    Err.Clear
    Resume ResErr:
    
End Sub

Public Sub DeriveLineNumbers(ByRef pLines() As Long, ByRef pPositions() As Long, ByRef pLineNumbers() As Long, Optional pDelimiterLength As Long = 2)
Attribute DeriveLineNumbers.VB_Description = "Generates an array of longs that represents the position of numbers within a range of numbers."

' ... this method is intended to return an array of numbers indicating the position of
' ... a number (from pPositions) within a range of numbers (in pLines).
' ... in so doing it is hoped to be able to find the line numbers corresponding to
' ... the start of a substring within a source string which represents lines of text.

' ... by rights, line pos numbers and char pos numbers should increase.

Dim lngLUBnd As Long
Dim lngPUBnd As Long
Dim lngTmpLine As Long
Dim lngTmpPos As Long
Dim lngPLoop As Long
Dim lngLastLine As Long
Dim lngLLoop As Long
Dim lngTmp As Long
Dim lngCounter As Long

    On Error GoTo ErrHan:
    
    lngLUBnd = UBound(pLines)
    lngPUBnd = UBound(pPositions)
    
    ReDim pLineNumbers(lngPUBnd)
    lngLastLine = 1             ' ... presume item 0 is 1.
    
    For lngPLoop = 0 To lngPUBnd
    
        lngTmpPos = pPositions(lngPLoop)
        
        For lngLLoop = lngLastLine To lngLUBnd - 1
            lngTmpLine = pLines(lngLLoop)
            lngTmp = lngTmpLine - pDelimiterLength
            If lngTmpPos <= lngTmp Then
                pLineNumbers(lngCounter) = lngLLoop - 1
                lngCounter = lngCounter + 1
                Exit For
            End If
        Next lngLLoop
    
    Next lngPLoop
    
        
Exit Sub
ErrHan:


End Sub

Public Sub ReadTextLines(ByRef pTheText As String, ByRef pTheLinesArray() As Long, ByRef pDelimiterUsed As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ReadTextLines.VB_Description = "This method takes a String and proceeds to read its lines returning the text of the file, and an array of line start positions."

' ... This method takes some text and proceeds to read its lines (presuming vbCrLf or Lf line delimiter)
' ... and returns an array of line start positions (0 = 1 to Len(File Text) = last char pos).

' ... The idea is to help parsing the text as lines afterwards using the arrays and the text contents e.g.
' ... Mid$ into the string with start amd length derived from array of line start positions,
' ... in so doing , tries to obviate the need to split a string into string array to read it in lines.

Dim bOK As Boolean
Dim sErrMsg As String
Dim lng13Found As Long
Dim lng10Found As Long
Dim sDel As String

    On Error GoTo ErrReading:
    
    If mReadingText = True Then
        ' ... avoid re-entrancy.
        Err.Raise vbObjectError + 1001, , cMsgInUse
    End If
    
    On Error GoTo ErrHan:
    
    mReadingText = True
        
    ' ... check if lines delimited by CRLF or just LF.
    lng10Found = InStr(1, pTheText, vbLf, vbBinaryCompare)
    
    lng13Found = InStr(1, pTheText, vbCr, vbBinaryCompare)
    
    ' ... default code file delimiter.
    sDel = vbCrLf
    
    bOK = lng10Found > 0 Or lng13Found > 0
    
    If (lng13Found + 1) <> lng10Found Then
        
        Let bOK = lng10Found > 0
        
        If bOK = True Then
            
            sDel = vbLf
        
        Else
            
            bOK = Len(pTheText) > 0
            
            If bOK = False Then
                
                Err.Raise vbObjectError + 1001, , "The program could not split the Text into Lines : No Line Delimiter Found."
            
            End If
        
        End If
    
    
    End If
    
    pDelimiterUsed = sDel
    
    If bOK = True Then
        ' ... try and get an array of line starting positions.
        modReader.DeriveLinesArray pTheText, pTheLinesArray, sDel, bOK, sErrMsg
        
    End If
        
ResumeErr:

    ' ... return stuff.
    pOK = bOK
    
    pErrMsg = sErrMsg
    ' ... clean-up.
    sErrMsg = vbNullString
    
    mReadingText = False
    
Exit Sub
ErrHan:
    
    bOK = False
    
    sErrMsg = Err.Description
    
    Debug.Print "Error.modReader.ReadTextLines: " & Err.Description
    
    GoTo ResumeErr:
    
ErrReading:
    
    Debug.Print "Error.modReader.ReadTextLines: Busy... " & Err.Description
    
End Sub ' ... ReadTextLines


Public Sub ReadTextFile(ByVal pTheFileName As String, ByRef pTheFileText As String, ByRef pTheLinesArray() As Long, ByRef pDelimiterUsed As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ReadTextFile.VB_Description = "This method takes a filename and proceeds to read its lines returning the text of the file, and an array of line start positions."

' ... This method takes a filename and proceeds to read its lines (presuming vbCrLf or Lf line delimiter)
' ... and returns the text of the file, and an array of line start positions ( valid return is 1 to Len(File Text)).
' ... Note: if upper bound line array returns 0 and linearray(0) = -1 then no processing occurred.

' ... The idea is to help parsing the text afterwards in lines.

Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
'Dim bFileExists As Boolean
Dim lngFileLength As Long
Dim sTmp As String

Dim bOK As Boolean
Dim sErrMsg As String

    On Error GoTo ErrReading:
    
    If mReadingFile = True Then
        ' ... avoid re-entrancy.
        Err.Raise vbObjectError + 1000, , cMsgInUse
    End If
    
    On Error GoTo ErrHan:
    
    mReadingFile = True
    
    ' ... set up the return array with -1 at index 0
    ReDim pTheLinesArray(0 To 0)
    pTheLinesArray(0) = -1
    
    sTmp = modReader.ReadFile(pTheFileName, bOK, sErrMsg)
    
    If bOK = True Then
    
        modReader.ReadTextLines sTmp, pTheLinesArray, pDelimiterUsed, bOK, sErrMsg
        
    End If
    
ResumeErr:

    ' ... return stuff.
    pTheFileText = sTmp
    
    pOK = bOK
    pErrMsg = sErrMsg
    ' ... clean-up.
    sTmp = vbNullString: sErrMsg = vbNullString
    lngFileLength = 0
    
    On Error GoTo 0
    If bFileIsOpen Then
        ' ... close the text file.
        Close #iFileNumber
    End If
    
    mReadingFile = False
    
Exit Sub
ErrHan:
    
    bOK = False
    
    sErrMsg = Err.Description
    
    Debug.Print "Error.modReader.ReadTextFile: " & Err.Description
    
    GoTo ResumeErr:
    
ErrReading:
    
    Debug.Print "Error.modReader.ReadTextFile: Busy... " & Err.Description
    
End Sub ' ... ReadTextFile

Public Function ReadFile(ByRef pTheFileName As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As String
Attribute ReadFile.VB_Description = "Returns the contents of a file as a string."

'... Parameters.
'    R__ pTheFileName: String        ' ... The name of the file to read.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim strReturn As String              ' ... a return value to this function.
Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
Dim bFileExists As Boolean
Dim lngLength As Long
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    bFileExists = FileLen(pTheFileName) > 0
    
    If bFileExists = True Then
        
        ' -------------------------------------------------------------------
        ' ... get the next free file number.
        iFileNumber = FreeFile()
        ' -------------------------------------------------------------------
        Open pTheFileName For Binary As #iFileNumber
        ' -------------------------------------------------------------------
        ' ... if execution flow got here, the file has been open without error.
        bFileIsOpen = True
        
        lngLength = LOF(iFileNumber)
        strReturn = String$(lngLength, Chr$(0))
        
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
        ReadFile = strReturn            ' ... return the string from the text file.
        pOK = True                      ' ... return success.
        ' -------------------------------------------------------------------
    
    End If
    
ErrResume:

    On Error Resume Next
    ' -------------------------------------------------------------------
    ' ... try closing file silently to error.
    If bFileIsOpen = True Then
        Close #iFileNumber
        If Err.Number <> 0 Then
            Debug.Print "modReader.ReadFile (Close File)", Err.Number, Err.Description
            Err.Clear
        End If
    End If
    ' -------------------------------------------------------------------
    strReturn = vbNullString
    lngLength = &H0
    iFileNumber = &H0
    
Exit Function
ErrHan:

    pErrMsg = Err.Description
    pOK = False
    Debug.Print "modReader.ReadFile", Err.Number, Err.Description
    
    Err.Clear
    
    Resume ErrResume:
    
End Function ' ... ReadFile: String.

Public Sub DeriveLinesArray(TheText As String, ByRef pTheLines() As Long, Optional ByVal pTheDelimiter As String = vbCrLf, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute DeriveLinesArray.VB_Description = "read a string as lines of text and return an array of line starting positions."

Dim lngFound As Long
Dim lngDelCount As Long
Dim lngStart As Long
Dim lngDelLength As Long
Dim lngTextLength As Long
Dim sDelimiter As String
Dim c As Long
Dim tmp As Long

' -------------------------------------------------------------------
' ... Helper: read a string as lines of text and return an array of line starting positions.
' -------------------------------------------------------------------
    
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    ' ... pre-process validation and set up.
    If mMaking Then
        Err.Raise vbObjectError + 1002, "", cMsgInUse
    End If
    
    pOK = False ' ... set return to false in case it comes in as true to start with.
    mMaking = True
    
    sDelimiter = pTheDelimiter

    lngDelLength = Len(sDelimiter)
    lngTextLength = Len(TheText)
    
    If lngTextLength = 0 Or lngDelLength = 0 Then
        Err.Raise vbObjectError + 1002, , "Parse For Lines Array: not a valid string or delimiter length."
    End If
    
    ' -------------------------------------------------------------------
    ' ... begin processing.
    ReDim Preserve pTheLines(0 To lngTextLength + 2)
    
    pTheLines(0) = 1
    c = 1
    
    tmp = InStr(TheText, sDelimiter)

    Do While tmp
        pTheLines(c) = tmp + lngDelLength
        tmp = InStr(pTheLines(c), TheText, sDelimiter)
        c = c + 1
    Loop
    
    pTheLines(c + 1) = lngTextLength ' ... ubound(lines) - 1 is last line in text while ubound(lines) is total length of text.
    
    ' -------------------------------------------------------------------
    ReDim Preserve pTheLines(c + 1)
    ' ... end processing
    ' -------------------------------------------------------------------
    
    pOK = True
    
ResumeError:

    mMaking = False
    
Exit Sub
ErrHan:
    pOK = False
    pErrMsg = Err.Description
    ReDim pTheLines(0)
    Debug.Print "Error.modReader.DeriveLinesArray: " & Err.Description
    Resume ResumeError:

ErrMaking:
    Debug.Print "Error.modReader.DeriveLinesArray: Busy... " & Err.Description
End Sub      ' ... DeriveLinesArray.

'**************************************
' next two functions from psc written by dogan.

' Name: HyperFast! Read/Write File Functions
' Description:These two functions are designed to read and write a file as fast as possible in VB. It is faster for some cases than WinApi Read/WriteFile functions because you don't have to convert binary to string in a loop. Thats why, it is very fast and useful for any purpose. I have created it for my Encryption programme and with Windows Crypto functions and this two functions, my encrytion programme works faster than most of the Encrytion software you can download on the net. Hope it will be useful for you. Thanks.
' Ozan Yasin Dogan
' www.uni-group.org (will be online in 01/06/02)
' By: tektus

Public Function ReadFileX(FileName As String) As String 'Returns STRING variable!
Attribute ReadFileX.VB_Description = "A super fast file reading method c/o Dogan @ PSC."

Const Buf As Integer = 30000
'Declarations
Dim FileLen As Long 'To keep file lenght information
Dim Multiply As Long 'It is required to find how many Buf
'bytes exist in the file. For ex: in a 125,000bytes file
'there are 4 multiply. The rest is recorded to Plus variable
Dim Temp As String * Buf 'Temporary string block
'It is necessary for use of Random Access methode.
'If not, you had to open it in Binary mode and convert
'binary data to text, and it is also a loop and slows
'down the process. This is the best methode i think..
Dim Content As String 'Content is the file content,
'the function allocates a space for it first and
'full it with Mid function. It is a very fast methode
'instead of using Content = Content & Something
Dim Plus As Long 'The plus part of the file after dividing
'to Buf variable. It is used when the file lenght is small
'than Buf and to find the rest of the bytes after dividing
'file lenght to Buf
Dim Point As Long 'Point shows on which byte the content is.
Dim FileNo As Byte 'To find a free file number
Dim Counter As Long 'Is required for loops

    FileNo = FreeFile 'Find a free file number
    
    Open FileName For Random As #FileNo Len = Buf 'Open the file as Random, each record will have the lenght of Buf
    FileLen = LOF(FileNo) 'File lenght
    Multiply = Int(FileLen \ Buf) 'How many loops required to read the file
    Content = Space(FileLen) 'Allocate a space for file content in the memory
    Plus = FileLen - (Multiply * Buf) 'After this loops, there might be also some bytes to read
    Point = 1 'Content is in this byte: 1
    
    If Multiply = 0 Then 'If the file is smaller than Buf (30000 bytes here, you can change it)
        Plus = FileLen: Counter = 1: GoTo Jump1
    End If

    'This loop reads the file as it was defined in a Type,
    'using random access methode and adds each records
    'to the content using Mid function.
    'Because Content = Content & Temp would slow down
    'the loop very much! And as you see, there is no transfer
    'beetween binary to string..
    
    For Counter = 1 To Multiply
        Get #FileNo, Counter, Temp
        Mid(Content, Point, Buf) = Temp
        Point = Point + Buf
    Next Counter
    
Jump1:
    
    'This is for the rest of the file after the loop.
    If Plus > 0 Then
        Get #FileNo, Counter, Temp
        Mid(Content, Point, Plus) = Left(Temp, Plus)
    End If
    
    Close #FileNo
    ReadFileX = Content
    
End Function

Public Sub WriteFile(FileName As String, Content As String)
Attribute WriteFile.VB_Description = "Write a string to a file (here only for reference)."

Dim FileNo As Byte 'To find a free file number

    FileNo = FreeFile
    Open FileName For Output As #FileNo
    Print #FileNo, Content; '; is required for Vb to not write another 2 charachters of new line in the file
    Close #FileNo
    
End Sub
