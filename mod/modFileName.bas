Attribute VB_Name = "modFileName"
Attribute VB_Description = "A module to provide functions to help access the various elements of a file name including a VB specific method to derive a full path and file name of files found in Project / Group files (VBP/VBG)."

' what?
'  a couple of methods to help handling file names with
'  a public stucture to boot.
' why?
'  just to simplify extracting the various attributes of a file name
'  in the heat of battle :)
' when?
'  when you have to read and manage file names somehow.
' how?
'  see methods for examples.
' who?
'  d.c.

Option Explicit

'== Subs: 2
'+  ParseFileNameEx:              "Parses a string representing the full path and file name of a file into a FileNameInfo Type."
'+  ParseFileName:                "Parses a string representing the full path and name of a file into File Name, Extension and Path,"
'
'== Functions: 1
'+  CheckValidName: Long          "Looks for invalid chars in a file name and returns the ainsi value of the first found, -1 if errored or 0 if file name is deemed valid."

' Requires:
' modStrings: InStrRevChar

Public Type FileNameInfo
    Extension As String
    FileName As String
    File As String
    Path As String
    PathAndName As String
End Type

Function FileExists(ByVal sFileName As String) As Boolean
' returns True when drive / folder / file exists:

Dim i As Long
    
    On Error GoTo NotFound
    
    i = GetAttr(sFileName)
        
    FileExists = True
    
Exit Function
NotFound:
    FileExists = False
    
End Function

Public Function TestFileName(ByRef pFileName As String, _
                    Optional ByRef pBadCharNum As Long = -1, _
                    Optional ByRef pBadCharPos As Long = -1, _
                    Optional ByRef pBadCharMsg As String = vbNullString) As Boolean
Attribute TestFileName.VB_Description = "Tests a file name for invalid characters; if found, returns false, the invalid char num, its position in the file name and a message in out params."

' Tests a file name for invalid characters; if found, returns false, the invalid char num, its position in the file name and a message in out params.
' Returns True if file name OK

' according to my english xp, the following characters are not allowed in the name of a file
' /\:*?"<>|
' 92,47,58,42,63,34,60,62,124


Dim x() As Byte

Dim l As Long
Dim i As Long
Dim c As Long

    ' ignoring dbcs or unicode, stepping thru' byte pairs
    
    On Error Resume Next
    
    If Len(pFileName) = 0 Then
        pBadCharMsg = "No Name to Test"
        Exit Function
    End If
    
    x = pFileName
    l = UBound(x)
    For i = 0 To l Step 2
        c = x(i)
        Select Case c
            Case 34, 42, 47, 58, 60, 62, 63, 92, 124
                pBadCharNum = c
                pBadCharPos = i / 2 + 1
                pBadCharMsg = "Bad File Name:" & vbNewLine & Chr$(c) & " is not allowed in the name of a file."
                GoTo Finish:
        End Select
    Next i
    TestFileName = True
    
Finish:

    If Len(pBadCharMsg) Then Debug.Print pBadCharMsg
    Erase x
    i = 0: l = 0: c = 0
    
End Function

Public Function CheckValidName(pTheFileName As String, pIncludesPath As Boolean) As Long
Attribute CheckValidName.VB_Description = "Looks for invalid chars in a file name and returns the ainsi value of the first found, -1 if errored or 0 if file name is deemed valid."

' ... checks to see if a file name includes invalid chars because
' ... a file write will fail if an illegal char exists in the name.
' ... this function returns the following
'   ... -1  ... an error, invalid request.
'   ... >0  ... the char number of an illegal character found in the name.
'   ...  0  ... if all went well and no invalid chars exist in the file name.

' ... assumptions:
'   ... either a file name or a full path and file name are passed.
'   ... only the file name is being tested.
' ... example:
' ...  test a file name (including full path) for invalid chars.
' ...       lngValidName = CheckValidName(sFileName, True)
' ...  test the result...
' ...       If lngValidName <> 0 Then
' ...           If lngValidName > 0 Then
' ...               sErrMsg = "The Char " & lngValidChar & " is not allowed."
' ...           ElseIf lngValidName < 0 Then
' ...               sErrMsg = "Unable to process request."
' ...           End If
' ...       End If

Dim xInfo As FileNameInfo
Dim lngLoop As Long
Dim lngNLen As Long
Dim sTmpChar As String
Dim sTmpName As String
Dim lngFound As Long
Dim sTmpInvalid As String
Const cInvalidFileNameChars As String = "`!$%^&*()-=+[]{}'#@~;,.<>/?\| " ' ... needs " as well.
    
    CheckValidName = -1
    
    If pIncludesPath Then
        ' ... make sure we have the name of the file.
        ParseFileNameEx pTheFileName, xInfo
        sTmpName = xInfo.FileName
    Else
        ' ... note: only want to check the file name itself
        ' ... ignoring the extension which includes an invalid name char " . ".
        sTmpName = pTheFileName
        lngFound = InStr(1, sTmpName, ".")
        If lngFound > 0 Then
            sTmpName = Left$(sTmpName, lngFound - 1)
        End If
    End If
    
    lngNLen = Len(sTmpName)
    
    If lngNLen > 0 Then
        sTmpInvalid = cInvalidFileNameChars & Chr$(34)
        For lngLoop = 1 To lngNLen
            sTmpChar = Mid$(sTmpName, lngLoop, 1)
            lngFound = InStr(1, sTmpInvalid, sTmpChar)
            If lngFound > 0 Then
                CheckValidName = Asc(sTmpChar)
                Exit For
            End If
        Next lngLoop
    End If
    
    ' ... if we got here and CheckValidName is still -1 then
    ' ... assume no error, things were valid and return 0 for not found.
    If CheckValidName = -1 Then CheckValidName = 0

End Function

Sub ParseFileNameEx(ByVal pTheFullName As String, ByRef pFileNameInfo As FileNameInfo)
Attribute ParseFileNameEx.VB_Description = "Parses a string representing the full path and file name of a file into a FileNameInfo Type."

' ... simplified wrapper to ParseFileNmae using Type: FileNameInfo for easier client coding (less declarations).

' ... example:
'   Dim x As FileNameInfo
'   Dim s As String
'       s = "C;\Stuff\Test.dat
'       ParseFileNameEx s, x
'       With x
'           Print "File Name: " & .FileName                     ' ... Test
'           Print "Extension: " & .Extension                    ' ... dat
'           Print "Path: " & .Path                              ' ... C:\Stuff
'           Print "Full Path And File Name: " & .PathAndName    ' ... C:\Stuff\Test.dat
'       End With

Dim sTheFileName As String
Dim sTheFile As String
Dim sThePath As String
Dim sTheExtension As String

    'ParseFileName pTheFullName, sTheFileName, sTheFile, sThePath, sTheExtension
    ShredFileName pTheFullName, sTheFileName, sTheFile, sThePath, sTheExtension
    
    With pFileNameInfo
    
        .FileName = sTheFileName
        .File = sTheFile
        .Extension = sTheExtension
        .Path = sThePath
        .PathAndName = pTheFullName
        
    End With

End Sub ' ... ParseFileNameEx.

Sub ParseFileName(ByVal pTheFullName As String, _
                  ByRef pTheFileName As String, _
         Optional ByRef pTheFile As String = vbNullString, _
         Optional ByRef pThePath As String = vbNullString, _
         Optional ByRef pTheExt As String = vbNullString)
Attribute ParseFileName.VB_Description = "Parses a string representing the full path and name of a file into File Name, Extension and Path,"

' ... Parameters:
' ... In:   pTheFullName            The full path and name of the file  e.g. "C:\Stuff\Test.dat"
' ... Out:  pTheFileName            The name of the file                e.g. "Test"
' ... Out:  pTheFile                The full name of the file           e.g. "Test.dat"
' ... Out:  pThePath                The full path of the file           e.g. "C:\Stuff"
' ... Out:  pTheExt                 The file's extension                e.g. "dat"


' Notes:    This converts the full file name into a Byte Array and then
'           loops from the end of the array (UBound - 1) to the start of the array (pos 0).
'           It tests for the (first) period for the extension.
'           It looks for the last back-slash/forward-slash (whatever?)
'           The extension may have more than one period e.g. C:\Text.exe.manifest in which case exe.manifest will be returned as the extension.
'           It removes the last ' \ ' from the path so that "C:\Test.dat" will return C: as the path rather than C:\

' IT'S GOT A BUG, IT CAN'T GET THIS FILENAME CORRECT, ITS A VBP NOT A TXT
' ... C:\codevb\4psc\codeviewer\unzipfiles\urdu_text_2206796222011\Urdu Text Handling As Plane Text In Text File(.TXT).vbp

' ... defer to ShredFileName

    ShredFileName pTheFullName, pTheFileName, pTheFile, pThePath, pTheExt
    
    Exit Sub
    
Dim lngLen As Long          ' ... length of the input full file name.
Dim sBytes() As Byte        ' ... the byte array derived from the full file name.
Dim uBnd As Long            ' ... upper limit of the byte array.
Dim lBnd As Long            ' ... lower limit of the byte array (added just in case... ).
Dim lngCharPos As Long      ' ... index of the current character being read in the loop.
Dim lngChar As Long         ' ... the current character being read.
Dim lngLoop As Long         ' ... the loop thingy(?) [self schooled no learny this].
Dim lngExtStart As Long     ' ... the start of the last extension marker found (which is the first one when read left to right).
Dim lngFNStart As Long      ' ... start of the file name proper.
Dim sTmpString As String    ' ... copy of the input string to avoid re-referencing the parameter.

    lngLen = Len(pTheFullName)
    
    If lngLen > 0 Then                              ' ... make sure there's something to parse.
        
        sTmpString = pTheFullName
        sBytes = sTmpString                       ' ... grab the bytes.
        lBnd = 0
        uBnd = UBound(sBytes) - 1                   ' ... minus 1 due to byte pair e.g. (x,0) else will always read 0.
        lngCharPos = lngLen                         ' ... start of current char goes from end to beginning.
        
        For lngLoop = uBnd To lBnd Step -2          ' ... go from the last char byte to the first by steps of 2.
            
            lngChar = sBytes(lngLoop)               ' ... read the character as a number.
            
            Select Case lngChar
                Case 46                             ' ... the period char for the extension e.g. ' . '
                    lngExtStart = lngCharPos + 1    ' ... reading backwards so if don't add 1 extension will include the period, e.g. ".dat" rather than "dat".
                Case 58, 92                         ' ... ' : ' and ' \ ' respectively, drive or directory delimiter.
                    lngFNStart = lngCharPos         ' ... read the start of the file name proper.
                    Exit For                        ' ... should have the start of the file name so quit else will mis-read file name.
                                                    ' ... e.g Stuff\Test.dat rather than Test.dat.
            End Select
            
            lngCharPos = lngCharPos - 1             ' ... decrement the current char pos.
            
        Next lngLoop
        
        ' ... reckon we done the processing, so now mid into the full file name to extract the bits we want (or at least try).
        If lngExtStart > 0 Then
            ' ... Mid no likey 0 start so make sure to have a valid start in this respect
            ' ... before attempting to extract the extension.
            pTheExt = Mid$(sTmpString, lngExtStart)
        End If
        
        If lngFNStart > 0 Then
            ' ... see above and extract the full name of the file.
            pTheFile = Mid$(sTmpString, lngFNStart + 1)
            ' ... see whether there is an extension and remove it for the file name without the extension.
            pTheFileName = pTheFile
            If lngExtStart > 0 Then
                ' ... remove the extension from the file name member: note: calc length to remove by subtracting ext start + 2 from length of input).
                pTheFileName = Left$(pTheFile, Len(pTheFileName) - (lngLen - lngExtStart + 2))
            End If
            pThePath = Mid$(sTmpString, 1, lngFNStart - 1)
        
        End If
        
        Erase sBytes
        sTmpString = vbNullString
    
    End If
    
End Sub ' ... ParseFileName.


Sub ShredFileName(ByVal pTheFullName As String, _
         Optional ByRef pTheFileName As String, _
         Optional ByRef pTheFile As String = vbNullString, _
         Optional ByRef pThePath As String = vbNullString, _
         Optional ByRef pTheExt As String = vbNullString)

' ... Parameters:
' ... In:   pTheFullName            The full path and name of the file  e.g. "C:\Stuff\Test.dat"
' ... Out:  pTheFileName            The name of the file                e.g. "Test"
' ... Out:  pTheFile                The full name of the file           e.g. "Test.dat"
' ... Out:  pThePath                The full path of the file           e.g. "C:\Stuff"
' ... Out:  pTheExt                 The file's extension                e.g. "dat"

' ... Notes:
'     ParseFileName doesn't work correctly, this is its replacement.
'     This version doesn't try so hard, uses less vars and should be faster
'     because there is less string setting stuff going on.
'     Of course, it's designed to work on valid file names;
'     if it doesn't work then the in/out params will remain unchanged from how they entered.

' -------------------------------------------------------------------
' ... This is the file name that killed ParseFileName
' ... C:\codevb\4psc\codeviewer\unzipfiles\urdu_text_2206796222011\Urdu Text Handling As Plane Text In Text File(.TXT).vbp
' ... ShredFileName "C:\codevb\4psc\codeviewer\unzipfiles\urdu_text_2206796222011\Urdu Text Handling As Plane Text In Text File(.TXT).vbp"
' -------------------------------------------------------------------

Dim x() As Byte             ' ... bytes of the full file name
Dim l As Long               ' ... upper bound of x bytes
Dim p As Long               ' ... last position of \
Dim i As Long               ' ... x byte loop
Dim j As Long               ' ... char number of x(i)
    
    On Error Resume Next
    
    ' ... make sure there's something to parse.
    If Len(pTheFullName) Then
    
        ' -------------------------------------------------------------------
        x = pTheFullName
        l = UBound(x)
        ' ... get the last folder delimiter ' \ ' position, start of full file name > forward
        For i = 0 To l Step 2
            j = x(i)
            If j = 92 Then
                p = i / 2 + 1
            End If
        Next i
        
        If p Then
            ' ... extract the file and the file path
            pTheFile = Mid$(pTheFullName, p + 1)
            pThePath = Left$(pTheFullName, p - 1)
            ' -------------------------------------------------------------------
ResFileNameOnly:

            p = 0
            x = pTheFile
            l = UBound(x) - 1   ' ... sync to leading byte
            ' ... look for an extension, identified by a period ' . ', end of file < backward
            For i = l To 0 Step -2
                j = x(i)
                If j = 46 Then
                    p = i / 2 + 1
                    Exit For
                End If
            Next i
            ' -------------------------------------------------------------------
            If p Then
                ' ... extract the extension and the name (only) of the file.
                pTheExt = Mid$(pTheFile, p + 1)
                pTheFileName = Left$(pTheFile, p - 1)
            End If
        
        Else
            ' ... just a file name? no path? / after thought/fix
            pTheFile = pTheFullName
            GoTo ResFileNameOnly:
            
        End If
        
'        Debug.Print pTheExt & vbNewLine & pTheFileName & vbNewLine & pTheFile & vbNewLine & pThePath & vbNewLine & pTheFullName
        
    End If
    
    Erase x
    p = 0: i = 0: j = 0: l = 0
    
End Sub ' ... ShredFileName.


