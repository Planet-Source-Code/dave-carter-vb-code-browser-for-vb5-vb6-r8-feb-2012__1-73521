Attribute VB_Name = "modUnZip"
Attribute VB_Description = "A module to help utilise Info-Zip's UnZip32.dll for reading and unzipping compressed files."
' what?
'  a module to read and unzip zip files.
' why?
'  to help program read and unzip zipped projects from psc and elsewhere.
' when?
'  when one needs to check or unzip the contents of a zip file.
' how?
'   Example:
'   Read the file names in a zip file:
'
'   ... pTheZipFileName is passed to us...
'   Dim sZipFileNames As String
'       ' ... read the zip file.
'       modUnZip.ReadZip pTheZipFileName, sZipFileNames
'
'   ' ... sZipFileNames returns a string, row del: vbCrLf, col. del: |, with the
'   ' ... names and info for each file member in the zip.
'   File Name | File Folder | Full Member Name | Date | Uncomp. Size | Comp. Size | Zip Index
'   Date (above) is written as a double to avoid parsing a date string on the receiving end if required.

Option Explicit

' Notes:
'   I used various sources to learn what I had to do here and experimented a little
'   to see what I could get away with to just do the minimum that I wanted.
'   Sources:

'   ' Unzipping files using the free Info-Zip Unzip DLL with VB ' @
'   http://www.vbaccelerator.com/home/vb/code/libraries/compression/unzipping_files/article.asp

'   Zip and Unzip Using VB5 or VB6, Chris Eastwood.
'   http://www.codeguru.com/vb/gen/vb_graphics/fileformats/article.php/c6743

'   ZipSearch, Rde. (no unzip to memory without this, see method, Thank you Rohan)
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72288&lngWId=1

'   Zipping with Info-Zip, Jox.
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=70968&lngWId=1

'   Technical information @
'   http://www.opensource.apple.com/source/zip/zip-6/unzip/unzip/windll/windll.txt?txt
'   and finally, the creators of this amazing compression program.
'   http://www.info-zip.org

'   There are two easy ways to get the required dlls;
'   follow the links to vbAccelerator or CodeGuru (preferable) above:
'   If you use CodeGuru, check for the skip add button top right and click that first.
'   If you use vbAccelerator then rename the files downloaded to UnZip32.dll & ... .
'   (See the API declarations below, the name for Lib is the name of the dll, minus the extemsion.)
'   Once you have the files, copy them to Windows\System32.
'   The dlls used by this source are older versions than currently available
'   but they still seem to work.

'   I'm ok about downloading stuff from codeguru and vbaccelerator,
'   but if you are reluctant to download the files then the worst is
'   that this program will not support unzipping, that's it! good luck deciding.

' Special Note:
'   Rde's ZipSearch was the last project I studied.
'   He makes light work of extracting value from the Unzip32 library.
'   Including Declarations he wraps it all up in 214 lines.
'
'   The other projects researched focus more on providing a full interface
'   to unzipping, but I was after the bare bones of it all

Private Type DCLIST
    ExtractOnlyNewer As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nfflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    nPrivilege As Long
    szZipName As String
    szExtractDir As String
End Type

Private Type USERFUNCTION
    lptrPrnt As Long
    lptrSound As Long
    lptrReplace As Long
    lptrPassword As Long
    lptrMessage As Long
    lptrService As Long
    lTotalSizeComp As Long
    lTotalSize As Long
    lCompFactor As Long
    lNumMembers As Long
    cchComment As Integer
End Type

Private Type CBChar
    ch(0 To 32800) As Byte
End Type

Private Type CBCh
    ch(0 To 255) As Byte
End Type

Private Type ZipNames
    zFiles(1024) As String
End Type

Private mZMemIndex As Long      ' ... set this to 0 when starting an unzip op.
Private msZipFiles As String    ' ... captures the names of files processed in UnzipMessageCallBack.
Private mbUnipToMem As Boolean  ' ... flag to say just unzipping to memory so no need to
                                ' ... do extra processing in ServiceCallBack.

Private Type UzpBuffer
    strLength As Long ' length of string
    strPointer As Long ' pointer to string
End Type

Private Declare Function Wiz_SingleEntryUnzip Lib "Unzip32" (ByVal ifnc As Long, _
                                                             ByRef ifnv As ZipNames, _
                                                             ByVal xfnc As Long, _
                                                             ByRef xfnv As ZipNames, _
                                                             ByRef lpDCL As DCLIST, _
                                                             ByRef lpUserFuncs As USERFUNCTION) _
                                                             As Long
                                                             
Private Declare Function Wiz_UnzipToMemory Lib "Unzip32" (ByVal lpszZip As String, _
                                                          ByVal lpszFile As String, _
                                                          ByRef lpUserFuncs As USERFUNCTION, _
                                                          ByRef lpRetStr As UzpBuffer) As Long

Private Declare Sub UzpFreeMemBuffer Lib "Unzip32" (ByRef lpRestr As UzpBuffer)
'void WINAPI UzpFreeMemBuffer(UzpBuffer *retstr)
'   Use this routine to release the return data space allocated by the function
'   Wiz_UnzipToMemory().

Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lLenB As Long)
Attribute CopyMemByV.VB_Description = "RTLMoveMemory API"

Private Function pAddressOf(ByVal lPtr As Long) As Long
Attribute pAddressOf.VB_Description = "work around in vb to get a pointer to the address of a method."
   ' VB Bug workaround fn
   pAddressOf = lPtr
End Function

Public Function UnzipToMemory(ByVal pvTheZipFileToRead As String, _
                              ByVal pvTheZipMember As String, _
                              ByRef prTheMemberText As String) As Boolean
Attribute UnzipToMemory.VB_Description = "Method to unzip a file to memory so we have its string value data."

' ... Decompress/Unzip a file, in the zip, to memory, as a string.

Dim tZipCBPointers As USERFUNCTION
Dim sZipFile As String
Dim sZipMember As String
Dim sTmp As String
Dim lngRet As Long
Dim lngLoop As Long
Dim lngZLength As Long
Dim lngZMLength As Long
Dim sRetString As UzpBuffer
'Dim lngRetLength As Long
Dim bRetString() As Byte
'Dim lngChar As Long
'Dim bb() As Byte
'Dim ss As String

    On Error GoTo ErrHan:
    
    mbUnipToMem = True
    
    lngZLength = Len(pvTheZipFileToRead)
    If lngZLength = 0 Then Err.Raise vbObjectError + 1000, , "No Zip File Name to Read."
    
    lngZMLength = Len(pvTheZipMember)
    If lngZMLength = 0 Then Err.Raise vbObjectError + 1000, , "No Zip Member File Name to Read."
    
    For lngLoop = 1 To lngZLength
        sTmp = Mid$(pvTheZipFileToRead, lngLoop, 1)
        If sTmp = "\" Then sTmp = "/"
        sZipFile = sZipFile & sTmp
    Next lngLoop
    
    For lngLoop = 1 To lngZMLength
        sTmp = Mid$(pvTheZipMember, lngLoop, 1)
        If sTmp = "\" Then sTmp = "/"
        sZipMember = sZipMember & sTmp
    Next lngLoop
    
    tZipCBPointers.lptrPrnt = pAddressOf(AddressOf UnzipPrintCallback)
    tZipCBPointers.lptrReplace = pAddressOf(AddressOf UnzipReplaceCallback)
    tZipCBPointers.lptrPassword = pAddressOf(AddressOf UnzipPasswordCallBack)
    tZipCBPointers.lptrMessage = pAddressOf(AddressOf UnzipMessageCallBack)
    tZipCBPointers.lptrService = pAddressOf(AddressOf UnZipServiceCallback)
    
    lngRet = Wiz_UnzipToMemory(sZipFile, sZipMember, tZipCBPointers, sRetString)
    
    UnzipToMemory = lngRet <> 0 ' ... zero means fail.
    
    If UnzipToMemory Then
        
        ' ... Credit to Rde (Happy Coding :) )
        ' ... I couldn't figure out how to get the memory chars into a string...
        ' ... as ever, Rde, came to the rescue, see ZipSearch, it's fab.
        
        ReDim bRetString(sRetString.strLength - 1&) As Byte
        CopyMemByV VarPtr(bRetString(0&)), sRetString.strPointer, sRetString.strLength
    
        prTheMemberText = StrConv(bRetString, vbUnicode)
                
    End If
    
ResumeError:
    
    mbUnipToMem = False
    
    On Error GoTo 0 ' ... just in case we get an error clearing mem.
    ' ... clear the unzip to mem. buffer.
    UzpFreeMemBuffer sRetString
    
Exit Function

ErrHan:

    Debug.Print "modUnZip.UnzipToMemory.Error: " & Err.Number & "; " & Err.Description
    Err.Clear
    Resume ResumeError:

End Function

Public Function ReadZip(ByVal pvTheZipFileToRead As String, _
               Optional ByRef prTheZipMemberString As String = vbNullString, _
               Optional ByRef prTheMemberCount As Long = 0) As Boolean
Attribute ReadZip.VB_Description = "Method to Read Contents of Zip file and built delimited string of file names & info."

Dim tZipCBPointers As USERFUNCTION
Dim tDCList As DCLIST
Dim lngLoop As Long
Dim lngLength As Long
Dim sZipFile As String
Dim sTmp As String
Dim lngRet As Long
Dim uZipNames As ZipNames
Dim uExcNames As ZipNames

    
    On Error GoTo ErrHan:
    
    mbUnipToMem = False
    
    prTheZipMemberString = vbNullString
    prTheMemberCount = 0
    
    mZMemIndex = 0
    msZipFiles = vbNullString
    
    lngLength = Len(pvTheZipFileToRead)
    If lngLength = 0 Then Err.Raise vbObjectError + 1000, , "No File Name to Read."
            
    tZipCBPointers.lptrPrnt = pAddressOf(AddressOf UnzipPrintCallback)
    tZipCBPointers.lptrReplace = pAddressOf(AddressOf UnzipReplaceCallback)
    tZipCBPointers.lptrPassword = pAddressOf(AddressOf UnzipPasswordCallBack)
    tZipCBPointers.lptrMessage = pAddressOf(AddressOf UnzipMessageCallBack)
    tZipCBPointers.lptrService = pAddressOf(AddressOf UnZipServiceCallback)
    
    For lngLoop = 1 To lngLength
        sTmp = Mid$(pvTheZipFileToRead, lngLoop, 1)
        If sTmp = "\" Then sTmp = "/"
        sZipFile = sZipFile & sTmp
    Next lngLoop
    
    tDCList.szZipName = sZipFile
    tDCList.nvflag = 1
    
    lngRet = Wiz_SingleEntryUnzip(0, uZipNames, 0, uExcNames, tDCList, tZipCBPointers)
    
    prTheZipMemberString = msZipFiles   ' ... return the contents string.
    prTheMemberCount = mZMemIndex       ' ... return the contents count.
    
    ReadZip = lngRet = 0
    
Exit Function
ErrHan:
    MsgBox "The program could not Read the Zip File and gave the following reason:" & vbNewLine & Err.Description, vbExclamation, "Read Zip"
    Debug.Print "modUnZip.ReadZip.Error: " & Err.Description
    
End Function

Public Function Unzip(pTheFileToUnzip As String, _
                      pTheUnzipFolder As String, _
       Optional ByRef prTheZipMemberString As String = vbNullString, _
       Optional ByRef prTheMemberCount As Long = 0) As Long
Attribute Unzip.VB_Description = "Method to Unzip a zip file to a folder."
                             
Dim tZipCBPointers As USERFUNCTION
Dim tDCList As DCLIST
Dim lngLoop As Long
Dim lngLength As Long
Dim sZipFile As String
Dim sZipDirectory As String
Dim sTmp As String
Dim lngRet As Long
Dim uZipNames As ZipNames
Dim uExcNames As ZipNames

    
    On Error GoTo ErrHan:
    
    mbUnipToMem = False
    
    prTheZipMemberString = vbNullString
    prTheMemberCount = 0
    
    mZMemIndex = 0
    msZipFiles = vbNullString
    
    lngLength = Len(pTheFileToUnzip)
    If lngLength = 0 Then Err.Raise vbObjectError + 1000, , "No File Name to Read."
    
    For lngLoop = 1 To lngLength
        sTmp = Mid$(pTheFileToUnzip, lngLoop, 1)
        If sTmp = "\" Then sTmp = "/"
        sZipFile = sZipFile & sTmp
    Next lngLoop
        
    lngLength = Len(pTheUnzipFolder)
    If lngLength > 0 Then
        For lngLoop = 1 To lngLength
            sTmp = Mid$(pTheUnzipFolder, lngLoop, 1)
            If sTmp = "\" Then sTmp = "/"
            sZipDirectory = sZipDirectory & sTmp
        Next lngLoop
    End If
    
    tZipCBPointers.lptrPrnt = pAddressOf(AddressOf UnzipPrintCallback)
    tZipCBPointers.lptrReplace = pAddressOf(AddressOf UnzipReplaceCallback)
    tZipCBPointers.lptrPassword = pAddressOf(AddressOf UnzipPasswordCallBack)
    tZipCBPointers.lptrMessage = pAddressOf(AddressOf UnzipMessageCallBack)
    tZipCBPointers.lptrService = pAddressOf(AddressOf UnZipServiceCallback)
    
    tDCList.szZipName = sZipFile            ' ... file to unzip.
    tDCList.szExtractDir = sZipDirectory    ' ... folder to unzip into.
    tDCList.ndflag = True                   ' ... honour directories.
    tDCList.nvflag = 0                      ' ... unzip.
    tDCList.noflag = True
    
    lngRet = Wiz_SingleEntryUnzip(0, uZipNames, 0, uExcNames, tDCList, tZipCBPointers)
    
    Unzip = lngRet
    
    prTheZipMemberString = msZipFiles   ' ... return the contents string.
    prTheMemberCount = mZMemIndex       ' ... return the contents count.
    
Exit Function
ErrHan:
    Unzip = -1
    Debug.Print "modUnZip.Unzip.Error: " & Err.Description
    
End Function

Private Sub UnzipMessageCallBack(ByVal ucsize As Long, _
                                 ByVal csiz As Long, _
                                 ByVal cfactor As Integer, _
                                 ByVal mo As Integer, _
                                 ByVal dy As Integer, _
                                 ByVal yr As Integer, _
                                 ByVal hh As Integer, _
                                 ByVal mm As Integer, _
                                 ByVal c As Byte, _
                                 ByRef fname As CBCh, _
                                 ByRef meth As CBCh, _
                                 ByVal crc As Long, _
                                 ByVal fCrypt As Byte)

' ... this callback is in response to list zip contents request.
' ... the list is requested by setting DCList.nvFlag to 1 (if 0 then file is unzipped).
' ... the parameters provide all the information required per member of the zip file.
' ... there is a unix convention of using / rather than \ to delimit folders.
' ... info on parameters of interest (for now).

'   ucsize: Long        ' ... the uncompressed size.
'   csiz: Long          ' ... compressed size.
'   mo: Integer         ' ... month
'   dy: Integer         ' ... day
'   yr: Integer         ' ... year
'   hh: Integer         ' ... hour
'   mm: Integer         ' ... minute
'   fname: CBCh         ' ... name of zip member (including zip folder, see above / and \)

Dim lngLoop As Long
Dim sTmpName As String

Dim sTmpFName As String
Dim sTmpFPath As String
Dim dDate As Date
Dim sDate As String
Dim sUnCSize As String
Dim sCSize As String
Dim sTmpMember As String
Dim sEncrypted As String
Dim lngLastSlash As Long
Dim lngChar As Long

    If crc = 0 Then Exit Sub
    
    sUnCSize = CStr(ucsize)
    sCSize = CStr(csiz)
    
    ' ... the name of the member can be extracted as follows
    ' ... presuming the name's length isn't greater than 256
    For lngLoop = 0 To 255
        lngChar = fname.ch(lngLoop)
        If lngChar = 0 Then Exit For
        If lngChar = 47 Then            ' ... e.g. if / then convert to \.
            lngLastSlash = lngLoop + 1  ' ... left of this gives the path and right gives the file name.
            lngChar = 92
        End If
        sTmpName = sTmpName & Chr$(lngChar)
    Next
    
    sTmpFName = sTmpName
    If lngLastSlash > 0 Then
        ' ... split the member into path and file name.
        sTmpFPath = Left$(sTmpName, lngLastSlash - 1)
        sTmpFName = Mid$(sTmpName, lngLastSlash + 1)
    End If
    
    dDate = DateSerial(yr, mo, dy)
    dDate = dDate + TimeSerial(hh, mm, 0)
    sDate = CStr(CDbl(dDate))
    
    sEncrypted = IIf((fCrypt And 64) = 64, "1", "0")
    
    mZMemIndex = mZMemIndex + 1
    
    ' ... Zip Member described as follows:
    '   File Name | File Folder | Full Member Name | Date | Uncomp. Size | Comp. Size | Zip Index | Encrypted
    sTmpMember = sTmpFName & "|" & sTmpFPath & "|" & sTmpName & "|" & sDate & "|" & sUnCSize & "|" & sCSize & "|" & CStr(mZMemIndex) & "|" & sEncrypted
    
    If Len(msZipFiles) Then
        sTmpMember = vbNewLine & sTmpMember
    End If
    msZipFiles = msZipFiles & sTmpMember
    
    sTmpMember = vbNullString
    sTmpFName = vbNullString
    sDate = vbNullString
    dDate = 0

End Sub

Private Function UnZipServiceCallback(ByRef mname As CBChar, ByVal X As Long) As Long
'
' ... this call back is in response to unzipping rather than reading.
' ... The mname param is the name of a file being unzipped.
' ... This is called when files are unzipped physically or just to memory.
' Note:
'   ... when unzipping to memory no stats. are changed.

'Dim s0 As String
'Dim xx As Long
Dim sTmpMember As String
Dim lngLoop As Long
Dim lngChar As Long

    On Error Resume Next    ' ... necessary in all call backs.
    ' -------------------------------------------------------------------
    For lngLoop = 0 To X - 1
        lngChar = mname.ch(lngLoop)
        If lngChar = 0 Then Exit For
        If lngChar = 47 Then            ' ... e.g. if / then convert to \.
            lngChar = 92
        End If
        sTmpMember = sTmpMember & Chr$(lngChar)
    Next lngLoop
    ' -------------------------------------------------------------------
    If mbUnipToMem = False Then
        mZMemIndex = mZMemIndex + 1
        If Len(msZipFiles) Then
            sTmpMember = vbNewLine & sTmpMember
        End If
        msZipFiles = msZipFiles & sTmpMember
    End If
    ' -------------------------------------------------------------------
    UnZipServiceCallback = 0 ' ... Setting this to 1 will abort the zip!
    ' -------------------------------------------------------------------
    sTmpMember = vbNullString
    lngChar = 0&

End Function

Private Function UnzipPasswordCallBack(ByRef pwd As CBCh, ByVal X As Long, ByRef s2 As CBCh, ByRef Name As CBCh) As Long

' ... password call back; get a password from the user.
' ... Note: drafted, not tested.

Dim sPwd As String
Dim lngLoop As Long

    On Error Resume Next
    
    ' ... set default return (cancel unzip).
    UnzipPasswordCallBack = 1
        
    ' ... ask for a password.
    sPwd = InputBox("A Password is required, please provide it.", "Password Protected")
    sPwd = Trim$(sPwd)
    
    If Len(sPwd) = 0 Then Exit Function
    
    ' ... following Chris Eastwood's code from here.
    For lngLoop = 0 To X - 1
        pwd.ch(lngLoop) = 0
    Next lngLoop
    
    For lngLoop = 0 To Len(sPwd) - 1
        pwd.ch(lngLoop) = Asc(Mid$(sPwd, lngLoop + 1, 1))
    Next
    
    ' ... null terminator for c, apparently.
    pwd.ch(lngLoop) = vbNullChar
    
     ' ... return wish to continue.
    UnzipPasswordCallBack = 0
            
End Function

Private Function UnzipPrintCallback(ByRef fname As CBChar, ByVal X As Long) As Long
'
End Function

Private Function UnzipReplaceCallback(ByRef fname As CBChar) As Long
'
End Function

