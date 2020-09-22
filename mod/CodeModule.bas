Attribute VB_Name = "modZip"
Attribute VB_Description = "A simple Zip module providing a very basic Zip interface."

' what?
'  a bare bones zip module using Zip32.dll from Info-Zip.
' why?
'  compliment project copy function.
' when?
'  when copying a project using the copy project function.
' how?
'  make a call to VBZipEx passing the name of the file
'  that should be written.
' note:
'  the zip file name's folder is presumed to be the
'  folder to zip.
'  take care with the path, once it starts it will try to go on.
' who?
'  d.c. adapted from CodeModule written by Chris Eastwood.

Option Explicit

' ... used to pass the names of the files to zip
' ... to ZpArchive.
Private Type ZipNames
    FileNames(0 To 99) As String
End Type

' ... ZipOptions is used to set options in the zip32.dll.
Private Type ZipOptions
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    Date As String ' 8 bytes long
    szRootDir As String ' up to 256 bytes long
End Type

' ... used to initialise Zip Archiving with the
' ... addresses of Call Back methods.
Private Type ZipCallBacks
    ZipPrintCallback As Long
    ZipPasswordCallback As Long
    ZipCommentCallback As Long
    ZipServiceCallback As Long
End Type


' Call back "string" (sic)
Private Type CBChar
    ch(4096) As Byte
End Type


' ... This assumes zip32.dll is in your \windows\system directory!
Private Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZipCallBacks) As Long ' Set Zip Callbacks
Private Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZipOptions) As Long ' Set Zip options
Private Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZipNames) As Long ' Real zipping action

Public msOutput As String
Private mWithDebug As Boolean

Private Function GetFolders(ByVal pPath As String, ByRef pResults As String) As String
Attribute GetFolders.VB_Description = "attempts to return the first level of sub folders to a folder."

' ... attempts to return the first level of sub folders to a folder.
' ... pResults can be used to accumulate all the folders found.

Dim sDirName As String
Dim sTmpPath As String
Dim sResult As String

    On Error Resume Next

    sTmpPath = pPath
    sResult = sTmpPath
    
    If Right$(sTmpPath, 1) <> "\" Then sTmpPath = sTmpPath & "\"
    sDirName = Dir$(sTmpPath & "*.*", vbDirectory)

    Do While Len(sDirName)
    
        If sDirName <> "." And sDirName <> ".." And Len(sDirName) Then
            If GetAttr(sTmpPath & sDirName) And vbDirectory Then
                If Len(sResult) Then sResult = sResult & vbNewLine
                sResult = sResult & sTmpPath & sDirName
            End If
        End If
        
        sDirName = Dir$(, vbDirectory)
        
    Loop
    
    GetFolders = sResult
    
    If Len(sResult) Then
        If Len(pResults) Then pResults = pResults & vbNewLine
        pResults = pResults & sResult
    End If
    
    sDirName = vbNullString
    sResult = vbNullString
    sTmpPath = vbNullString
    
End Function

Private Function GetFolderDirectories(ByVal pPath As String, ByRef pZipNames As ZipNames, ByRef pCount As Long) As String
Attribute GetFolderDirectories.VB_Description = "get a list of all sub folders to process."

Dim sFoldersString As String
Dim res() As String
Dim sTmp As String
Dim lngLoop As Long
    
    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... get a list of all sub folders to process.
    GetFolderList pPath, sFoldersString
    ' -------------------------------------------------------------------
    
    If Len(sFoldersString) Then
        ' -------------------------------------------------------------------
        ' ... split the folder list string.
        modStringArrays.SplitString sFoldersString, res, vbCrLf, pCount
            
        If pCount Then
            
            ' -------------------------------------------------------------------
            ' ... add a dummy ZipNames item in first position.
            pZipNames.FileNames(0) = "qwerty"
            
            ' -------------------------------------------------------------------
            ' ... loop the sub folders adding *.* on the end to include all files.
            For lngLoop = 0 To pCount - 1
                
                sTmp = res(lngLoop)
                
                If Right$(sTmp, 1) <> "\" Then sTmp = sTmp & "\"
                ' -------------------------------------------------------------------
                ' ... update ZipNames entry.
                pZipNames.FileNames(lngLoop + 1) = sTmp & "*.*"
                
            Next lngLoop
        
        End If
            
    End If

ResumeError:
    
    On Error GoTo 0
    
    Erase res
    sFoldersString = vbNullString
    sTmp = vbNullString
    lngLoop = 0&
    
Exit Function
ErrHan:

    Debug.Print "modZip.GetFolderDirectories.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function

Private Sub GetFolderList(ByVal pPath As String, pResults As String)
Attribute GetFolderList.VB_Description = "recursive routine to build a list of sub folders to a parent folder."

' ... recursive routine to build a list of sub folders to a parent folder.
' ... pResults will return a vbCrLF delimited string naming all sub folders found.

Dim sFoldersArray() As String
Dim lngLoop As Long
Dim sFoldersString As String
        
    On Error GoTo ErrHan:
    
    If Len(pPath) = 0 Then
        Err.Raise vbObjectError + 1000, , "Not a Valid Path to query."
    End If
    
    If Right$(pPath, 1) <> "\" Then pPath = pPath & "\"
    
    sFoldersString = GetFolders(pPath, pResults)
    modStringArrays.SplitString sFoldersString, sFoldersArray, vbCrLf
    
    For lngLoop = 1 To UBound(sFoldersArray)
        GetFolderList sFoldersArray(lngLoop), pResults
    Next

ResumeError:
    
    On Error GoTo 0
    
    Erase sFoldersArray
    sFoldersString = vbNullString
    lngLoop = 0&
    
Exit Sub
ErrHan:

    Debug.Print "modZip.GetFolderList.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

' ... Callback for zip32.dll.
Private Function ZipPrintCallback(ByRef fname As CBChar, ByVal x As Long) As Long
Attribute ZipPrintCallback.VB_Description = "Zip32 CallBack: called as each file to zip is processed by zip32."

Dim s0$
Dim xx As Long
Dim sVbZipInf As String
    
    ' always put this in callback routines!
    On Error Resume Next
    s0 = ""
    For xx = 0 To x
        If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 + Chr(fname.ch(xx))
    Next xx
    
    If mWithDebug Then
        Debug.Print sVbZipInf & s0
    End If
    
    msOutput = msOutput & s0
    
    sVbZipInf = ""
    
    DoEvents
    ZipPrintCallback = 0
    
End Function

' ... Callback for Zip32.dll ?
Private Function ZipServiceCallback(ByRef fname As CBChar, ByVal x As Long) As Long
Attribute ZipServiceCallback.VB_Description = "Zip32 CallBack: not sure of this callback's purpose and not used other than to return 0."

'Dim s0 As String
'Dim xx As Long
'
'    On Error Resume Next
'
'    s0 = ""
'
'    For xx = 0 To X - 1
'        If fname.ch(xx) = 0 Then Exit For
'        s0 = s0 & Chr$(fname.ch(xx))
'    Next
    
    ' -------------------------------------------------------------------
    ' ... not interested?
    
    ZipServiceCallback = 0
    
    
End Function

' ...Callback for zip32.dll
Private Function ZipPasswordCallback(ByRef s1 As Byte, x As Long, _
                                     ByRef s2 As Byte, _
                                     ByRef s3 As Byte) As Long
Attribute ZipPasswordCallback.VB_Description = "Zip32 CallBack: not used here other than to return 1, must be something to do with applying password to zipped files."

    On Error Resume Next
    
    ' ... not supported - always return 1
    ZipPasswordCallback = 1
    
End Function

' ... Callback for zip32.dll.
Private Function ZipCommentCallback(ByRef s1 As CBChar) As CBChar
    
    ' always put this in callback routines!
    On Error Resume Next
    ' not supported always return \0
    s1.ch(0) = vbNullString
    ZipCommentCallback = s1
    
End Function

Public Function VBZipEx(pTheZipFileName As String, Optional pWIthDebug As Boolean = False) As Long
Attribute VBZipEx.VB_Description = "Public Interface to running Zip32.  Very basic with no nice refinements as yet."

' ... very basic zip routine, no nice refinements, no point (as yet).

Dim lngReturn As Long
Dim tZipCallBacks As ZipCallBacks
Dim tZipOptions As ZipOptions
Dim lngFileCount As Long
Dim prtZipNames As ZipNames
Dim tFileNameInfo As FileNameInfo
    
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    ' ... shred the name of the zip file for its folder.
    modFileName.ParseFileNameEx pTheZipFileName, tFileNameInfo
    ' -------------------------------------------------------------------
    mWithDebug = pWIthDebug
    msOutput = ""
    ' -------------------------------------------------------------------
    ' ... set addresses of callback functions.
    tZipCallBacks.ZipPrintCallback = pAddressOf(AddressOf ZipPrintCallback)
    tZipCallBacks.ZipPasswordCallback = pAddressOf(AddressOf ZipPasswordCallback)
    tZipCallBacks.ZipCommentCallback = pAddressOf(AddressOf ZipCommentCallback)
    tZipCallBacks.ZipServiceCallback = 0& ' not coded yet :-)
    ' -------------------------------------------------------------------
    ' ... set zip options
    tZipOptions.fSuffix = 0             ' include suffixes (not yet implemented)
    tZipOptions.fEncrypt = 0            ' 1 if encryption wanted
    tZipOptions.fSystem = 0             ' 1 to include system/hidden files
    tZipOptions.fVolume = 0             ' 1 if storing volume label
    tZipOptions.fExtra = 0              ' 1 if including extra attributes
    tZipOptions.fNoDirEntries = 0       ' 1 if ignoring directory entries
    tZipOptions.fExcludeDate = 0        ' 1 if excluding files earlier than a specified date
    tZipOptions.fIncludeDate = 0        ' 1 if including files earlier than a specified date
    tZipOptions.fVerbose = 0            ' 1 if full messages wanted
    tZipOptions.fQuiet = 0              ' 1 if minimum messages wanted
    tZipOptions.fCRLF_LF = 0            ' 1 if translate CR/LF to LF
    tZipOptions.fLF_CRLF = 0            ' 1 if translate LF to CR/LF
    tZipOptions.fJunkDir = 0            ' 1 if junking directory names
    tZipOptions.fRecurse = 0            ' 1 if recursing into subdirectories
    tZipOptions.fGrow = 0               ' 1 if allow appending to zip file
    tZipOptions.fForce = 0              ' 1 if making entries using DOS names
    tZipOptions.fMove = 0               ' 1 if deleting files added or updated
    tZipOptions.fDeleteEntries = 0      ' 1 if files passed have to be deleted
    tZipOptions.fUpdate = 0             ' 1 if updating zip file--overwrite only if newer
    tZipOptions.fFreshen = 0            ' 1 if freshening zip file--overwrite only
    tZipOptions.fJunkSFX = 0            ' 1 if junking sfx prefix
    tZipOptions.fLatestTime = 0         ' 1 if setting zip file time to time of latest file in archive
    tZipOptions.fComment = 0            ' 1 if putting comment in zip file
    tZipOptions.fOffsets = 0            ' 1 if updating archive offsets for sfx Files
    tZipOptions.fPrivilege = 0          ' 1 if not saving privelages
    tZipOptions.fEncryption = 0         ' Read only property!
    tZipOptions.fRepair = 0             ' 1=> fix archive, 2=> try harder to fix
    tZipOptions.flevel = 0              ' compression level - should be 0!!!
    tZipOptions.Date = vbNullString     ' "12/31/79"? US Date?
    tZipOptions.szRootDir = "\"
    ' -------------------------------------------------------------------
    ' ... pass call back methods.
    lngReturn = ZpInit(tZipCallBacks)
    ' -------------------------------------------------------------------
    ' ... convey options.
    lngReturn = ZpSetOptions(tZipOptions)
    ' -------------------------------------------------------------------
    ' ... build list of files to zip from the shredded zip path info.
    ' ... all we're doing is adding each folder + "\*.*", not actual file names.
    GetFolderDirectories tFileNameInfo.Path, prtZipNames, lngFileCount
    ' -------------------------------------------------------------------
    lngReturn = -1 ' ... preset a default having used above.
    ' ... try running the zip command, file count is plus one because we add a dummy
    ' ... entry to the list of folders to zip (apparently the first zip entry always fails).
    lngReturn = ZpArchive(lngFileCount + 1, pTheZipFileName, prtZipNames)
    ' -------------------------------------------------------------------
    
ResumeError:
    
    ' -------------------------------------------------------------------
    VBZipEx = lngReturn
    ' -------------------------------------------------------------------
    lngReturn = 0&
    lngFileCount = 0&
    
Exit Function

ErrHan:

    Debug.Print "modZip.VBZipEx.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function


Private Function pAddressOf(ByVal lPtr As Long) As Long
Attribute pAddressOf.VB_Description = "VB Bug workaround to gaining member address pointer for CallBacks."
   ' VB Bug workaround fn
   pAddressOf = lPtr
End Function

