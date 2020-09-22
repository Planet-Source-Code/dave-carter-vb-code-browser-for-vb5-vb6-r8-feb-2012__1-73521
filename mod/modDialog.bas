Attribute VB_Name = "modDialog"
Attribute VB_Description = "A module to help utilise the Windows COMDLG and SHBrowseForFolder APIs for retrieving file and folder names from the user."

' what?
'  a module with a few common dialog related methods
'  including browse for folder dialog.
' why?
'  reduce dependence upon common dialog control.
' when?
'  when simple common dialog box access is required
'  or when a filter for a common dialog is required.
' how?
'  see examples.
' who?
'  d.c.

'== Subs: 1
'   psMakeAPIFilter:              "replaces | and : chars in a common dialog filter string for passing to API."
'
'== Functions: 4
'+  GetOpenFileName: String       "Get a file name from the Common Dialog Open File Name dialog."
'+  GetFolder: String             "Returns the name of a folder by way of the Browse for folder Dialog."
'+  MakeDialogFilter: String      "Creates a Common Dialog Filter String a single filter for opening and saving."
'+  MakeDialogMultiFilter: St..   "Creates a Common Dialog Filter string for a multiple filter for Opening and Saving files."

Option Explicit

' -------------------------------------------------------------------
' ... Open File Name Related.

'Private Const OFN_DONTADDTORECENT As Long = &H2000000
'Private Const OFN_EXPLORER As Long = &H80000
'Private Const OFN_FILEMUSTEXIST As Long = &H1000
'Private Const OFN_PATHMUSTEXIST As Long = &H800
'Private Const OFN_NOCHANGEDIR As Long = &H8
' ... cOpenFLags is the sum of the above constants for one flag msg to open file name dialog.
Private Const cOpenFlags As Long = &H2081808


' ... structure members not used are prefixed with z so they are at the end of
' ... the intelliesense pop-up.
' ... structure members prefixed with x are set internally with defaults.

Private Type OpenFileInfo
    StructSize As Long              ' Filled with UDT size
    OwnerHandle As Long             ' Tied to vOwner
    xInstance As Long               ' Ignored (used only by templates)
    Filter As String                ' Tied to vFilter
    zCustomFilter As String         ' Ignored (exercise for reader)
    zMaxCustFilter As Long          ' Ignored (exercise for reader)
    FilterIndex As Long             ' Tied to vFilterIndex
    File As String                  ' Tied to vFileName
    xMaxFile As Long                ' Handled internally
    FileTitle As String             ' Tied to vFileTitle
    xMaxFileTitle As Long           ' Handle internally
    InitialDir As String            ' Tied to vInitDir
    DialogTitle As String           ' Tied to vTitle
    flags As Long                   ' Tied to vFlags
    zFileOffset As Integer          ' Ignored (exercise for reader)
    zFileExtension As Integer       ' Ignored (exercise for reader)
    DefaultExt As String            ' Tied to vDefExt
    zCustData As Long               ' Ignored (needed for hooks)
    zHook As Long                   ' Ignored (no hooks in Basic)
    zTemplateName As Long           ' Ignored (no templates in Basic)
End Type

Private Declare Function fGetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (File As OpenFileInfo) As Long
Private Const cMaxPath = 260
Private Const cMaxFile = 260

' -------------------------------------------------------------------
' ... Browse For Folder Related.
' ... check the link below for a more complete version of the get folder function,
' ... I used it to help me with the CallBack bit.
' ... The lpszTitle member of the BrowseInfo Type came out as a long in my
' ... API viewer, but OnError's version has it as a string...
' ... I guess if it were a long I'd have to do a StrPtr to the string with the title.

'By OnError @
'http://www.xtremevbtalk.com/showthread.php?t=213821
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal csidl As Long, ByRef ppidl As Long)

'Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Const WM_USER As Long = &H400

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_USENEWUI As Long = &H40

Private Const BFFM_INITIALIZED As Long = 1                  ' ... message from browser.
Private Const BFFM_SETSELECTIONA As Long = WM_USER + 102    ' ... message to browser.

Private Const CSIDL_DRIVES As Long = &H11

Private Type BrowseInfo
  hwndOwner         As Long
  pIDLRoot          As Long
  pszDisplayName    As Long
  lpszTitle         As String
  ulFlags           As Long
  lpfnCallback      As Long
  lParam            As Long
  iImage            As Long
End Type

Private Const cCharAsterix As String = "*"
Private Const cCharBang As String = "|"
Private Const cCharPeriod As String = "."
Private Const cCharColon As String = ";"
Private Const cCharComma As String = ","
Private Const cCharLBrack As String = "("
Private Const cCharRBrack As String = ")"

Private Const cWrdAllFiles As String = "All Files"

Private Function BrowseFolderCallBack(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal sData As String) As Long
Attribute BrowseFolderCallBack.VB_Description = "Call Back method to the SHBrowseForFolder API Function, used here just to set the initial directory."
    If uMsg = BFFM_INITIALIZED Then
        SendMessage hwnd, BFFM_SETSELECTIONA, True, ByVal sData
    End If
End Function

Public Function GetFolder(Optional ByVal pTheBrowserTitle As String = vbNullString, _
                          Optional ByVal pInitialDirectory As String = vbNullString) As String
Attribute GetFolder.VB_Description = "Get a Folder from the user via the Browse Folder Dialog."

' ... get a folder from the user via the browse folder dialog.
' ... example:
'       ' using from a Form: note a Window Handle is Required,
'       ' open the browse folder dialog setting its initial directory to C:\.
'       sFolder = GetFolder("Select a Folder","C:\")
'       If Len(sFolder) Then        ' ... we have a folder selected.

Dim iNull As Integer
Dim lpIDList As Long
Dim sPath As String
Dim udtBI As BrowseInfo
Dim ppidl As Long
        
    On Error GoTo ErrHan:
    
    With udtBI
        
        ' ... Set the owner window, if no active window then fall into error.
        .hwndOwner = Screen.ActiveForm.hwnd
        
        SHGetSpecialFolderLocation .hwndOwner, CSIDL_DRIVES, ppidl
        .pIDLRoot = ppidl
        
        .lpszTitle = pTheBrowserTitle
        
        If Dir$(pInitialDirectory, vbDirectory) <> "" Then
            .lpfnCallback = AddressOfMethod(AddressOf BrowseFolderCallBack)
            .lParam = StrPtr(pInitialDirectory)
        End If
        
        ' ... Return only if the user selected a directory.
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_USENEWUI

    End With

    ' ... Show the 'Browse for folder' dialog.
    lpIDList = SHBrowseForFolder(udtBI)
    
    If lpIDList Then
        
        ' ... create buffer for folder path.
        sPath = String$(cMaxPath, 0)
        
        ' ...Get the path from the IDList.
        SHGetPathFromIDList lpIDList, sPath
        
        ' ... free the block of memory; developer is responisible for this.
        CoTaskMemFree lpIDList
        
        ' ... read the folder path.
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
        
    End If

ResumeError:

    GetFolder = sPath
    
    sPath = vbNullString

Exit Function

ErrHan:

    Debug.Print "modDialog.GetFolder.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:

End Function

Function GetOpenFileName(Optional ByRef prFileName As String = vbNullString, _
                         Optional ByRef prFileTitle As String = vbNullString, _
                         Optional ByVal pvFilter As String = "All Files (*.*)| *.*", _
                         Optional ByVal pvFilterIndex As Long = 1, _
                         Optional ByVal pvInitialDirectory As String = vbNullString, _
                         Optional ByVal pvDialogTitle As String = vbNullString, _
                         Optional ByVal pvDefaultExtension As String = vbNullString, _
                         Optional ByVal pvOwnerHandle As Long = -1) As String
Attribute GetOpenFileName.VB_Description = "Get a file name from the Common Dialog Open File Name dialog."

' ... Get a file name from the Common Dialog Open File Name dialog.
' ... Over simplified version of calling COMDLG.GetOpenFileName.

' ... Sets the following flags with the cOpenFLags Constant,
' ... OFN_DONTADDTORECENT As Long = &H2000000
' ... OFN_EXPLORER As Long = &H80000
' ... OFN_FILEMUSTEXIST As Long = &H1000
' ... OFN_PATHMUSTEXIST As Long = &H800
' ... OFN_NOCHANGEDIR As Long = &H8

' ... Returns the full path and name of file selected or empty string.
' ... prFileName is updated with just the name (no extension/path) of the file.
' ... prFileTitle is updated with the name of the file (without its path).
' ... pvDefaultExtension is updated with the extension of the file selected.

' ... example.
'   ... No Filter, just get a file name back.
'       ... sFileName = GetOpenFileName()
'   ... Filter VB Projects,
'       ... sFilter = MakeDialogFilter("VB Projects",,"vbp)
'       ... sFileName = GetOpenFileName(,,sFilter)
'   ... Set Name of file to open and include it in the filter,
'       ... sFilter = modDialog.MakeDialogFilter("VB Projects", "CodeViewer", "vbp")
'       ... sFileName = modDialog.GetOpenFileName("CodeViewer.vbp", , sFilter, , , "Open VBP")

'       ... If Len(sFileName) Then  ' ... we have a file name.

Dim xOpenFileInfo As OpenFileInfo
Dim lngFindCharZero As Long
Dim lngFindCharPeriod As Long
Dim sFilter As String
Dim sTmpString As String
Dim lngResult As Long
    
    
    With xOpenFileInfo
        
        .StructSize = Len(xOpenFileInfo)
        
        .flags = cOpenFlags ' ... flags is a composite of several flags, see top of module.
        
        sTmpString = prFileName & String$(cMaxPath - Len(prFileName), 0)
        .File = sTmpString
        
        sTmpString = prFileTitle & String$(cMaxFile - Len(prFileTitle), 0)
        .FileTitle = sTmpString
        
        If Len(pvInitialDirectory) Then
            .InitialDir = pvInitialDirectory
        End If
        
        If Len(pvDialogTitle) Then
            .DialogTitle = pvDialogTitle
        End If
        
        If Len(pvDefaultExtension) Then
            .DefaultExt = pvDefaultExtension
        End If
        
        If pvOwnerHandle <> -1 Then
            .OwnerHandle = pvOwnerHandle
        End If
        
        ' ... make the filter ok for the use within the api call.
        psMakeAPIFilter pvFilter, sFilter
        
        .Filter = sFilter
        .FilterIndex = pvFilterIndex
        
        .xMaxFile = cMaxPath
        .xMaxFileTitle = cMaxFile
        
        lngResult = fGetOpenFileName(xOpenFileInfo)
        
        If lngResult <> 0 Then ' 0 = file selected and ok clicked.
            
            If lngResult = 1 Then
                
                lngFindCharZero = InStr(1, .FileTitle, vbNullChar)
                If lngFindCharZero > 0 Then
                    sTmpString = Left$(.FileTitle, lngFindCharZero - 1)
                    prFileTitle = sTmpString
                    lngFindCharPeriod = InStr(1, sTmpString, cCharPeriod)
                    If lngFindCharPeriod > 0 Then
                        prFileName = Left$(sTmpString, lngFindCharPeriod - 1)
                        pvDefaultExtension = Mid$(sTmpString, lngFindCharPeriod + 1)
                    End If
                End If
            
                lngFindCharZero = InStr(1, .File, vbNullChar)
                If lngFindCharZero > 0 Then
                    sTmpString = Left$(.File, lngFindCharZero - 1)
                    GetOpenFileName = sTmpString
                End If
            End If
            
        End If
        
    End With


End Function

Public Function MakeDialogFilter(Optional pThePrompt As String = cWrdAllFiles, _
                                 Optional pTheFileName As String = cCharAsterix, _
                                 Optional pTheExtension As String = cCharAsterix) As String
Attribute MakeDialogFilter.VB_Description = "Creates a Common Dialog Filter String a single filter for opening and saving."

' ... return a common dialog filter.
' ... calling this with no parameters returns 'All Files (*.*) | *.*

' ... pThePrompt is the All Files bit as above.
' ... pTheFileName is the name of the file and defaults to asterix if not passed.
' ... pTheExtension is the extension to filter on and defaults to asterix as pTheFileName.

' ... examples:
'   ... Return a filter for VB Projects (vbp)
'   ... MakeDialogFilter("VB Projects","*","vbp")           = "VB Projects (*.vbp) | *.vbp"

'   ... Return a filter for a VB Project (Test.vbp)
'   ... MakeDialogFilter("Test VB Project","Test","vbp")    = "Test VB Project (Test.vbp) | Test.vbp"

Dim sTmp As String
Dim sTmpName As String
    
    sTmpName = pTheFileName & cCharPeriod & pTheExtension
    
    sTmp = pThePrompt & Space$(1) & cCharLBrack & sTmpName & cCharRBrack & cCharBang & sTmpName
    
    MakeDialogFilter = sTmp
    
    sTmp = vbNullString
    sTmpName = vbNullString
    
End Function

Public Function MakeDialogMultiFilter(pThePrompt As String, ParamArray pTheExtension() As Variant) As String
Attribute MakeDialogMultiFilter.VB_Description = "Creates a Common Dialog Filter string for a multiple filter for Opening and Saving files."

' ... return a common dialog filter for a range of types of file.
' ... e.g. Metafiles|*.wmf;*.emf
' ... the call to make this return would be
'   ... MakeDialogMultiFilter("MetaFiles","wmf","emf")
' ... Note: this does not use a file name.
' ...       the ParamArray is a comma delimited list of strings wanted as available extensions
' ...       e.g.
'   ... MakeDialogMultiFilter("Pic Files","bmp","wmf","gif")    = "Pic Files|*.bmp;*.wmf;*.gif"

Dim sTmp As String
Dim sTmpName As String
Dim sPrompt As String
Dim sTmpPrompt As String
Dim lngLoop As Long

Dim lngUBnd As Long
Dim lngLBnd As Long

    On Error GoTo ErrHan:
    
    lngUBnd = UBound(pTheExtension)
    lngLBnd = LBound(pTheExtension)
    
    sPrompt = pThePrompt
    
    For lngLoop = lngLBnd To lngUBnd
        If Len(sTmpName) > 0 Then
            sTmpName = sTmpName & cCharColon
        End If
        If Len(sTmpPrompt) > 0 Then
            sTmpPrompt = sTmpPrompt & cCharComma & Space$(1)
        End If
        sTmp = pTheExtension(lngLoop)
        sTmpName = sTmpName & cCharAsterix & cCharPeriod & sTmp
        sTmpPrompt = sTmpPrompt & sTmp
    Next lngLoop
    
    If Len(sTmpPrompt) > 0 Then
        sPrompt = sPrompt & Space$(1) & cCharLBrack & sTmpPrompt & cCharRBrack
    End If
    
    sTmp = sPrompt & cCharBang & sTmpName
    
    MakeDialogMultiFilter = sTmp
    
ResumeErr:

    sTmp = vbNullString
    sTmpName = vbNullString
    
Exit Function
ErrHan:
    Debug.Print "Error.modDialog.MakeDialogMultiFiler: " & Err.Description
    Resume ResumeErr:
    
End Function

Private Sub psMakeAPIFilter(ByVal pvTheFilter As String, ByRef prTheAPIFilter As String)
Attribute psMakeAPIFilter.VB_Description = "replaces | and : chars in a common dialog filter string for passing to API."

' ... replaces | and : chars in a common dialog filter string for passing to API.
' ... the API filter is returned in prTheAPIFilter.

Dim bFilter() As Byte
Dim sReturn As String
Dim lngLoop As Long
Dim lngPos As Long
Dim lngChar As Long

    If Len(pvTheFilter) = 0 Then Exit Sub
    
    bFilter = pvTheFilter
    
    ' ... create a return buffer of char zeros plus 2 for default termination of the filter.
    sReturn = String$(Len(pvTheFilter) + 2, 0)
    
    ' ... loop the bytes of the filter and write the return, replacing as required.
    For lngLoop = 0 To UBound(bFilter) Step 2
        lngPos = lngPos + 1
        lngChar = bFilter(lngLoop)
        Select Case lngChar
            Case 124, 58 ' ... | and : respectively, after RR's TinyGFX.
                lngChar = 0
        End Select
        Mid$(sReturn, lngPos, 1) = Chr$(lngChar)
    Next lngLoop
    
    prTheAPIFilter = sReturn
    
    sReturn = vbNullString
    Erase bFilter
        
End Sub

Function GetFolders(pPath As String) As String
Dim pResults As String
'    GetFolderList pPath, pResults
'    GetFolders = pResults
    GetFolders = pGetFolders(pPath, pResults)
End Function

Private Function pGetFolders(ByVal pPath As String, ByRef pResults As String) As String

' ... attempts to return the first level of sub folders to a folder.
' ... pResults can be used to accumulate all the folders found.

Dim sDirName As String
Dim sTmpPath As String
Dim sResult As String

    On Error Resume Next

    sTmpPath = pPath
    sResult = sTmpPath
    
    If Right$(sTmpPath, 1) <> "\" Then sTmpPath = sTmpPath & "\"
    'sDirName = Dir$(sTmpPath & "*.*", vbDirectory)
    sDirName = Dir$(sTmpPath, vbDirectory)
    
    Do While Len(sDirName)
    
        If sDirName <> "." And sDirName <> ".." And Len(sDirName) Then
            If GetAttr(sTmpPath & sDirName) And vbDirectory Then
                If Len(sResult) Then sResult = sResult & vbNewLine
                sResult = sResult & sTmpPath & sDirName
            End If
        End If
        
        sDirName = Dir$(, vbDirectory)
        
    Loop
    
    pGetFolders = sResult
    
    If Len(sResult) Then
        If Len(pResults) Then pResults = pResults & vbNewLine
        pResults = pResults & sResult
    End If
    
    sDirName = vbNullString
    sResult = vbNullString
    sTmpPath = vbNullString
    
End Function
'
'Private Sub GetFolderList(ByVal pPath As String, pResults As String)
'
'' ... recursive routine to build a list of sub folders to a parent folder.
'' ... pResults will return a vbCrLF delimited string naming all sub folders found.
'
'Dim sFoldersArray() As String
'Dim lngLoop As Long
'Dim sFoldersString As String
'
'    On Error GoTo ErrHan:
'
'    If Len(pPath) = 0 Then
'        Err.Raise vbObjectError + 1000, , "Not a Valid Path to query."
'    End If
'
'    If Right$(pPath, 1) <> "\" Then pPath = pPath & "\"
'
'    sFoldersString = pGetFolders(pPath, pResults)
'    modStringArrays.SplitString sFoldersString, sFoldersArray, vbCrLf
'
'    For lngLoop = 1 To UBound(sFoldersArray)
'        GetFolderList sFoldersArray(lngLoop), pResults
'    Next
'
'ResumeError:
'
'    On Error GoTo 0
'
'    Erase sFoldersArray
'    sFoldersString = vbNullString
'    lngLoop = 0&
'
'Exit Sub
'ErrHan:
'
'    Debug.Print "modZip.GetFolderList.Error: " & Err.Number & "; " & Err.Description
'    Resume ResumeError:
'
'End Sub
'
