Attribute VB_Name = "modFileIcons"
Attribute VB_Description = "A module to help utilise the Windows SHGetFileInfo API to retrieve the small icon and description of a registered file type on the system."
' what?
'  get the file shell description and small icon and add them
'  to an Image List.
'  Also, just get the icons and file type info from a file extension.
' why?
'  to help display the contents of a zip file with the system
'  associated icon for its type.
'  bits from file extension only are just because...
' when?
'  loading zip file info into list view with an image list.
' how?
'   make a call to AddIconToImageList passing a file name, an image list and a default key
'   to return the key to the image in the list and then use this to associate an icon in a list view.
'   Example:
'   ... Add some file names to a list view.
'
'    For i = 1 To FileCount
'
'        sFile = Filenames(i)
'
'        sIcon = AddIconToImageList(sFile, ImageList1, "DEFAULT")
'
'        Set itmX = ListView1.ListItems.Add(, "File" & i, sFile, , sIcon)
'
'    Next i
'
' who?
'  d.c. (originally inspired by Steve McMahon of vbAccelerator).
'

Option Explicit

' Requires:
'   ... modHandCursor
'   ... modFileName

' Notes:
'   This was originally based upon vbAccelerator's (Steve McMahon) mFileIcons.bas found in
'   ' Unzipping files using the free Info-Zip Unzip DLL with VB ' @
'   http://www.vbaccelerator.com/home/vb/code/libraries/compression/unzipping_files/article.asp
'
'   I have changed a couple of things...
'   the main thing is that the getting of icon and file info
'   is not dependent upon creating a file first.  The SHGetFileInfo
'   will accept a file extension as well as a file name.

' Update Notes: 09 April 2011
'   Added public interfaces to getting icons and file info
'   by file extension only (not file name);
'   the extension (in all cases) needs to include the period as prefix
'   e.g. ".txt" <> "txt", ".doc" <> "doc"

Private Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Const SHGFI_ICON As Long = &H100
Private Const SHGFI_TYPENAME As Long = &H400
Private Const SHGFI_SMALLICON As Long = &H1
Private Const SHGFI_LARGEICON As Long = &H0
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10

' -------------------------------------------------------------------
' v6, adding public interface to getting small and large icons from a file extension.
'     includes File Type Info.

Public Function GetFileSmallIcon(ByVal pTheFileExtension As String, Optional pOK As Boolean = False) As IPicture
Attribute GetFileSmallIcon.VB_Description = "Return the small icon of a file from the file's extension.  pOK returns True if successful else False."
        
    pOK = False
    
    Set GetFileSmallIcon = pGetFileIconSmall(pTheFileExtension)
    
    If Not GetFileSmallIcon Is Nothing Then
        pOK = True
    End If
    
End Function

Public Function GetFileLargeIcon(ByVal pTheFileExtension As String, Optional pOK As Boolean = False) As IPicture
Attribute GetFileLargeIcon.VB_Description = "Return the large icon of a file from the file's extension.  pOK returns True if successful else False."
    
    pOK = False
    
    Set GetFileLargeIcon = pGetFileIconLarge(pTheFileExtension)
    
    If Not GetFileLargeIcon Is Nothing Then
        pOK = True
    End If
    
End Function

Public Function GetFileTypeInfo(ByVal pTheFileExtension As String, Optional pOK As Boolean = False) As String
Attribute GetFileTypeInfo.VB_Description = "Returns the File Type Info about a file from its extension (e.g. the description provided in explorer for a file).  pOK returns True if successful else False."
    
    ' ... K, so this isn't icon getting, may not be so well placed in this module
    ' ... but added because private version exists anyway and is handy to have.
    
    ' Note:
    '   ... the extension requires a preceding dot e.g ".txt" not just "txt"
    '   ... without the dot we just get the word "File" in return.
    
    pOK = False
    
    GetFileTypeInfo = pGetFileTypeInfo(pTheFileExtension)
    
    pOK = CBool(Len(GetFileTypeInfo) > 0)
    
End Function

' -------------------------------------------------------------------

Private Function pGetFileIconSmall(ByVal pTheExt As String) As IPicture
Attribute pGetFileIconSmall.VB_Description = "Returns the icon of a file (type) by its extension as a Picture."

' ... get a file's small icon by file extension.

Dim lRet As Long
Dim hIcon As Long
Dim tSHI As SHFILEINFO
Dim lFlags As Long

    ' ... use SHGFI_LARGEICON As Long ( = &H0 )
    ' ... to get the 32 by 32 large icon in place of SHGFI_SMALLICON in the flags below.
    
    lFlags = SHGFI_USEFILEATTRIBUTES + SHGFI_ICON + SHGFI_SMALLICON

    lRet = SHGetFileInfo(pTheExt, FILE_ATTRIBUTE_NORMAL, tSHI, Len(tSHI), lFlags)
    
    If (lRet <> 0) Then
        
        hIcon = tSHI.hIcon
        
        If hIcon <> 0 Then
            Set pGetFileIconSmall = modHandCursor.HandleToPicture(hIcon, False)
        End If
    
    End If
    
End Function

Private Function pGetFileIconLarge(ByVal pTheExt As String) As IPicture
Attribute pGetFileIconLarge.VB_Description = "Get a File's Large Icon by File Extension."

' ... get a file's large icon by file extension.

' v6, added to compliment small icon function.

Dim lRet As Long
Dim hIcon As Long
Dim tSHI As SHFILEINFO
Dim lFlags As Long
    
    lFlags = SHGFI_USEFILEATTRIBUTES + SHGFI_ICON + SHGFI_LARGEICON

    lRet = SHGetFileInfo(pTheExt, FILE_ATTRIBUTE_NORMAL, tSHI, Len(tSHI), lFlags)
    
    If (lRet <> 0) Then
        
        hIcon = tSHI.hIcon
        
        If hIcon <> 0 Then
            Set pGetFileIconLarge = modHandCursor.HandleToPicture(hIcon, False)
        End If
    
    End If
    
End Function

Private Function pGetFileTypeInfo(ByVal pTheExt As String) As String
Attribute pGetFileTypeInfo.VB_Description = "Returns the description of a file (type) by its extension."
        
' ... get file type info by extension only.
        
Dim lRet As Long
Dim tSHI As SHFILEINFO
Dim iPos As Long
Dim lFlags As Long

    lFlags = SHGFI_USEFILEATTRIBUTES + SHGFI_TYPENAME
    
    lRet = SHGetFileInfo(pTheExt, FILE_ATTRIBUTE_NORMAL, tSHI, Len(tSHI), lFlags)
    
    If lRet <> 0 Then
        iPos = InStr(tSHI.szTypeName, vbNullChar)
        If iPos = 0 Then
            pGetFileTypeInfo = tSHI.szTypeName
        ElseIf iPos > 1 Then
            pGetFileTypeInfo = Left$(tSHI.szTypeName, (iPos - 1))
        End If
    End If
    
End Function

Public Function AddIconToImageList(ByVal pTheFileName As String, _
                                   ByRef pTheImageList As ImageList, _
                                   ByVal pTheDefault As String, _
                          Optional ByRef prFileInfo As String = vbNullString) As String
Attribute AddIconToImageList.VB_Description = "Get the Icon for a File Type and add it to the Image List if does not already exist."

' ... get the icon for a file type and add to image list if does not already exist.
' ... returns the key to the icon else the default passed.
' ... adds file type info to the tag of the image added.
' ... The Default parameter is the key that will be returned if file info not found.

Dim sExt As String
Dim iIndex As Long
Dim xfInfo As FileNameInfo

    On Error GoTo ErrHan:
        
    ' ... set the Default return key to the icon.
    AddIconToImageList = pTheDefault
    prFileInfo = vbNullChar
    
    ' ... split the file name into its component parts
    ' ... and use the extension only ( xfInfo.Extension ).
    modFileName.ParseFileNameEx pTheFileName, xfInfo
'    modFileName.ShredFileName pTheFileName, xfInfo
    
    If Len(xfInfo.Extension) Then
        
        sExt = "." & UCase$(xfInfo.Extension)
        
        On Error Resume Next
        iIndex = pTheImageList.ListImages(sExt).Index
        
        If Err.Number <> 0 Then
            
            Err.Clear
            
            On Error GoTo ErrHan:
            pTheImageList.ListImages.Add , sExt, pGetFileIconSmall(sExt)
            
            prFileInfo = pGetFileTypeInfo(sExt)
            pTheImageList.ListImages(sExt).Tag = prFileInfo
            
        Else
            
            prFileInfo = pTheImageList.ListImages(sExt).Tag
        
        End If
        
        ' ... update the actual icon used for return.
        AddIconToImageList = sExt
        
    End If
       
Exit Function
ErrHan:

End Function

