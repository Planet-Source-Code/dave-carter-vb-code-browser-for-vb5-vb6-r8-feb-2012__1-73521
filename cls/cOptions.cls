VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class dedicated to reading and saving various user preferences to this program."
' what?
'  a class to read and write user preferred settings from/to file.
' why?
'   it's the done thing.
' when?
'   this class is dedicated to this program...
'   use it to load user preferences and save them.
' how?
'   Load Settings:
'   Dim xOptions As cOptions
'       Set xOptions = New cOptions
'       xOptions.Read
'       ' ... user preferences are available via properties
'       ' ... e.g.
'       sFontName = xOptions.FontName
'       sFontSize = xOptions.FontSize
'       ' ... etc.
'
'   Store Settings:
'   Dim xOptions As cOptions
'       Set xOptions = New cOptions
'       xOptions.FontName = "Courier New"
'       xOptions.FontSize = "10"
'       ' ... etc.
'       xOptions.Save
' who?
'  d.c.

Option Explicit

' Note:
'   This module reads and writes to a file named Options.dat which will be created
'   in the root folder to the program.
'   These settings could just as easily be stored in the Windows Registry but
'   I prefer to use a file because it makes the info saved transparent to the owner.
'   If you no like then just comment the ' xSave.ToFile mFileName, , , , pOK, pErrMsg ' code line in Save,
'   the worst that will happen is that default settings will always be applied.

' S   Read.              ' ... Reads contents of Options file and parses data into various properties.
' S   Save.              ' ... Saves the Option Property Values to file.

' P   AttributeTextColour: Long.  ' ... Sets / Returns the User's Viewer Attribute Text Colour.
' P   AutoCleanUnzipFolder: Boolean.  ' ... Sets / Returns a Boolean describing whether the Unzip folder is emptied when the program closes.
' P   AutoLoadAllCode: Boolean.  ' ... Sets / Returns a Boolean describing whether the complete content of a Code File is loaded into the code viewer when selected from the Project Explorer (True = Yes / False = No).
' P   AutoLoadProjects: Boolean.  ' ... Sets / Returns a Boolean describing whether projects in zipped files are automatically loaded (true) or not (false) after unzipping.
' P   AutoRTFEncoding: Boolean.  ' ... Sets / Returns a Boolean describing whether Code Text is Automatically converted to RTF when loaded into the code viewer (True = Yes / False = No).
' P   AutoUnZip: Boolean.  ' ... Sets / Returns a Boolean describing whether zipped files are automatically unzipped (true) or not (false).
' P   CommentTextColour: Long.  ' ... Sets / Returns the User's Viewer Comment Text Colour.
' P   ConfirmExit: Boolean.  ' ... Option to confirm exit program, true, or not, false (ie just quit).
' P   FontBold: Boolean.  ' ... Sets / Returns the whether the User's Viewer writes Text in Bold (True) or Not (False).
' P   FontName: String.  ' ... Sets / Returns which Font Name is used to write the text in the viewer.
' P   FontSize: String.  ' ... Sets / Returns the Font Size used to write the text in the viewer.
' P   HideChildProject: Boolean.  ' ... Option to hide the project explorer when window opened as a child of a project, true, or not, ie show project explorer.
' P   HideChildToolbar: Boolean.  ' ... Option to hide the toolbar when window opened as a child of a project, true, or not, ie show toolbar.
' P   KeywordTextColour: Long.  ' ... Sets / Returns the User's Viewer Keyword Text Colour.
' P   LineNoTextColour: Long.  ' ... Sets / Returns the User's Viewer Line Number Text Colour.
' P   LineNumbers: Boolean.  ' ... Sets / Returns the whether the User's Viewer includes Line Number Text.
' P   NormalTextColour: Long.  ' ... Sets / Returns the User's Viewer Normal Text Colour.
' P   PathToTextEditor: String.  ' ... Sets / Returns the path to the preferred  text editor.
' P   PathToVB432: String.  ' ... Sets / Returns the path to the VB32 IDE.
' P   PathToVB5: String.  ' ... Sets / Returns the path to the VB5 IDE.
' P   PathToVB6: String.  ' ... Sets / Returns the path to the VB6 IDE.
' P   ShowAttributes: Boolean.  ' ... Show / Hide Attributes in encoded text.
' P   ShowVB432IDE: Boolean.  ' ... Sets / Returns whether the VB4 32 IDE Option is available in the Project Explorer Pop-Up Menu.
' P   ShowClassHeadCount: Boolean.  ' ... Option to Show Count of Items on Class Explorer, yes = true, no = false.
' P   ShowTextEditor: Boolean.  ' ... Sets / Returns whether the Text Editor Option is available in the Project Explorer Pop-Up Menu.
' P   ShowVB5IDE: Boolean.  ' ... Sets / Returns whether the VB5 IDE Option is available in the Project Explorer Pop-Up Menu.
' P   ShowVB6IDE: Boolean.  ' ... Sets / Returns whether the VB6 IDE Option is available in the Project Explorer Pop-Up Menu.
' P   UnZipFolder: String.  ' ... Sets / |Returns the default folder for unzipping files into.
' P   UseChildWindows: Boolean.  ' ... Option to open new files in a new window, true, or not false (ie just keep to a single window per project).
' P   UseOwnColours: Boolean.  ' ... Sets / Returns a Boolean describing if User Colours (True) are used or Default Colours (False).
' P   ViewerBackColour: Long.  ' ... Sets / Returns the User's Viewer BackColour.

Private m_AttributeTextColour As Long ' ... private field for property AttributeTextColour.
Private m_AutoCleanUnzipFolder As Boolean ' ... private field for property AutoCleanUnzipFolder.
Private m_AutoLoadAllCode As Boolean ' ... private field for property AutoLoadAllCode.
Private m_AutoRTFEncoding As Boolean ' ... private field for property AutoRTFEncoding.
Private m_AutoLoadProjects As Boolean ' ... private field for property AutoLoadProjects.
Private m_AutoUnZip As Boolean ' ... private field for property AutoUnZip.
Private m_CommentTextColour As Long ' ... private field for property CommentTextColour.
Private m_ConfirmExit As Boolean ' ... private field for property ConfirmExit.
Private m_FontBold As Boolean ' ... private field for property FontBold.
Private m_FontName As String ' ... private field for property FontName.
Private m_FontSize As String ' ... private field for property FontSize.
Private m_HideChildProject As Boolean ' ... private field for property HideChildProject.
Private m_HideChildToolbar As Boolean ' ... private field for property HideChildToolbar.
Private m_KeywordTextColour As Long ' ... private field for property KeywordTextColour.
Private m_LineNoTextColour As Long ' ... private field for property LineNoTextColour.
Private m_LineNumbers As Boolean ' ... private field for property LineNumbers.
Private m_NormalTextColour As Long ' ... private field for property NormalTextColour.
Private m_PathToTextEditor As String ' ... private field for property PathToTextEditor.
Private m_PathToVB432 As String ' ... private field for property PathToVB432.
Private m_PathToVB5 As String ' ... private field for property PathToVB5.
Private m_PathToVB6 As String ' ... private field for property PathToVB6.
Private m_ShowAttributes As Boolean ' ... private field for property ShowAttributes.
Private m_ShowClassHeadCount As Boolean ' ... private field for property ShowClassHeadCount.
Private m_ShowTextEditor As Boolean ' ... private field for property ShowTextEditor.
Private m_ShowVB432IDE As Boolean ' ... private field for property ShowVB432IDE.
Private m_ShowVB5IDE As Boolean ' ... private field for property ShowVB5IDE.
Private m_ShowVB6IDE As Boolean ' ... private field for property ShowVB6IDE.
Private m_UnzipFolder As String ' ... private field for property UnZipFolder.
Private m_UseChildWindows As Boolean ' ... private field for property UseChildWindows.
Private m_UseOwnColours As Boolean ' ... private field for property UseOwnColours.
Private m_ViewerBackColour As Long ' ... private field for property ViewerBackColour.

Private mFileName As String
Private Const cFileName As String = "Options.dat"

Private Const cConfirmExit As String = "ConfirmExit"
Private Const cHideChildToolbar As String = "HideChildToolbar"
Private Const cHideChildProject As String = "HideChildProject"
Private Const cUseChildWindows As String = "UseChildWindows"
Private Const cShowClassHeadCount As String = "ShowClassHeadCount"

Private Const cUseOwnColours As String = "UseOwnColours"
Private Const cViewerBackColour As String = "ViewerBackColour"
Private Const cNormalTextColour As String = "NormalTextColour"
Private Const cKeyWordTextColour As String = "KeyWordTextColour"
Private Const cCommentTextColour As String = "CommentTextColour"
Private Const cAttributeTextColour As String = "AttributeTextColour"
Private Const cLineNoTextColour As String = "LineNoTextColour"
Private Const cLineNos As String = "LineNos"
Private Const cFontSize As String = "FontSize"
Private Const cFontName As String = "FontName"
Private Const cFontBold As String = "FontBold"
Private Const cShowTextEditor As String = "ShowTextEditor"
Private Const cShowVB432 As String = "ShowVB432"
Private Const cShowVB5 As String = "ShowVB5"
Private Const cShowVB6 As String = "ShowVB6"
Private Const cAutoUnzip As String = "AutoUnZip"
Private Const cAutoLoadVBP As String = "AutoLoadProject"
Private Const cUnZipFolder As String = "UnZipFolder"
Private Const cAutoCleanUnzip As String = "AutoCleanUnZipFolder"
Private Const cShowAttributes As String = "ShowAttributes"

Private Const cAutoRTFEncoding As String = "AutoRTFEncoding"
Private Const cAutoLoadAllCode As String = "AutoLoadAllCode"

' ... default paths to text editor, vb5 and vb6 exes.
Private Const cAppTextEditor As String = "C:\Windows\system32\notepad.exe"
Private Const cAppVB4IDE As String = "C:\Program Files\Microsoft Visual Basic\vb32.exe"
Private Const cAppVB5IDE As String = "C:\Program Files\DevStudio\VB\VB5.exe"
Private Const cAppVB6IDE As String = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe"

'Private Const c_def_UnZipFolder As String = "UnZipFiles" ' ... default value for property UnZipFolder.

' ... v6
Public Property Get AutoLoadAllCode() As Boolean
Attribute AutoLoadAllCode.VB_Description = "Sets / Returns a Boolean describing whether the complete content of a Code File is loaded into the code viewer when selected from the Project Explorer (True = Yes / False = No)."

    Let AutoLoadAllCode = m_AutoLoadAllCode

End Property ' ... AutoLoadAllCode: Boolean.

Public Property Let AutoLoadAllCode(ByVal pNewValue As Boolean)

    Let m_AutoLoadAllCode = pNewValue

End Property ' ... AutoLoadAllCode: Boolean.

Public Property Get AutoRTFEncoding() As Boolean
Attribute AutoRTFEncoding.VB_Description = "Sets / Returns a Boolean describing whether Code Text is Automatically converted to RTF when loaded into the code viewer (True = Yes / False = No)."

    Let AutoRTFEncoding = m_AutoRTFEncoding

End Property ' ... AutoRTFEncoding: Boolean.

Public Property Let AutoRTFEncoding(ByVal pNewValue As Boolean)

    Let m_AutoRTFEncoding = pNewValue

End Property ' ... AutoRTFEncoding: Boolean.

' ... v5
Public Property Get ShowClassHeadCount() As Boolean
Attribute ShowClassHeadCount.VB_Description = "Option to Show Count of Items on Class Explorer, yes = true, no = false."

    Let ShowClassHeadCount = m_ShowClassHeadCount

End Property ' ... ShowClassHeadCount: Boolean.

Public Property Let ShowClassHeadCount(ByVal pNewValue As Boolean)

    Let m_ShowClassHeadCount = pNewValue

End Property ' ... ShowClassHeadCount: Boolean.

Public Property Get AttributeTextColour() As Long
Attribute AttributeTextColour.VB_Description = "Sets / Returns the User's Viewer Attribute Text Colour."

    Let AttributeTextColour = m_AttributeTextColour

End Property ' ... AttributeTextColour: Long.

Public Property Let AttributeTextColour(ByVal pNewValue As Long)

    Let m_AttributeTextColour = pNewValue

End Property ' ... AttributeTextColour: Long.

Public Property Get AutoCleanUnzipFolder() As Boolean
Attribute AutoCleanUnzipFolder.VB_Description = "Sets / Returns a Boolean describing whether the Unzip folder is emptied when the program closes."

    Let AutoCleanUnzipFolder = m_AutoCleanUnzipFolder

End Property ' ... AutoCleanUnzipFolder: Boolean.

Public Property Let AutoCleanUnzipFolder(ByVal pNewValue As Boolean)

    Let m_AutoCleanUnzipFolder = pNewValue

End Property ' ... AutoCleanUnzipFolder: Boolean.

Public Property Get AutoLoadProjects() As Boolean
Attribute AutoLoadProjects.VB_Description = "Sets / Returns a Boolean describing whether projects in zipped files are automatically loaded (true) or not (false) after unzipping."

    Let AutoLoadProjects = m_AutoLoadProjects

End Property ' ... AutoLoadProjects: Boolean.

Public Property Let AutoLoadProjects(ByVal pNewValue As Boolean)

    Let m_AutoLoadProjects = pNewValue

End Property ' ... AutoLoadProjects: Boolean.

Public Property Get AutoUnZip() As Boolean
Attribute AutoUnZip.VB_Description = "Sets / Returns a Boolean describing whether zipped files are automatically unzipped (true) or not (false)."

    Let AutoUnZip = m_AutoUnZip

End Property ' ... AutoUnZip: Boolean.

Public Property Let AutoUnZip(ByVal pNewValue As Boolean)

    Let m_AutoUnZip = pNewValue

End Property ' ... AutoUnZip: Boolean.

Private Sub Class_Initialize()
    mFileName = App.Path & "\" & cFileName
End Sub

Public Property Get CommentTextColour() As Long
Attribute CommentTextColour.VB_Description = "Sets / Returns the User's Viewer Comment Text Colour."

    Let CommentTextColour = m_CommentTextColour

End Property ' ... CommentTextColour: Long.

Public Property Let CommentTextColour(ByVal pNewValue As Long)

    Let m_CommentTextColour = pNewValue

End Property ' ... CommentTextColour: Long.

Public Property Get ConfirmExit() As Boolean
Attribute ConfirmExit.VB_Description = "Option to confirm exit program, true, or not, false (ie just quit)."

    Let ConfirmExit = m_ConfirmExit

End Property ' ... ConfirmExit: Boolean.

Public Property Let ConfirmExit(ByVal pNewValue As Boolean)

    Let m_ConfirmExit = pNewValue

End Property ' ... ConfirmExit: Boolean.

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Sets / Returns the whether the User's Viewer writes Text in Bold (True) or Not (False)."

    Let FontBold = m_FontBold

End Property ' ... FontBold: Boolean.

Public Property Let FontBold(ByVal pNewValue As Boolean)

    Let m_FontBold = pNewValue

End Property ' ... FontBold: Boolean.

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Sets / Returns which Font Name is used to write the text in the viewer."

    Let FontName = m_FontName

End Property ' ... FontName: String.

Public Property Let FontName(ByVal pNewValue As String)

    Let m_FontName = pNewValue

End Property ' ... FontName: String.

Public Property Get FontSize() As String
Attribute FontSize.VB_Description = "Sets / Returns the Font Size used to write the text in the viewer."

    Let FontSize = m_FontSize

End Property ' ... FontSize: String.

Public Property Let FontSize(ByVal pNewValue As String)

    Let m_FontSize = pNewValue

End Property ' ... FontSize: String.

Public Property Get HideChildProject() As Boolean
Attribute HideChildProject.VB_Description = "Option to hide the project explorer when window opened as a child of a project, true, or not, ie show project explorer."

    Let HideChildProject = m_HideChildProject

End Property ' ... HideChildProject: Boolean.

Public Property Let HideChildProject(ByVal pNewValue As Boolean)

    Let m_HideChildProject = pNewValue

End Property ' ... HideChildProject: Boolean.

Public Property Get HideChildToolbar() As Boolean
Attribute HideChildToolbar.VB_Description = "Option to hide the toolbar when window opened as a child of a project, true, or not, ie show toolbar."

    Let HideChildToolbar = m_HideChildToolbar

End Property ' ... HideChildToolbar: Boolean.

Public Property Let HideChildToolbar(ByVal pNewValue As Boolean)

    Let m_HideChildToolbar = pNewValue

End Property ' ... HideChildToolbar: Boolean.

Public Property Get KeywordTextColour() As Long
Attribute KeywordTextColour.VB_Description = "Sets / Returns the User's Viewer Keyword Text Colour."

    Let KeywordTextColour = m_KeywordTextColour

End Property ' ... KeywordTextColour: Long.

Public Property Let KeywordTextColour(ByVal pNewValue As Long)

    Let m_KeywordTextColour = pNewValue

End Property ' ... KeywordTextColour: Long.

Public Property Get LineNoTextColour() As Long
Attribute LineNoTextColour.VB_Description = "Sets / Returns the User's Viewer Line Number Text Colour."

    Let LineNoTextColour = m_LineNoTextColour

End Property ' ... LineNoTextColour: Long.

Public Property Let LineNoTextColour(ByVal pNewValue As Long)

    Let m_LineNoTextColour = pNewValue

End Property ' ... LineNoTextColour: Long.

Public Property Get LineNumbers() As Boolean
Attribute LineNumbers.VB_Description = "Sets / Returns the whether the User's Viewer includes Line Number Text."

    Let LineNumbers = m_LineNumbers

End Property ' ... LineNumbers: Boolean.

Public Property Let LineNumbers(ByVal pNewValue As Boolean)

    Let m_LineNumbers = pNewValue

End Property ' ... LineNumbers: Boolean.

Public Property Get NormalTextColour() As Long
Attribute NormalTextColour.VB_Description = "Sets / Returns the User's Viewer Normal Text Colour."

    Let NormalTextColour = m_NormalTextColour

End Property ' ... NormalTextColour: Long.

Public Property Let NormalTextColour(ByVal pNewValue As Long)

    Let m_NormalTextColour = pNewValue

End Property ' ... NormalTextColour: Long.

Public Property Get PathToTextEditor() As String
Attribute PathToTextEditor.VB_Description = "Sets / Returns the path to the preferred  text editor."

    Let PathToTextEditor = m_PathToTextEditor

End Property ' ... PathToTextEditor: String.

Public Property Let PathToTextEditor(ByVal pNewValue As String)

    Let m_PathToTextEditor = pNewValue

End Property ' ... PathToTextEditor: String.

Public Property Get PathToVB432() As String
Attribute PathToVB432.VB_Description = "Sets / Returns the path to the VB32 IDE."

    Let PathToVB432 = m_PathToVB432

End Property ' ... PathToVB432: String.

Public Property Let PathToVB432(ByVal pNewValue As String)

    Let m_PathToVB432 = pNewValue

End Property ' ... PathToVB432: String.

Public Property Get PathToVB5() As String
Attribute PathToVB5.VB_Description = "Sets / Returns the path to the VB5 IDE."

    Let PathToVB5 = m_PathToVB5

End Property ' ... PathToVB5: String.

Public Property Let PathToVB5(ByVal pNewValue As String)

    Let m_PathToVB5 = pNewValue

End Property ' ... PathToVB5: String.

Public Property Get PathToVB6() As String
Attribute PathToVB6.VB_Description = "Sets / Returns the path to the VB6 IDE."

    Let PathToVB6 = m_PathToVB6

End Property ' ... PathToVB6: String.

Public Property Let PathToVB6(ByVal pNewValue As String)

    Let m_PathToVB6 = pNewValue

End Property ' ... PathToVB6: String.

Public Sub Read(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Read.VB_Description = "Reads contents of Options file and parses data into various properties."

' Sub:             Read
' Description:     Reads contents of Options file and parses data into various properties.

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim sOptionsText As String
Dim sOptions() As String
Dim lngCount As Long
Dim lngLoop As Long
Dim sTmp As String
Dim sLeft As String
Dim sRight As String

    On Error GoTo ErrHan:
    
    ' ... set up default underlying values to properties, regardless.
    m_ViewerBackColour = &HFFFFFF
    m_NormalTextColour = &H0
    m_KeywordTextColour = RGB(0, 0, 128)
    m_CommentTextColour = RGB(0, 128, 0)
    m_AttributeTextColour = RGB(128, 128, 128)
    m_LineNoTextColour = RGB(148, 148, 148)
    m_LineNumbers = False
    m_FontName = "Courier New"
    m_FontSize = "09"

    m_ShowTextEditor = True
    m_ShowVB432IDE = False 'True
    m_ShowVB5IDE = True
    m_ShowVB6IDE = True
    
    m_PathToTextEditor = cAppTextEditor
    m_PathToVB432 = cAppVB4IDE
    m_PathToVB5 = cAppVB5IDE
    m_PathToVB6 = cAppVB6IDE
    
    m_AutoUnZip = False
    m_AutoLoadProjects = False
    m_AutoCleanUnzipFolder = False ' ... v6, changed to False so user is more aware of having set this up.
    
    m_ShowAttributes = True     ' ... v7, amended from false to true.
    
    m_UseChildWindows = False
    m_ConfirmExit = True
    m_ShowClassHeadCount = True ' ... v5
    m_AutoRTFEncoding = True    ' ... v6
    m_AutoLoadAllCode = False   ' ... v6
    
    m_UnzipFolder = App.Path & "\" & c_def_UnZipFolder
    
    sOptionsText = modReader.ReadFile(mFileName, pOK, pErrMsg)
    
    If pOK = True Then
        pOK = Len(sOptionsText) > 0
    Else ' ... v7, options not yet written, fed up of debug printing error file not found.
        WriteFile mFileName, ""
    End If
    
    If pOK = True Then
        
        modStringArrays.SplitString sOptionsText, sOptions, vbCrLf, lngCount, pOK, pErrMsg
        
        If pOK And lngCount > 0 Then
                    
            For lngLoop = 0 To lngCount - 1
            
                sTmp = sOptions(lngLoop)
                
                modStrings.SplitStringPair sTmp, "=", sLeft, sRight, True, True
                
                Select Case sLeft
                    
                    Case cUseOwnColours:        UseOwnColours = CBool(Val(sRight))
                    Case cViewerBackColour:     ViewerBackColour = CLng(Val(sRight))
                    Case cNormalTextColour:     NormalTextColour = CLng(Val(sRight))
                    Case cKeyWordTextColour:    KeywordTextColour = CLng(Val(sRight))
                    Case cCommentTextColour:    CommentTextColour = CLng(Val(sRight))
                    Case cAttributeTextColour:  AttributeTextColour = CLng(Val(sRight))
                    Case cLineNoTextColour:     LineNoTextColour = CLng(Val(sRight))
                    Case cLineNos:              LineNumbers = CBool(Val(sRight))
                    Case cFontSize:             FontSize = IIf(Len(sRight), sRight, m_FontSize)
                    Case cFontName:             FontName = IIf(Len(sRight), sRight, m_FontName)
                    Case cFontBold:             FontBold = CBool(Val(sRight))
                    Case cAutoUnzip:            AutoUnZip = CBool(Val(sRight))
                    Case cAutoLoadVBP:          AutoLoadProjects = CBool(Val(sRight))
                    Case cUnZipFolder:
                        ' ... v6, not allowed to return app.path.
                        ' ... can happen if options file is copied to
                        ' ... a different folder.
                        UnzipFolder = sRight
                    Case cAutoCleanUnzip:       AutoCleanUnzipFolder = CBool(Val(sRight))
                    Case cShowAttributes:       ShowAttributes = CBool(Val(sRight))
                    Case cConfirmExit:          ConfirmExit = CBool(Val(sRight))
                    Case cHideChildProject:     HideChildProject = CBool(Val(sRight))
                    Case cHideChildToolbar:     HideChildToolbar = CBool(Val(sRight))
                    Case cUseChildWindows:      UseChildWindows = CBool(Val(sRight))
                    Case cShowClassHeadCount:   ShowClassHeadCount = CBool(Val(sRight))
                    Case cAutoRTFEncoding:      AutoRTFEncoding = CBool(Val(sRight))
                    Case cAutoLoadAllCode:      AutoLoadAllCode = CBool(Val(sRight))
                    
                    Case cShowTextEditor:       ShowTextEditor = CBool(Val(sRight))
                                                sTmp = Mid$(sRight, 2)
                                                If Len(sTmp) Then
                                                    PathToTextEditor = sTmp
                                                End If
                                                
                    Case cShowVB432:            ShowVB432IDE = CBool(Val(sRight))
                                                sTmp = Mid$(sRight, 2)
                                                If Len(sTmp) Then
                                                    PathToVB432 = sTmp
                                                End If
                                                
                    Case cShowVB5:              ShowVB5IDE = CBool(Val(sRight))
                                                sTmp = Mid$(sRight, 2)
                                                If Len(sTmp) Then
                                                    PathToVB5 = sTmp
                                                End If
                                                
                    Case cShowVB6:              ShowVB6IDE = CBool(Val(sRight))
                                                sTmp = Mid$(sRight, 2)
                                                If Len(sTmp) Then
                                                    PathToVB6 = sTmp
                                                End If
                    
                End Select
            
            Next lngLoop
        
        End If
        
    End If

    Let pErrMsg = vbNullString
    Let pOK = True

ErrResume:

    On Error GoTo 0
    
    Erase sOptions
    
    sOptionsText = vbNullString
    sTmp = vbNullString
    sLeft = vbNullString
    sRight = vbNullString
    
Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "cOptions.Read", Err.Number, Err.Description
    Resume ErrResume:

End Sub ' ... Read.

Public Sub Save(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Save.VB_Description = "Saves the Option Property Values to file."

' Sub:             Save
' Description:     Saves the Option Property Values to file.

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim xSave As SBuilder ' StringWorker

    On Error GoTo ErrHan:
    
    Set xSave = New SBuilder ' StringWorker
    
    xSave.AppendAsLine cUseOwnColours & "=" & IIf(UseOwnColours, "1", "0")
    xSave.AppendAsLine cViewerBackColour & "=" & CStr(ViewerBackColour)
    xSave.AppendAsLine cNormalTextColour & "=" & CStr(NormalTextColour)
    xSave.AppendAsLine cKeyWordTextColour & "=" & CStr(KeywordTextColour)
    xSave.AppendAsLine cCommentTextColour & "=" & CStr(CommentTextColour)
    xSave.AppendAsLine cAttributeTextColour & "=" & CStr(AttributeTextColour)
    xSave.AppendAsLine cLineNoTextColour & "=" & CStr(LineNoTextColour)
    xSave.AppendAsLine cLineNos & "=" & IIf(LineNumbers, "1", "0")
    xSave.AppendAsLine cFontSize & "=" & FontSize
    xSave.AppendAsLine cFontName & "=" & FontName
    xSave.AppendAsLine cFontBold & "=" & IIf(FontBold, "1", "0")
    
    xSave.AppendAsLine cShowTextEditor & "=" & IIf(ShowTextEditor, "1", "0") & m_PathToTextEditor
    xSave.AppendAsLine cShowVB432 & "=" & IIf(ShowVB432IDE, "1", "0") & m_PathToVB432
    xSave.AppendAsLine cShowVB5 & "=" & IIf(ShowVB5IDE, "1", "0") & m_PathToVB5
    xSave.AppendAsLine cShowVB6 & "=" & IIf(ShowVB6IDE, "1", "0") & m_PathToVB6
    
    xSave.AppendAsLine cAutoUnzip & "=" & IIf(AutoUnZip, "1", "0")
    xSave.AppendAsLine cAutoLoadVBP & "=" & IIf(AutoLoadProjects, "1", "0")
    xSave.AppendAsLine cAutoCleanUnzip & "=" & IIf(AutoCleanUnzipFolder, "1", "0")
    
    xSave.AppendAsLine cShowAttributes & "=" & IIf(ShowAttributes, "1", "0")
    
    xSave.AppendAsLine cConfirmExit & "=" & IIf(ConfirmExit, "1", "0")
    xSave.AppendAsLine cHideChildProject & "=" & IIf(HideChildProject, "1", "0")
    xSave.AppendAsLine cHideChildToolbar & "=" & IIf(HideChildToolbar, "1", "0")
    xSave.AppendAsLine cUseChildWindows & "=" & IIf(UseChildWindows, "1", "0")
    xSave.AppendAsLine cShowClassHeadCount & "=" & IIf(ShowClassHeadCount, "1", "0")
    xSave.AppendAsLine cAutoRTFEncoding & "=" & IIf(AutoRTFEncoding, "1", "0")      ' ... v6
    xSave.AppendAsLine cAutoLoadAllCode & "=" & IIf(AutoLoadAllCode, "1", "0")      ' ... v6
    
    xSave.AppendAsLine cUnZipFolder & "=" & m_UnzipFolder
    
    xSave.WriteToFile mFileName, , pOK, pErrMsg
    
ErrResume:
    
    On Error GoTo 0
    Set xSave = Nothing
    
Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "cOptions.Save", Err.Number, Err.Description
    Resume ErrResume:

End Sub ' ... Save.

Public Property Get ShowAttributes() As Boolean
Attribute ShowAttributes.VB_Description = "Show / Hide Attributes in encoded text."

    Let ShowAttributes = m_ShowAttributes

End Property ' ... ShowAttributes: Boolean.

Public Property Let ShowAttributes(ByVal pNewValue As Boolean)

    Let m_ShowAttributes = pNewValue

End Property ' ... ShowAttributes: Boolean.

Public Property Get ShowTextEditor() As Boolean
Attribute ShowTextEditor.VB_Description = "Sets / Returns whether the Text Editor Option is available in the Project Explorer Pop-Up Menu."

    Let ShowTextEditor = m_ShowTextEditor

End Property ' ... ShowTextEditor: Boolean.

Public Property Let ShowTextEditor(ByVal pNewValue As Boolean)

    Let m_ShowTextEditor = pNewValue

End Property ' ... ShowTextEditor: Boolean.

Public Property Get ShowVB432IDE() As Boolean
Attribute ShowVB432IDE.VB_Description = "Sets / Returns whether the VB4 32 IDE Option is available in the Project Explorer Pop-Up Menu."

    Let ShowVB432IDE = m_ShowVB432IDE

End Property ' ... ShowVB432IDE: Boolean.

Public Property Let ShowVB432IDE(ByVal pNewValue As Boolean)

    Let m_ShowVB432IDE = pNewValue

End Property ' ... ShowVB432IDE: Boolean.

Public Property Get ShowVB5IDE() As Boolean
Attribute ShowVB5IDE.VB_Description = "Sets / Returns whether the VB5 IDE Option is available in the Project Explorer Pop-Up Menu."

    Let ShowVB5IDE = m_ShowVB5IDE

End Property ' ... ShowVB5IDE: Boolean.

Public Property Let ShowVB5IDE(ByVal pNewValue As Boolean)

    Let m_ShowVB5IDE = pNewValue

End Property ' ... ShowVB5IDE: Boolean.

Public Property Get ShowVB6IDE() As Boolean
Attribute ShowVB6IDE.VB_Description = "Sets / Returns whether the VB6 IDE Option is available in the Project Explorer Pop-Up Menu."

    Let ShowVB6IDE = m_ShowVB6IDE

End Property ' ... ShowVB6IDE: Boolean.

Public Property Let ShowVB6IDE(ByVal pNewValue As Boolean)

    Let m_ShowVB6IDE = pNewValue

End Property ' ... ShowVB6IDE: Boolean.

Public Property Get UnzipFolder() As String
Attribute UnzipFolder.VB_Description = "Sets / Returns the default folder for unzipping files into."

    ' ... v6, update, not allowed to return app.path.
    ' ... don't want to delete the program folder itself.
    
    If UCase$(App.Path) = UCase$(m_UnzipFolder) Then
        m_UnzipFolder = App.Path & "\" & c_def_UnZipFolder
    End If

    Let UnzipFolder = m_UnzipFolder

End Property ' ... UnZipFolder: String.

Public Property Let UnzipFolder(ByVal pNewValue As String)
    
    If Len(Trim$(pNewValue)) = 0 Then Exit Property
    ' -------------------------------------------------------------------
    ' ... v6, update, not allowed to be app.path.
    If UCase$(App.Path) = UCase$(pNewValue) Then
        pNewValue = App.Path & "\" & c_def_UnZipFolder
    End If
    ' -------------------------------------------------------------------
    Let m_UnzipFolder = pNewValue

End Property ' ... UnZipFolder: String.

Public Property Get UseChildWindows() As Boolean
Attribute UseChildWindows.VB_Description = "Option to open new files in a new window, true, or not false (ie just keep to a single window per project)."

    Let UseChildWindows = m_UseChildWindows

End Property ' ... UseChildWindows: Boolean.

Public Property Let UseChildWindows(ByVal pNewValue As Boolean)

    Let m_UseChildWindows = pNewValue

End Property ' ... UseChildWindows: Boolean.

Public Property Get UseOwnColours() As Boolean
Attribute UseOwnColours.VB_Description = "Sets / Returns a Boolean describing if User Colours (True) are used or Default Colours (False)."

    Let UseOwnColours = m_UseOwnColours

End Property ' ... UseOwnColours: Boolean.

Public Property Let UseOwnColours(ByVal pNewValue As Boolean)

    Let m_UseOwnColours = pNewValue

End Property ' ... UseOwnColours: Boolean.

Public Property Get ViewerBackColour() As Long
Attribute ViewerBackColour.VB_Description = "Sets / Returns the User's Viewer BackColour."

    Let ViewerBackColour = m_ViewerBackColour

End Property ' ... ViewerBackColour: Long.

Public Property Let ViewerBackColour(ByVal pNewValue As Long)

    Let m_ViewerBackColour = pNewValue

End Property ' ... ViewerBackColour: Long.

