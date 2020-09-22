VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "VBCB R6c"
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12810
   Icon            =   "mdiMainCV.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnzip 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1140
      Top             =   3240
   End
   Begin ComctlLib.ImageList liMember 
      Left            =   1230
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   26
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":058A
            Key             =   "main"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":0ADC
            Key             =   "declarations"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":102E
            Key             =   "event"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":1580
            Key             =   "constpublic"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":18D2
            Key             =   "constprivate"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":1E24
            Key             =   "type"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":2376
            Key             =   "enum"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":28C8
            Key             =   "apipublic"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":29DA
            Key             =   "apiprivate"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":2F2C
            Key             =   "implements"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":347E
            Key             =   "subpublic"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":37D0
            Key             =   "subprivate"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":3D22
            Key             =   "subfriend"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":4074
            Key             =   "funcpublic"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":43C6
            Key             =   "funcprivate"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":4918
            Key             =   "funcfriend"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":4C6A
            Key             =   "proppublic"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":4FBC
            Key             =   "propprivate"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":550E
            Key             =   "propfriend"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":5860
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":5DB2
            Key             =   "info"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":6304
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":6856
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":6DA8
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":72FA
            Key             =   "User Control"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":784C
            Key             =   "folders"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilProject 
      Left            =   1260
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":795E
            Key             =   "notfound"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":7EB0
            Key             =   "project"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":8402
            Key             =   "reference"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":8954
            Key             =   "component"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":8EA6
            Key             =   "formmdi"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":93F8
            Key             =   "formmdichild"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":994A
            Key             =   "formnormal"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":9E9C
            Key             =   "class"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":A3EE
            Key             =   "module"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":A940
            Key             =   "uc"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":AE92
            Key             =   "resource"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":B3E4
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":B4F6
            Key             =   "db"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":B608
            Key             =   "info"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":BB5A
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMainCV.frx":BEAC
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Begin VB.Menu mnuNewWindow 
         Caption         =   ""
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
         Index           =   0
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   ""
      Begin VB.Menu mnuProjSearch 
         Caption         =   "Project Search"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefViewer 
         Caption         =   "Mini Reference Viewer"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsSep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUnzip 
         Caption         =   "UnZip"
      End
      Begin VB.Menu mnuFOpenZipFolder 
         Caption         =   "Open UnZip Folder"
      End
      Begin VB.Menu mnuUnzipPSC 
         Caption         =   "Scan for PSC Read Me Files"
      End
      Begin VB.Menu mnuPSCDownloads 
         Caption         =   "PSC Read Me files in Unzip Folder"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMgnHist 
         Caption         =   "Manage VBP History"
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompResource 
         Caption         =   "Manifest Resource Maker"
      End
      Begin VB.Menu mnuFileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyProject 
         Caption         =   "Copy Project"
      End
      Begin VB.Menu mnuCopyFile 
         Caption         =   "Copy File"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHTMLHelpGen 
         Caption         =   "HTML Member Pages"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   ""
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "Arrange"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   ""
      Begin VB.Menu mnuHelpFile 
         Caption         =   "Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowHelp 
         Caption         =   "Show Help"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTipODay 
         Caption         =   "Tip of the Day"
      End
      Begin VB.Menu mnuHSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "MDI Backdrop to viewer form/s."
Option Explicit

' ... Responsibilities:
'   ... Load New Viewer Forms.
'   ... Read/Save VBP History
'   ... Load VBP History menu.
'   ... Process OLE Drag'n'Drop.
'   ... Check and Process OLE Zip Drop.
'   ... Load other forms;
'       ... frmOptions.
'       ... frmMngVBPHistory.
'       ... frmUnZip.
'       ... frmAbout.

' ... Memory stuff.
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "PSAPI" (ByVal hProcess As Long) As Long

' -------------------------------------------------------------------
' ... v5/6 ... busking, open zip folder.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' -------------------------------------------------------------------

Private mChildCount As Long
Private moForms() As frmViewer
' -------------------------------------------------------------------
Private moFileHistory As StringArray
' -------------------------------------------------------------------
Private mOptions As cOptions
' -------------------------------------------------------------------
Private mDropZip As String
Private mHaveReadZip As Boolean
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
'
'
'
'Private Sub ClearMemory()
'Dim lngProcHandle As Long
'Dim lngRet As Long
'' -------------------------------------------------------------------
'' Helper: clears up application memory use.
'' -------------------------------------------------------------------
'    On Error GoTo LogMemory_Err
'    lngProcHandle = GetCurrentProcess()
'    lngRet = EmptyWorkingSet(lngProcHandle)
'LogMemory_Exit:
'Exit Sub
'LogMemory_Err:
'End Sub

Public Sub ChildLoaded()
    ClearMemory
End Sub

Public Sub ChildUnloaded(Index As Long)
    On Error GoTo ErrHan:
    ' ... could resize the array.
    ' ... could go for a collection for easier memory management e.g. resizing the array.
'    Set moForms(Index) = Nothing
ResumeError:
    ClearMemory
Exit Sub
ErrHan:
    Debug.Print "mdiMain.ChildUnloaded.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

Private Sub MDIForm_Initialize()
    InitCommonControlsVB
End Sub

Private Sub MDIForm_Load()
'Dim lngLoop As Long
'Dim lngCount As Long
'Dim sCaption As String
Dim bHaveUnzip As Boolean
Dim ShowAtStartup As Long
    

    On Error GoTo ErrHan:
    
    Screen.MousePointer = vbHourglass
    ' -------------------------------------------------------------------
    Caption = AppTitle
    ' -------------------------------------------------------------------
    pLoadHistMenu
    ' -------------------------------------------------------------------
    pLoadTextStrings
    ' -------------------------------------------------------------------
'    pLoadViewer
    ' -------------------------------------------------------------------
    bHaveUnzip = modGeneral.CheckUnzipAPI
    mnuUnzip.Enabled = bHaveUnzip
    ' -------------------------------------------------------------------
    modGeneral.CheckHelp
    ' -------------------------------------------------------------------
ResumeError:
    ' -------------------------------------------------------------------
    ' ... v3/4: mod level options.
    Set mOptions = New cOptions
    mOptions.Read
    ' -------------------------------------------------------------------
    ' See if we should be shown at startup
'    DeleteSetting App.EXEName, "Options"
    
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If ShowAtStartup = 1 Then
        frmTip.Show
'        BringWindowToTop frmTip.hwnd
'        frmTip.ZOrder
    End If
    
    ClearMemory
    
    Screen.MousePointer = vbDefault
Exit Sub
ErrHan:
    Debug.Print "mdiMain.MDIForm_Load.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

Private Sub pLoadHistMenu()
Dim sFile As String
Dim lngCount As Long
Dim lngMCount As Long
Dim lngLoop As Long
    
    ' ... clear out any existing hist menu items.
    lngMCount = mnuHistory.Count
    If lngMCount > 1 Then
        For lngLoop = 1 To lngMCount - 1
            Unload mnuHistory(lngLoop)
        Next lngLoop
    End If
    ' ... set up a new history string array
    If Not moFileHistory Is Nothing Then
        moFileHistory.Clear
        Set moFileHistory = Nothing
    End If
    
    Set moFileHistory = New StringArray
    moFileHistory.DuplicatesAllowed = False
    
    ' ... read the history file for the items to load.
    sFile = App.Path & "\History.dat"
    If Dir$(sFile, vbNormal) <> "" Then
        moFileHistory.FromFile App.Path & "\History.dat", vbCrLf
        lngCount = moFileHistory.Count
        If lngCount > 0 Then
            ' ... load the items found in the history file.
            For lngLoop = 1 To lngCount
                Load mnuHistory(lngLoop)
                mnuHistory(lngLoop).Caption = moFileHistory(lngLoop)
                mnuHistory(lngLoop).Visible = True
            Next lngLoop
        End If
    End If
    
    ' ... sort out history menu.
    mnuFileSep2.Visible = lngCount > 0
    mnuHistory(0).Visible = False
    
End Sub

Private Sub pLoadTextStrings()
    On Error GoTo ErrHan:
    mnuFile.Caption = LoadResString(128)
    mnuNewWindow.Caption = LoadResString(129)
    mnuExit.Caption = LoadResString(130)
    mnuWindow.Caption = LoadResString(131)
    mnuHelp.Caption = LoadResString(133)
'    mnuHelpFile.Caption = LoadResString(134)
    mnuHAbout.Caption = LoadResString(135)
    mnuTools.Caption = LoadResString(165)
Exit Sub
ErrHan:
    Debug.Print "mdiMain.pLoadTextStrings.Error: " & Err.Number & "; " & Err.Description
    Err.Clear
    Resume Next
End Sub

Sub LoadFile(pTheFile As String, _
    Optional pKey As String = vbNullString, _
    Optional pHideProject As Boolean = False, _
    Optional pHideToolbar As Boolean = False)
    
    pLoadViewer pTheFile, pKey, pHideProject, pHideToolbar
    
End Sub

Private Sub pLoadViewer(Optional pTheFile As String = vbNullString, _
                        Optional pKey As String = vbNullString, _
                        Optional pHideProject As Boolean = False, _
                        Optional pHideToolbar As Boolean = False, _
                        Optional ByVal pPSCReadMeText As String = vbNullString)
                        
Dim xF As frmViewer
Dim i As VBRUN.MousePointerConstants
Dim lngLoop As Long

    If Len(pKey) Then
        If mChildCount > 0 Then
            For lngLoop = 0 To mChildCount - 1
                If Not moForms(lngLoop) Is Nothing Then
                    If moForms(lngLoop).CodeFileName = pKey Then
                        moForms(lngLoop).ZOrder
                        Exit Sub
                    End If
                End If
            Next lngLoop
        End If
    End If

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    i = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    ' -------------------------------------------------------------------
    Set xF = New frmViewer
    xF.Hide
    ' -------------------------------------------------------------------
    ReDim Preserve moForms(mChildCount)
    Set moForms(mChildCount) = xF
    ' -------------------------------------------------------------------
    moForms(mChildCount).ChildIndex = mChildCount

    If pTheFile <> vbNullString Then
        moForms(mChildCount).LoadVBP pTheFile, pKey, pHideProject, pHideToolbar, pPSCReadMeText
    End If

    moForms(mChildCount).Show

    mChildCount = mChildCount + 1
    
'    frmPB.Show
'    frmPB.LoadVBP pTheFile
'    frmProject2.Show
'    frmProject2.LoadVBP pTheFile


ResumeError:
    If i <> Screen.MousePointer Then
        Screen.MousePointer = i
    End If
Exit Sub
ErrHan:
    Debug.Print "mdiMain.pLoadViewer.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim stext As String
Dim sFile As String
Dim xVBP As VBPInfo
Dim xArray As StringArray
Dim lngLoop As Long
'Dim lngCount As Long

    On Error GoTo ErrHan:
    
    If Data.GetFormat(vbCFText) Then
        stext = Data.GetData(1)
        If Left$(stext, Len(cFileSig)) = cFileSig Then
            stext = Mid$(stext, Len(cFileSig) + 1)
            ' ... process dropping vbp as request to open new form with same vbp.
            ' ... could go further much interface required.
            If LCase$(Right$(stext, 4)) = ".vbp" Then
                If Dir$(stext, vbNormal) <> "" Then
                    pLoadViewer stext
                End If
            End If
        End If
    ElseIf Data.GetFormat(vbCFFiles) Then
        sFile = Data.Files(1)
        sFile = LCase$(sFile)
        If Right$(sFile, 4) = ".vbg" Then
            Set xVBP = New VBPInfo
            Set xArray = xVBP.ReadVBG(sFile)
            For lngLoop = 1 To xArray.Count
                pLoadViewer xArray(lngLoop)
                If Not moFileHistory Is Nothing Then
                    moFileHistory.AddItemString xArray(lngLoop)
                End If
            Next lngLoop
        ElseIf Right$(sFile, 4) = ".vbp" Then
            pLoadViewer sFile
            If Not moFileHistory Is Nothing Then
                moFileHistory.AddItemString sFile
            End If
        ElseIf Right$(sFile, 4) = ".zip" Then
            If Len(mDropZip) = 0 Then
                ' ... mDropZip with a value means the unzip form is loaded
                ' ... and we don't want to process drops.
                mDropZip = sFile
                ' ... note: opening zip file from timer event
                ' ... because if open modal form here then
                ' ... ole is stuck until it closes so this is a cheap
                ' ... way to let ole drag complete first.
                tmrUnzip.Enabled = True
            End If
        End If
        If Not xVBP Is Nothing Then
            Set xVBP = Nothing
        End If
        If Not xArray Is Nothing Then
            Set xArray = Nothing
        End If
    End If

ResumeError:

Exit Sub

ErrHan:

    Debug.Print "mdiMain.MDIForm_OLEDragDrop.Error: " & Err.Number & "; " & Err.Description
    Err.Clear
    Resume Next
    
    Resume ResumeError:

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim iAnswer As VbMsgBoxResult
    
    If mOptions.ConfirmExit Then
        iAnswer = MsgBox("Are you sure you want to Close the Program:" & vbNewLine & App.Title & "?", vbQuestion + vbYesNo, "Close: " & App.Title)
        If iAnswer = vbNo Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim i As Long
Dim iAnswer As VbMsgBoxResult
Dim sUnZipFolder As String

    On Error Resume Next    ' ... some forms will already have been unloaded.
    ' -------------------------------------------------------------------
    Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
    ' -------------------------------------------------------------------
'    Unload frmTip
'    Unload frmRefViewer
    Unload frmPSCHistory
'    Unload frmSearch2
'    Unload frmProject2
    ' -------------------------------------------------------------------
'    If mChildCount > 0 Then
'        For i = 0 To mChildCount - 1
'            If Not moForms(i) Is Nothing Then
'                Unload moForms(i)
'            End If
'        Next i
'    End If
'    ' -------------------------------------------------------------------
'    Erase moForms
    ' -------------------------------------------------------------------
    If Not moFileHistory Is Nothing Then
        ' ... save drag'n'drop vbp history.
        If moFileHistory.Count > 0 Then
            moFileHistory.Remove "", True
            If moFileHistory.Count > 0 Then
                moFileHistory.ToFile App.Path & "\History.dat", vbCrLf
            End If
        End If
        Set moFileHistory = Nothing
    End If
    
    ' -------------------------------------------------------------------
    ' ... clean unzip folder?
    ' ... you is gonna laugh
    ' ... I managed to KillFolder on my project path and deleted
    ' ... the entire project's source :|
    ' ... good job I uploaded to PSC after all :D
    
    If mOptions.AutoCleanUnzipFolder Then
        sUnZipFolder = mOptions.UnzipFolder
        If Dir$(sUnZipFolder, vbDirectory) <> "" Then
            If UCase$(App.Path) <> UCase$(sUnZipFolder) Then
                ' ... v6, added the name of the folder to delete, see above.
                iAnswer = MsgBox("Are you sure you want to remove the unzip folder and its contents?" & vbNewLine & sUnZipFolder, vbQuestion + vbYesNo, "Auto-Clean Unzip Folder")
                If iAnswer = vbYes Then
                    KillFolder mOptions.UnzipFolder
                    If Dir$(mOptions.UnzipFolder, vbDirectory) <> "" Then
                        MsgBox "Not all files were deleted from the unzip folder.", vbInformation, "Auto-Clean Unzip Folder"
                    End If
                End If
            Else
                MsgBox "Unzip folder not removed because it is the same as the program's own folder.", vbInformation, "Auto-Clean Unzip Folder"
            End If
        End If
    End If
    Set mOptions = Nothing
    DoEvents
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    ' -------------------------------------------------------------------
End Sub

Private Sub mnuArrange_Click()
    Arrange VBRUN.vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    Arrange VBRUN.vbCascade
End Sub

Private Sub mnuCompResource_Click()
    frmCompileResource.Show vbModal
End Sub

Private Sub mnuCopyFile_Click()
'    frmCopyProjectFile.Show vbModal
End Sub

Private Sub mnuCopyProject_Click()
    frmCopyProject.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMgnHist_Click()
    frmMngVBPHistory.Show vbModal
    pLoadHistMenu
End Sub

Private Sub mnuFOpenZipFolder_Click()
Dim x As cOptions
Dim sUnZipFolder As String
Dim lngRet As Long

    Set x = New cOptions
    x.Read
    
    sUnZipFolder = x.UnzipFolder
    
    Set x = Nothing
    
    If Len(sUnZipFolder) Then
    
        lngRet = ShellExecute(0&, vbNullString, sUnZipFolder & "\", vbNullString, vbNullString, vbNormalFocus)
    
    End If
    
    If lngRet < 33 Then ' ... ShellExecute returns > 32 if successful, else it failed.
        ' ... confirm, did not open zip folder, no reason.
        MsgBox "The Program could not open the UnZip Folder:" & vbNewLine & sUnZipFolder & vbNewLine & vbNewLine & "ShellExecute Error Code: " & lngRet, vbInformation, "Zip Folder"
    End If
    
End Sub

Private Sub mnuHAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHistory_Click(Index As Integer)
    pLoadViewer Trim$(mnuHistory(Index).Caption)
'    frmProject2.LoadVBP Trim$(mnuHistory(Index).Caption)
'    frmProject2.Show
End Sub

Private Sub mnuHTMLHelpGen_Click()
'    frmHelpGen.Show ' vbModal
Dim x As DevHelpGen
Dim xVBP As VBPInfo
    
    Set xVBP = New VBPInfo
    xVBP.ReadVBP App.Path & "\codeviewer.vbp"
    
    Set x = New DevHelpGen
    x.Init xVBP
    

End Sub

Private Sub mnuNewWindow_Click()
    pLoadViewer
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
    If Not mOptions Is Nothing Then
        ' ... update options regardless.
        mOptions.Read
    End If
End Sub

Private Sub mnuProjSearch_Click()
'    frmSearchVBProject.Show
'    DoEvents
'    frmSearchVBP.ZOrder
End Sub

Private Sub mnuPSCDownloads_Click()
'    WinSeek.Show vbModal, Me
End Sub

Private Sub mnuRefViewer_Click()
    frmRefViewer.Show
    frmRefViewer.ZOrder
End Sub

Private Sub mnuShowHelp_Click()
'    modGeneral.ShowHelp
End Sub

Private Sub mnuTileHorizontally_Click()
    Arrange VBRUN.vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
    Arrange VBRUN.vbTileVertical
End Sub

Private Sub mnuTipODay_Click()
    frmTip.Show
End Sub

Private Sub mnuUnZip_Click()

Dim sFile As String
Dim sVBPFiles() As String
Dim lngCount As Long
Dim lngLoop As Long
Dim sFolder As String
Dim iAnswer As VbMsgBoxResult
Dim bOK As Boolean
Dim bAutoOpen As Boolean
Dim lngCountOfFiles As Long
Dim bShowUnzipForm As Boolean
    
    If mHaveReadZip = False Then
        bShowUnzipForm = True
    Else
        If frmUnZip.CountOfUnzippedFiles = 0 And frmUnZip.CountOfFilesInZip > 0 Then
            ' ... read but didn't unzip.
            bShowUnzipForm = True
        End If
    End If
    
    If bShowUnzipForm = True Then
        frmUnZip.Show vbModal
    End If
    
    If frmUnZip.Cancelled = True Then
        If frmUnZip.CountOfFilesInZip > 0 Then
            MsgBox "No Files were Unzipped, Unzip was cancelled.", vbInformation, "Unzip"
        End If
        GoTo Quit:
    End If
    
    lngCountOfFiles = frmUnZip.CountOfUnzippedFiles
    
    If Len(frmUnZip.VBPFileString) Then
        
        sFolder = mOptions.UnzipFolder & "\"
        bAutoOpen = mOptions.AutoLoadProjects
        
        modStringArrays.SplitString frmUnZip.VBPFileString, sVBPFiles, vbCrLf, lngCount
        If lngCount > 0 Then
            If bAutoOpen = False And lngCount > 1 Then
                iAnswer = MsgBox("Zip File contains " & Format$(lngCount, cNumFormat) & " VBP files:" & vbNewLine & "Would you like to open all of them regardless or confirm each one?" & vbNewLine & "Click Yes to Open all of them or No to confirm each one first.", vbQuestion + vbYesNo, "Unzip: Auto-Load VBP")
                bAutoOpen = iAnswer = vbYes
            End If
            For lngLoop = 0 To lngCount - 1
                sFile = sFolder & sVBPFiles(lngLoop)
                If Dir$(sFile, vbNormal) <> "" Then
                    If bAutoOpen = False Then
                        iAnswer = MsgBox("Would you like to load the following VBP?" & vbNewLine & sFile, vbQuestion + vbYesNo, "Unzip: Auto-Load VBP")
                        bOK = iAnswer = vbYes
                    End If
                    If bOK Or bAutoOpen Then
                        pLoadViewer sFile, "", mOptions.HideChildProject, mOptions.HideChildToolbar, frmUnZip.PSCText
                    End If
                End If
            Next lngLoop
        End If
        
    End If
    
'    If Len(frmUnZip.PSCText) Then
'        MsgBox frmUnZip.PSCText
'    End If
    
    If lngCountOfFiles > 0 Then
'        sFolder = frmUnZip.UnZipFolder
'        MsgBox "Unzip operation worked." & vbNewLine & "Files Unzipped: " & Format$(lngCountOfFiles, cNumFormat) & ", to the following folder ... " & vbNewLine & sFolder, vbInformation, "Unzip"
    Else
        MsgBox "Either there were no files to Unzip or the Unzip operation failed.", vbExclamation, "Unzip"
    End If
       
Quit:

End Sub

Private Sub mnuUnzipPSC_Click()
    If frmPSCHistory.WindowState = vbMinimized Then frmPSCHistory.WindowState = vbNormal
    frmPSCHistory.Show
    frmPSCHistory.ZOrder
End Sub

Private Sub tmrUnzip_Timer()

Dim bDoUnzip As Boolean
Dim bLoadForm As Boolean

    tmrUnzip.Enabled = False
    
    bDoUnzip = Len(mDropZip)
    
    If bDoUnzip Then
        bLoadForm = Not mOptions.AutoUnZip
        bDoUnzip = frmUnZip.ReadZip(mDropZip, bLoadForm)
        If bDoUnzip Then
            mHaveReadZip = True
            mnuUnZip_Click
        End If
    End If
    ' ... erase mDropZip after using frmUnZip
    ' ... ole drag'n'drop can still fire while it is open modally
    ' ... and events are cancelled if mDropZip is anything.
    mDropZip = vbNullString
    mHaveReadZip = False
    
End Sub
