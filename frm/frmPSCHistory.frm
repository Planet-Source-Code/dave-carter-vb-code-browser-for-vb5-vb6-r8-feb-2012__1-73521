VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPSCHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan for PSC Read Me Files"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   HelpContextID   =   16
   Icon            =   "frmPSCHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picVBP 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   1500
      ScaleHeight     =   3975
      ScaleWidth      =   8955
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   8955
      Begin ComctlLib.ListView lvVBP 
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Title"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "VBP Name"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Folder"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "DateNo."
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblPSCFolder 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PSC File Folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1440
         TabIndex        =   28
         ToolTipText     =   "Open PSC Read Me File Folder"
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblPSCFolderCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PSC File Folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblVBPFindCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "VBPs Found"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label lblVBPFound 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1320
         TabIndex        =   25
         ToolTipText     =   "No of Items Found"
         Top             =   3720
         Width           =   840
      End
      Begin VB.Image imgBtn 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   8520
         Picture         =   "frmPSCHistory.frx":058A
         Stretch         =   -1  'True
         ToolTipText     =   "Back to list of PSC files [Escape]"
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.PictureBox picFound 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   6360
      ScaleHeight     =   1755
      ScaleWidth      =   2115
      TabIndex        =   21
      Top             =   180
      Width           =   2175
      Begin VB.ListBox lstFoundFiles 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   2280
      ScaleHeight     =   1785
      ScaleWidth      =   3885
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   3915
      Begin VB.FileListBox filList 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   1860
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.DirListBox dirList 
         Appearance      =   0  'Flat
         Height          =   990
         Left            =   180
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   1575
      End
      Begin VB.DriveListBox drvList 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7455
      ScaleWidth      =   8895
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   8895
      Begin ComctlLib.ListView lv 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "double-click"
         Top             =   3960
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Link"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Folder"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "DateNo."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtFindName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   1995
      End
      Begin VB.TextBox txtFindDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         TabIndex        =   6
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmPSCHistory.frx":0B14
         Top             =   1080
         Width           =   8655
      End
      Begin VB.Label lblFilterCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtered Count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5280
         TabIndex        =   20
         Top             =   3540
         Width           =   1020
      End
      Begin VB.Label lblFilterCount 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6540
         TabIndex        =   19
         ToolTipText     =   "No of Items Found"
         Top             =   3540
         Width           =   840
      End
      Begin VB.Image imgBtn 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   8460
         Picture         =   "frmPSCHistory.frx":0B2B
         Stretch         =   -1  'True
         ToolTipText     =   "Filter by Name / Description"
         Top             =   3480
         Width           =   285
      End
      Begin VB.Label lblLink 
         BackStyle       =   0  'Transparent
         Caption         =   "Link to PSC Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Goto Submission on PSC"
         Top             =   780
         Width           =   8715
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Submission Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Left            =   120
         TabIndex        =   18
         Top             =   60
         UseMnemonic     =   0   'False
         Width           =   8595
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFileName 
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   7200
         Width           =   8715
      End
   End
   Begin VB.Label lblScanningFolder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3060
      TabIndex        =   17
      ToolTipText     =   "Folder to Scan"
      Top             =   8340
      Width           =   1140
   End
   Begin VB.Label lblFound 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No of Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   16
      ToolTipText     =   "No of Items Found"
      Top             =   8760
      Width           =   840
   End
   Begin VB.Label lblFoundCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Items Found"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   540
      TabIndex        =   15
      Top             =   8760
      Width           =   900
   End
   Begin VB.Image imgBtn 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   120
      Picture         =   "frmPSCHistory.frx":10B5
      Stretch         =   -1  'True
      ToolTipText     =   "Cancel Scan [Escape]"
      Top             =   8760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblScanStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Scanned"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   14
      ToolTipText     =   "Folder to Scan"
      Top             =   8340
      Width           =   915
   End
   Begin VB.Label lblScanStatusCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   540
      TabIndex        =   13
      ToolTipText     =   "Folder to Scan"
      Top             =   8340
      Width           =   855
   End
   Begin VB.Label lblFolderCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Folder to scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   540
      TabIndex        =   12
      ToolTipText     =   "Folder to Scan"
      Top             =   7920
      Width           =   1020
   End
   Begin VB.Label lblScanFolder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Folder to Scan"
      Top             =   7920
      Width           =   840
   End
   Begin VB.Image imgBtn 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   120
      Picture         =   "frmPSCHistory.frx":163F
      Stretch         =   -1  'True
      ToolTipText     =   "Scan"
      Top             =   8340
      Width           =   285
   End
   Begin VB.Image imgBtn 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      Picture         =   "frmPSCHistory.frx":1BC9
      Stretch         =   -1  'True
      ToolTipText     =   "Select Folder to Scan"
      Top             =   7860
      Width           =   285
   End
   Begin VB.Menu mnuVBP 
      Caption         =   "VBP"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Browse Project"
         Index           =   1
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Search Project"
         Index           =   2
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Containing Folder"
         Index           =   3
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "VB5"
         Index           =   4
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "VB6"
         Index           =   5
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Methods"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmPSCHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mtPSCInfo() As PSCInfo

Private miSearching As Integer
Private miCancelled As Integer

Private miScanCount As Long
Private mLMouseDown As Boolean
Private msScanFolder As String

Private moOptions As cOptions

Private Sub DrvList_Change()
    On Error GoTo ErrHan:
    dirList.Path = drvList.Drive
Exit Sub
ErrHan:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub DirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
End Sub

Private Sub pPreScan()
    ' -------------------------------------------------------------------
    Screen.MousePointer = vbHourglass
    ' -------------------------------------------------------------------
    filList.Pattern = "@PSC*.txt"
    ' -------------------------------------------------------------------
    lstFoundFiles.Clear
    lvVBP.ListItems.Clear
    picVBP.Visible = False
    ' -------------------------------------------------------------------
    lblScanStatus.Caption = "Scanning ..."
    lblFound.Caption = "0"
    miScanCount = 0
    ReDim mtPSCInfo(0)
    ' -------------------------------------------------------------------
    miCancelled = False
    miSearching = True
    ' -------------------------------------------------------------------
    pClearDisplay
    pShowHideScan True
End Sub

Private Sub pPostScan()
    ' -------------------------------------------------------------------
    miSearching = False
    ' -------------------------------------------------------------------
    lblScanStatus.Caption = "Scanned"
    lblScanningFolder.Caption = msScanFolder
    If miCancelled = True Then lblScanStatus.Caption = "Scan Cancelled"
    ' -------------------------------------------------------------------
    pLoadLV 'txtFindName.Text, txtFindDesc.Text
    pShowHideScan False
    ' -------------------------------------------------------------------
    lv.SetFocus
    ' -------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    ' -------------------------------------------------------------------
End Sub

Private Sub pClearDisplay()
    lv.ListItems.Clear
    txtDesc.Text = ""
    lblName.Caption = ""
    lblLink.Caption = ""
    lblFileName.Caption = ""
    lblFilterCount.Caption = "0"
End Sub

Private Sub pLoadLV(Optional pNameFilter As String = "", _
                    Optional pDescFilter As String = "")
                    
Dim iListIndex As Long
Dim xItem As ListItem
Dim bCheck As Boolean
Dim bAdd As Boolean
    ' -------------------------------------------------------------------
    pClearDisplay
    ' -------------------------------------------------------------------
    If miScanCount Then
        bCheck = Len(pNameFilter) > 0 Or Len(pDescFilter) > 0
        For iListIndex = 0 To UBound(mtPSCInfo)
            bAdd = True
            With mtPSCInfo(iListIndex)
                If bCheck Then
                    bAdd = False
                    If Len(pNameFilter) > 0 And Len(pDescFilter) > 0 Then
                        bAdd = InStrB(1, .Name, pNameFilter)
                        bAdd = bAdd And InStrB(1, .Description, pDescFilter)
                    ElseIf Len(pNameFilter) > 0 Then
                        bAdd = InStrB(1, .Name, pNameFilter)
                    Else
                        bAdd = InStrB(1, .Description, pDescFilter)
                    End If
                End If
                If bAdd Then
                    Set xItem = lv.ListItems.Add(, , .Name)
                    xItem.SubItems(1) = Format$(.FileDate, "dd MMM YYYY")
                    xItem.SubItems(2) = .Link
                    xItem.SubItems(3) = .Description
                    xItem.SubItems(4) = .UnzipFolder
                    xItem.SubItems(5) = .FileDate
                    ' -------------------------------------------------------------------
                    xItem.Tag = iListIndex
                End If
                ' -------------------------------------------------------------------
            End With
        Next iListIndex
        ' -------------------------------------------------------------------
        If lv.ListItems.Count Then
            Set xItem = lv.ListItems(1)
            xItem.Selected = True
            lv_ItemClick xItem
            LVSizeColumn lv, 3
            ' -------------------------------------------------------------------
            lblFilterCount.Caption = Format$(lv.ListItems.Count, "#,##0") & " of " & Format$(miScanCount, "#,##0")
            ' -------------------------------------------------------------------
        End If
        ' -------------------------------------------------------------------
    End If
End Sub

Private Function pbScan() As Boolean
Dim iRet As Integer
Dim iListIndex As Long
Dim sFile As String
Dim tFileInfo As FileNameInfo
Dim xString As SBuilder ' StringWorker
    ' -------------------------------------------------------------------
    pPreScan
    ' -------------------------------------------------------------------
    iRet = piFindFiles(msScanFolder, "", miScanCount)
    ' -------------------------------------------------------------------
    If miScanCount Then
        ReDim mtPSCInfo(miScanCount - 1)
        For iListIndex = 0 To miScanCount - 1
            sFile = lstFoundFiles.List(iListIndex)
            Set xString = New SBuilder ' StringWorker
            xString.ReadFromFile sFile
            ' get psc submission info -------------------------------------------
            ParsePSCInfo xString.TheString, mtPSCInfo(iListIndex)
            ' get file name info ------------------------------------------------
            ParseFileNameEx sFile, tFileInfo
            ' update psc info with extra file related info ----------------------
            With mtPSCInfo(iListIndex)
                .UnzipFolder = tFileInfo.Path
                .FileName = tFileInfo.PathAndName
                .FileDate = CLng(FileDateTime(tFileInfo.PathAndName))
            End With
            Set xString = Nothing
        Next iListIndex
    End If
    ' -------------------------------------------------------------------
    pPostScan
    ' -------------------------------------------------------------------
End Function

Private Sub pShowHideScan(pShow As Boolean)
    picDetail.Visible = Not pShow
    picFound.Visible = pShow
    lstFoundFiles.Visible = pShow
    imgBtn(0).Visible = Not pShow
    imgBtn(1).Visible = Not pShow
    imgBtn(2).Visible = pShow
    picFound.BorderStyle = 0 ' initially picFound has border for cosmetics, this removes it so only found lists border is showing
End Sub

Private Function pbCancelScan() As Boolean
    miSearching = False
    miCancelled = True
    pShowHideScan False
    pPostScan
End Function

Private Function piFindFiles(ByVal pNewFolder As String, _
                             ByVal pRestoreFolder As String, _
                             ByRef pNoOfFinds As Long) As Integer

Dim iRet As Integer
Dim iAbandon As Integer
Dim iFoldersToPeek As Long

Dim iFileCount As Long
Dim iFileIndex As Long

Dim sOldFolder As String
Dim sFolder As String
Dim sFileName As String

    ' -------------------------------------------------------------------
    miSearching = True
    piFindFiles = False
    lblScanningFolder.Caption = pNewFolder
    iRet = DoEvents()
    ' -------------------------------------------------------------------
    If miSearching = False Then
        piFindFiles = True
        Exit Function
    End If
    ' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    iFoldersToPeek = dirList.ListCount
    Do While iFoldersToPeek > 0 And miSearching = True
        sOldFolder = dirList.Path
        dirList.Path = pNewFolder
        If dirList.ListCount Then
            dirList.Path = dirList.List(iFoldersToPeek - 1)
            iAbandon = piFindFiles((dirList.Path), sOldFolder, pNoOfFinds)
        End If
        iFoldersToPeek = iFoldersToPeek - 1
        If iAbandon = True Then Exit Function
    Loop
    ' -------------------------------------------------------------------
    iFileCount = filList.ListCount
    If iFileCount Then
        sFolder = pNewFolder
        If Len(sFolder) > 3 Then sFolder = sFolder & "\"
        For iFileIndex = 0 To iFileCount - 1
            sFileName = sFolder & filList.List(iFileIndex)
            pNoOfFinds = pNoOfFinds + 1
            ' -------------------------------------------------------------------
            lstFoundFiles.AddItem sFileName
            lblFound.Caption = Format$(pNoOfFinds, "#,##0")
        Next iFileIndex
    End If
    ' -------------------------------------------------------------------
    If pRestoreFolder <> "" Then dirList.Path = pRestoreFolder
    ' -------------------------------------------------------------------
Exit Function
ErrHan:
    piFindFiles = True
    Debug.Print "frmPSCHistory.piFindFiles.Error: " & Err.Number & "; " & Err.Description
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If miSearching = True Then
            pbCancelScan
        Else
            If picVBP.Visible Then
                picVBP.Visible = False
            Else
            
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Width = (2 * picDetail.Left) + picDetail.Width + (3 * Screen.TwipsPerPixelX)
    ' -------------------------------------------------------------------
    Set moOptions = New cOptions
    moOptions.Read
    ' -------------------------------------------------------------------
    pLoadHandCursor
    msScanFolder = moOptions.UnzipFolder ' drvList.Drive & "\"
    If Dir$(msScanFolder, vbDirectory) = "" Then msScanFolder = App.Path
    dirList.Path = msScanFolder
    lblScanFolder.Caption = msScanFolder
    lblScanningFolder.Caption = ""
    ' -------------------------------------------------------------------
    LVFullRowSelect lv.hwnd
    LVFullRowSelect lvVBP.hwnd
    ' -------------------------------------------------------------------
    SetTextBoxCueBanner txtFindName.hwnd, "Filter on Name"
    SetTextBoxCueBanner txtFindDesc.hwnd, "Filter on Description"
    ' -------------------------------------------------------------------
    With picDetail
        picFound.Move .Left, .Top, .Width, .Height
        With picFound
            lstFoundFiles.Move 0, 0, .Width, .Height
        End With
        picVBP.Left = .Left
    End With
    ' -------------------------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If miSearching = True Then
        Cancel = True
    Else
        pClearDisplay
        Set moOptions = Nothing
    End If
End Sub

Private Sub imgBtn_Click(Index As Integer)
Dim sFolder As String
Dim bOK As Boolean
    Select Case Index
        Case 0 ' select folder to scan
            sFolder = GetFolder("Select Folder to Scan", msScanFolder)
            If Len(sFolder) > 0 Then
                msScanFolder = sFolder
                lblScanFolder.Caption = msScanFolder
                dirList.Path = msScanFolder
            End If
        Case 1 ' scan
            bOK = pbScan
        Case 2 ' cancel scan
            bOK = pbCancelScan
        Case 3 ' filter results
            picVBP.Visible = False
            pLoadLV txtFindName.Text, txtFindDesc.Text
        Case 4 ' hide vbp list
            picVBP.Visible = False
    End Select
End Sub

Private Sub lblFileName_Click()
Dim sTmp As String
    sTmp = Trim$(lblFileName.Caption)
    If Len(sTmp) > 0 Then
        RunProgram sTmp
    End If
End Sub

Private Sub lblLink_Click()
Dim sTmp As String
    sTmp = Trim$(lblLink.Caption)
    If Len(sTmp) > 0 Then
        OpenWebPage sTmp
    End If
End Sub

Private Sub lblPSCFolder_Click()
Dim sTmp As String
    sTmp = Trim$(lblPSCFolder.Caption)
    If Len(sTmp) > 0 Then
        OpenFolder sTmp
    End If
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Dim iColIndex As Long
    iColIndex = ColumnHeader.Index
    If iColIndex = 2 Then iColIndex = 6
'    Debug.Print ColumnHeader.Index, ColumnHeader.Text
    LVSortTextCol lv, iColIndex - 1 ' ColumnHeader.Index - 1
End Sub

Private Sub lv_DblClick()
Dim tFileInfo As FileNameInfo
Dim sOldPath As String
Dim i As Long
Dim iFound As Long
Dim iRet As Integer
Dim xVBP As VBPInfo
Dim iIndex As Long
Dim xItem As ListItem
Dim sTmp As String
Dim iDate As Long

    On Error GoTo ErrHan:
    If lv.ListItems.Count = 0 Then Exit Sub
    If lv.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    ' -------------------------------------------------------------------
    lvVBP.ListItems.Clear
    lstFoundFiles.Clear
    ' -------------------------------------------------------------------
    filList.Pattern = "*.vbp"
    sOldPath = dirList.Path
    sTmp = lblScanStatus.Caption
    lblScanStatus.Caption = "Scan 4 VBPs ..."
    lblFound.Caption = "0"
    ' -------------------------------------------------------------------
    i = lv.ListItems(lv.SelectedItem.Index).Tag
    ParseFileNameEx mtPSCInfo(i).FileName, tFileInfo
    lblPSCFolder.Caption = tFileInfo.Path
    dirList.Path = tFileInfo.Path
    iRet = piFindFiles((dirList.Path), sOldPath, iFound)
    ' -------------------------------------------------------------------
    If iFound Then
        lblScanStatus.Caption = "Reading VBPs ..."
        For iIndex = 0 To lstFoundFiles.ListCount - 1
            lblScanningFolder.Caption = lstFoundFiles.List(iIndex)
            lblScanFolder.Refresh
            Set xVBP = New VBPInfo
            xVBP.ReadVBP lstFoundFiles.List(iIndex)
            If xVBP.Initialised Then
                With xVBP
                iDate = CLng(FileDateTime(.FileNameAndPath))
                    Set xItem = lvVBP.ListItems.Add(, , Trim$(.Title))
                    xItem.SubItems(1) = .FileName
                    xItem.SubItems(2) = Format$(iDate, "dd MMM YYYY")
                    xItem.SubItems(3) = .Description
                    xItem.SubItems(4) = .FilePath
                    xItem.SubItems(5) = iDate
                    xItem.Tag = .FileNameAndPath
                End With
            End If
            Set xVBP = Nothing
        Next iIndex
        ' -------------------------------------------------------------------
        LVSizeColumn lvVBP, 3
        picVBP.Visible = True
        picVBP.ZOrder
        ' -------------------------------------------------------------------
    End If
ResumeError:
    On Error GoTo 0
    Set xVBP = Nothing
    miSearching = False
    ' -------------------------------------------------------------------
    lblScanStatus.Caption = sTmp
    lblFound.Caption = Format$(miScanCount, "#,##0")
    lblScanningFolder.Caption = msScanFolder
    lblVBPFound.Caption = Format$(iFound, "#,##0")
    ' -------------------------------------------------------------------
    Screen.MousePointer = vbDefault
Exit Sub
ErrHan:
    Debug.Print "frmPSCHistory.lv_DblClick.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim i As Long
    If Item Is Nothing Then Exit Sub
    i = Item.Tag
    On Error GoTo ErrHan:
    With mtPSCInfo(i)
        lblName.Caption = .Name
        lblFileName.Caption = .FileName
        lblLink.Caption = .Link
        ' ... replace char 10 with new line as intended.
        txtDesc.Text = modStrings.Replace(.Description, Chr$(10), vbCrLf)
    End With
    picVBP.Visible = False
    
Exit Sub
ErrHan:
    Debug.Print "frmPSCHistory.lv_ItemClick.Error: " & Err.Number & "; " & Err.Description
End Sub

Private Sub imgBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Image: Shift Right and Down.
    If mLMouseDown Then Exit Sub
    imgBtn(Index).Move imgBtn(Index).Left + 15, imgBtn(Index).Top + 15
    mLMouseDown = True
End Sub

Private Sub imgBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Image: Shift Left and Up.
    If mLMouseDown = False Then Exit Sub
    imgBtn(Index).Move imgBtn(Index).Left - 15, imgBtn(Index).Top - 15
    mLMouseDown = False
End Sub

Private Sub pLoadHandCursor()
' ... try and load the hand cursor for image buttons.
Dim lngCount As Long
Dim lngLoop As Long
Dim iHandle As Long
Dim iCursor As StdPicture
    iHandle = modHandCursor.LoadHandCursor
    If iHandle <> 0 Then
        Set iCursor = modHandCursor.HandleToPicture(iHandle, False)
        If iHandle = 0 Then Exit Sub
        ' -------------------------------------------------------------------
        On Error Resume Next
        ' -------------------------------------------------------------------
        lngCount = imgBtn.Count
        For lngLoop = 0 To lngCount - 1
            imgBtn(lngLoop).MouseIcon = iCursor
            imgBtn(lngLoop).MousePointer = vbCustom
        Next lngLoop
        ' -------------------------------------------------------------------
        lblLink.MouseIcon = iCursor
        lblLink.MousePointer = vbCustom
        lblFileName.MouseIcon = iCursor
        lblFileName.MousePointer = vbCustom
        lblPSCFolder.MouseIcon = iCursor
        lblPSCFolder.MousePointer = vbCustom
    End If
End Sub

Private Sub lvVBP_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Dim iColIndex As Long
    iColIndex = ColumnHeader.Index
    If iColIndex = 3 Then iColIndex = 6
    LVSortTextCol lvVBP, iColIndex - 1
End Sub

Private Sub lvVBP_DblClick()
    mnuOpen_Click 1 ' open in viewer window
End Sub

Private Sub lvVBP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xItem As ListItem
    If Button = 2 Then
        If lvVBP.ListItems.Count = 0 Then Exit Sub
        If lvVBP.SelectedItem Is Nothing Then Exit Sub
        ' -------------------------------------------------------------------
        mnuOpen(4).Enabled = moOptions.ShowVB5IDE
        mnuOpen(5).Enabled = moOptions.ShowVB6IDE
        ' -------------------------------------------------------------------
        PopupMenu mnuVBP
        ' -------------------------------------------------------------------
    End If
End Sub

Private Sub mnuOpen_Click(Index As Integer)
Dim tFileInfo As FileNameInfo
Dim sApp As String
Dim lngRet As Long
Dim xVBP As VBPInfo
Dim xRForm As IReportForm

    If lvVBP.ListItems.Count = 0 Then Exit Sub
    If lvVBP.SelectedItem Is Nothing Then Exit Sub
    ParseFileNameEx lvVBP.SelectedItem.Tag, tFileInfo
    Select Case Index
        Case 1 ' browse project
            mdiMain.LoadFile tFileInfo.PathAndName
        Case 2 ' search project
            frmSearchVBProject.LoadVBP tFileInfo.PathAndName
        Case 3 ' containing folder
            OpenFolder tFileInfo.Path
        Case 4, 5 ' vb5, vb6
            If Index = 4 Then
                sApp = moOptions.PathToVB5
            ElseIf Index = 5 Then
                sApp = moOptions.PathToVB6
            End If
            If Len(sApp) > 0 Then
                If Dir$(sApp, vbNormal) <> "" Then
                    lngRet = Shell(sApp & " " & Chr$(34) & tFileInfo.PathAndName & Chr$(34), vbNormalFocus)
                End If
            End If
        Case 6 ' members
            Set xVBP = New VBPInfo
            xVBP.ReadVBP tFileInfo.PathAndName
            If xVBP.Initialised Then
                Set xRForm = frmMembers
                xRForm.Init xVBP
                xRForm.ZOrder
                Set xRForm = Nothing
                Set xVBP = Nothing
            End If
    End Select
        
End Sub

