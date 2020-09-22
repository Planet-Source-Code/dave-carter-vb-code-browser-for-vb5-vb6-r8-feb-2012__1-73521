VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmViewer 
   Caption         =   "Code Viewer"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   11880
   Begin CodeViewer.ucMenus PMenu 
      Left            =   480
      Top             =   3360
      _ExtentX        =   1508
      _ExtentY        =   1931
   End
   Begin VB.Timer tmrPNode 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1740
      Top             =   3780
   End
   Begin VB.PictureBox picTB 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   11880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11880
      Begin VB.CheckBox chkBoldRTF 
         Caption         =   "Bold"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3750
         TabIndex        =   17
         Top             =   180
         Width           =   615
      End
      Begin VB.CheckBox chkLineNos 
         Alignment       =   1  'Right Justify
         Caption         =   "Line Nos."
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4680
         TabIndex        =   18
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox chkAlign 
         Caption         =   "Align Right"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4140
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkColour 
         Caption         =   "In Colour"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4140
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2100
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkWordWrap 
         Caption         =   "Word Wrap"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5460
         TabIndex        =   19
         Top             =   630
         Width           =   1155
      End
      Begin VB.ComboBox cboFont 
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   0
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   585
      End
      Begin VB.ComboBox cboFont 
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   1
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkFHighlight 
         Caption         =   "Highlight Found"
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   4140
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1740
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkFMatchCase 
         Caption         =   "Match Case"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   8205
         TabIndex        =   22
         Top             =   630
         Width           =   1245
      End
      Begin VB.CheckBox chkFWholeWord 
         Caption         =   "Whole Word"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   6840
         TabIndex        =   21
         Top             =   630
         Width           =   1245
      End
      Begin VB.TextBox txtFind 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   6840
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   135
         Width           =   2055
      End
      Begin VB.Image imgBtn 
         Height          =   225
         Index           =   14
         Left            =   4380
         Picture         =   "frmViewer.frx":058A
         Stretch         =   -1  'True
         ToolTipText     =   "PSC Info"
         Top             =   195
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   9
         Left            =   9000
         Picture         =   "frmViewer.frx":0B14
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   240
         Index           =   13
         Left            =   6240
         Picture         =   "frmViewer.frx":109E
         Top             =   195
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBtn 
         Height          =   240
         Index           =   12
         Left            =   5940
         Picture         =   "frmViewer.frx":1628
         Top             =   195
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   11
         Left            =   2580
         Picture         =   "frmViewer.frx":1BB2
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Label lblLineCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Count: "
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3780
         TabIndex        =   16
         Top             =   660
         Width           =   870
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   10
         Left            =   4980
         Picture         =   "frmViewer.frx":213C
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   8
         Left            =   3330
         Picture         =   "frmViewer.frx":26C6
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   7
         Left            =   2970
         Picture         =   "frmViewer.frx":2C50
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   6
         Left            =   2220
         Picture         =   "frmViewer.frx":31DA
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   5
         Left            =   1860
         Picture         =   "frmViewer.frx":3764
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   4
         Left            =   1440
         Picture         =   "frmViewer.frx":3CEE
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   3
         Left            =   1110
         Picture         =   "frmViewer.frx":4278
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   2
         Left            =   780
         Picture         =   "frmViewer.frx":4802
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   1
         Left            =   450
         Picture         =   "frmViewer.frx":4D8C
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   60
         Picture         =   "frmViewer.frx":5316
         Stretch         =   -1  'True
         Top             =   135
         Width           =   285
      End
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   660
         Width           =   390
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   9570
         TabIndex        =   23
         Top             =   210
         Width           =   180
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   9570
         TabIndex        =   24
         Top             =   660
         Width           =   180
      End
   End
   Begin ComctlLib.TreeView tvProj 
      Height          =   3495
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   6165
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   18
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picSplitProj 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   1980
      ScaleHeight     =   1065
      ScaleWidth      =   105
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   11820
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5235
      Width           =   11880
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   2460
      ScaleHeight     =   3510
      ScaleWidth      =   5145
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1350
      Width           =   5145
      Begin VB.PictureBox picSplitMain 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   2310
         ScaleHeight     =   1455
         ScaleWidth      =   75
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   75
      End
      Begin ComctlLib.TreeView tvMembers 
         Height          =   2295
         HelpContextID   =   29
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   4048
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   18
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
         OLEDragMode     =   1
      End
      Begin VB.PictureBox picCodeCanvas 
         Height          =   3195
         Left            =   2490
         ScaleHeight     =   3135
         ScaleWidth      =   2385
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   2445
         Begin VB.PictureBox picSplitRTB 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Height          =   90
            Left            =   810
            ScaleHeight     =   90
            ScaleWidth      =   705
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2820
            Visible         =   0   'False
            Width           =   705
         End
         Begin RichTextLib.RichTextBox rtb 
            Height          =   525
            Index           =   0
            Left            =   150
            TabIndex        =   12
            Top             =   2340
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   926
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmViewer.frx":58A0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox rtb 
            Height          =   525
            Index           =   1
            Left            =   1560
            TabIndex        =   10
            Top             =   2310
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   926
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmViewer.frx":5920
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFFFFF&
            Height          =   795
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   1995
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   2055
            Begin VB.Label lblDesc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Left            =   90
               TabIndex        =   9
               Tag             =   "..."
               Top             =   450
               UseMnemonic     =   0   'False
               Width           =   180
            End
            Begin VB.Label lblClassName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   285
               Left            =   90
               TabIndex        =   8
               Tag             =   "..."
               Top             =   90
               UseMnemonic     =   0   'False
               Width           =   225
            End
         End
         Begin VB.Image imgSplitRTB 
            Height          =   165
            Left            =   840
            MousePointer    =   7  'Size N S
            Top             =   2490
            Width           =   645
         End
      End
      Begin VB.Image imgSplitMain 
         Height          =   1425
         Left            =   2040
         MousePointer    =   9  'Size W E
         Top             =   180
         Width           =   105
      End
   End
   Begin MSComDlg.CommonDialog cdgHTM 
      Left            =   1710
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "htm"
      Filter          =   "HTML (*.htm) | *.htm"
   End
   Begin MSComDlg.CommonDialog cdgSave 
      Left            =   1710
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "rtf"
      DialogTitle     =   "Save As ..."
      Filter          =   "Rich Text File (*.rtf) | *.rtf"
   End
   Begin VB.Image imgSplitProj 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   1740
      MousePointer    =   9  'Size W E
      Top             =   1350
      Width           =   165
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Code Viewer form."


Option Explicit
' -------------------------------------------------------------------
' v8
Private mMemberDrag As Boolean
' -------------------------------------------------------------------
' v6
Private mWithSyntaxColours As Boolean
Private mCodeLoadAll As Boolean
' -------------------------------------------------------------------

Private mUseChildWindows As Boolean
Private mShowClassHeadCount As Boolean
Private mShowAttributes As Boolean
Private mPSCReadMeText As String

'Private Const EM_LINESCROLL = &HB6
'Private Const EM_GETFIRSTVISIBLELINE = &HCE

Event UnLoaded(Index As Long)
Public ChildIndex As Long

Private mHistIndex As Long
Private mMovingBack As Boolean

Private mBoldRTF As Boolean
Private mHistArray As StringArray

Private m_CurrentTreeText As String
Private m_CurrentTreeFileName As String
Private m_CurrentTreeGUID As String
'Private m_CurrentProjectText As String
Private mLoadingClass As Boolean

' ... the following fields and types aid the resizing source only.
Private mLMouseDown As Boolean
Private mRMouseDown As Boolean
Private mSizeEditorOnly As Boolean

Private Const cHBorder As Long = 60
Private Const cSplitLimit As Long = 660
Private Const cSplitterHeight As Long = 30
Private Const cSplitterWidth As Long = 60


'Private sFilePath As String
Private sFileName As String
'Private sFileExt As String
Private m_loading As Boolean

Private mShift As Boolean
Private mControl As Boolean


'Private Const cDLGCancelErr As Long = 32755
'Private Const WM_USER = &H400
'Private Const EM_SETTARGETDEVICE = (WM_USER + 72)

Private Const ECM_FIRST = &H1500                    ' ... Edit control messages.

Private Const EM_SETCUEBANNER = (ECM_FIRST + 1)

'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private moVBPInfo As VBPInfo
Private moCodeReader As CodeInfo  ' CodeReader

Private mF2LastNode As String
Private mProjectName As String

' v3
Private mLineNos As Boolean
Private mColoured As Boolean

' v4
Private mRTFColourTable As String
Private moLastProjNode As Node
Private mCodeFileName As String
Private mProjectDesc As String

' v5
' ... variables for the Hand Cursor.
Private myHandCursor As StdPicture
Private myHand_handle As Long
Private mLastCopyFolder As String
Private mLastSaveFolder As String
Private mLastOpenFolder As String

' -------------------------------------------------------------------
' v5 ... hand cursor for image buttons

Private Sub pLoadHandCursor()

Dim lngCount As Long
Dim lngLoop As Long

' ... try and load the hand cursor.
    myHand_handle = modHandCursor.LoadHandCursor
    
    If myHand_handle <> 0 Then
        
        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
        
        If myHand_handle = 0 Then Exit Sub
        
        On Error Resume Next
        
        lngCount = imgBtn.Count
        
        For lngLoop = 0 To lngCount - 1
            imgBtn(lngLoop).MouseIcon = myHandCursor
            imgBtn(lngLoop).MousePointer = vbCustom
        Next lngLoop
        
    End If
    
End Sub

' -------------------------------------------------------------------

Private Sub pClearHistory()
' ... clear visited history.
    If Not mHistArray Is Nothing Then
        mHistArray.Clear
        Set mHistArray = New StringArray
        mHistArray.DuplicatesAllowed = False
    End If
    mHistIndex = 0&
    imgBtn(12).Visible = False
    imgBtn(13).Visible = False
    
End Sub

Private Sub pAddHistory(pNode As Node)
' ... add to visit history.
Dim lngFirstLine As Long
Dim lngCharPos As Long
Dim lngSelLen As Long
Dim bAdd As Boolean
Dim sBase As String
Dim sExt As String
Dim lngLoop As Long
Dim lngCount As Long
Dim lngLen As Long
Dim sTmp As String
' note: was using no duplicates on string array until i added the
'       first visible line and char pos stuff making duplicates possible.

    On Error GoTo ErrHan:
    
    If mMovingBack = True Then Exit Sub
    
    If mHistArray Is Nothing Then
        Set mHistArray = New StringArray
        mHistArray.DuplicatesAllowed = False
    End If
    
    ' ... first build the unique bit of the history item
    sBase = CStr(pNode.Index) & Chr$(0) & pNode.Text & Chr$(0)
    lngLen = Len(sBase)
    
    ' ... then build the dynamic stuff for the item.
'    lngFirstLine = SendMessageLong(rtb(1).hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    lngFirstLine = modGeneral.GetFirstVisibleLineRTFBox(rtb(1).hwnd) ' v6, added/swapped with above.
    If lngFirstLine = -1 Then lngFirstLine = 0
    
    lngCharPos = rtb(1).SelStart
    lngSelLen = rtb(1).SelLength
    sExt = CStr(lngFirstLine) & Chr$(0) & CStr(lngCharPos) & Chr$(0) & CStr(lngSelLen)
    
    lngCount = mHistArray.Count
    bAdd = True
    
    If lngCount > 0 Then
        For lngLoop = 1 To lngCount
            ' ... check if item exists, if so, update it and exit for.
            sTmp = Left$(mHistArray(lngLoop), lngLen)
            If sTmp = sBase Then
                ' ... item exists, update it.
                mHistArray.Item(lngLoop) = sBase & sExt
                ' ... don't add below.
                bAdd = False
                Exit For
            End If
        Next lngLoop
    End If
        
    If bAdd = True Then
        bAdd = mHistArray.AddItemString(sBase & sExt)
    End If
    
    mHistIndex = mHistArray.Count
    
    imgBtn(12).Visible = mHistIndex > 1
'    imgBtn(13).Visible = False

Exit Sub
ErrHan:
    Debug.Print "frmViewer.pAddHistory.Error: " & Err.Description
    
End Sub

Friend Property Get VBPInfo() As QuickVBPInfo
' ... Read-Only Interface to quick vbp information.
Dim x As QuickVBPInfo
    
    If Not moVBPInfo Is Nothing Then
        x = moVBPInfo.QuickInfo
    Else
        x.Name = "Unknown Source"
        x.Description = "No VBP Information Available."
    End If
    ' -------------------------------------------------------------------
    VBPInfo = x
    ' -------------------------------------------------------------------
End Property

Public Sub LoadVBP(pTheFile As String, _
          Optional pKey As String = vbNullString, _
          Optional pHideProject As Boolean = False, _
          Optional pHideToolbar As Boolean = False, _
          Optional ByVal pPSCReadMeText As String = vbNullString)
          
' ... Cheap Interface instruction to load a new instance with a given project.
Dim sTheFile As String
    sTheFile = pTheFile
    sTheFile = LCase$(sTheFile)
    If Dir$(sTheFile, vbNormal) <> "" Then
        If Right$(sTheFile, 4) = ".vbp" Then
            pOpenVBP sTheFile, pKey, pHideProject, pHideToolbar, pPSCReadMeText
        End If
    End If
End Sub

Public Property Get CodeFileName() As String
    CodeFileName = mCodeFileName
End Property

Public Property Get ProjectName() As String
' ... Cheap Interface to the name of the current loaded project, if any.
    ProjectName = mProjectName
End Property

Private Sub cboFont_Click(Index As Integer)

' ... updated v7
' ... Respond to a Font Combo selection, force current Member Node re-CLick.

' ... Thanks to Zhu JinYong for notes regarding rich text box selection being lost on change of font & size.
' ... I was using HideSelection on the RT Box to avoid flicker when scrolling sizes / names
' ... it turns out that hiding the control during update is pretty ok for this.
' ... And I had previously ignored selected text.

' ... the following method defers to the last tv member node click
' ... as the most simple (though not fastest) and safe thing to do.
' ... i looked into just changing rtf header properties in the rtf text
' ... and reloading it but the start colour is defined inside the header
' ... rather than at the end and the header returned is altered in other ways as well.

Dim lngFSize As Long
Dim xNode As Node
Dim i As Integer

Dim lngFirstLine As Long
Dim lngSelStart As Long
Dim lngSelLength As Long

    On Error GoTo ErrHan:
    
    If m_loading Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If Not tvMembers.SelectedItem Is Nothing Then
                        
        rtb(1).Visible = False
        
        lngFirstLine = modGeneral.GetFirstVisibleLineRTFBox(rtb(1).hwnd)
        lngSelStart = rtb(1).SelStart
        lngSelLength = rtb(1).SelLength
        
        Set xNode = tvMembers.Nodes(tvMembers.SelectedItem.Index)
        
        If Not xNode Is Nothing Then
        
            tvMembers_NodeClick xNode
            
            modGeneral.ScrollRTFBox rtb(1).hwnd, lngFirstLine
                        
            rtb(1).SelStart = lngSelStart
            rtb(1).SelLength = lngSelLength
            
            rtb(1).Visible = True
            
            Set xNode = Nothing
            
        End If
        
    End If

ResumeError:

    lngFirstLine = 0&
    lngSelStart = 0&
    lngSelLength = 0&
    lngFSize = 0&
    
    Screen.MousePointer = vbDefault
    
Exit Sub

ErrHan:

    Debug.Print "frmViewer.cboFont_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
        
End Sub


Private Sub chkAlign_Click()
' ... Align code viewer text Right or Left.
' ... can do this with rtf but takes longer to process...
' ... e.g. can process rtf normally with instruction to align right in header.

'Dim iMPointer As MousePointerConstants
'
'    On Error Resume Next    ' ... ensure can restore mouse pointer.
'    ' -------------------------------------------------------------------
'    iMPointer = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'    ' -------------------------------------------------------------------
'    rtb(1).SelStart = 0
'    rtb(1).SelLength = Len(rtb(1).TextRTF)
'    If chkAlign.Value = VBRUN.vbChecked Then
'        If rtb(1).SelAlignment <> rtfRight Then
'            If chkWordWrap.Value <> VBRUN.vbChecked Then
'                ' ... force word wrap on else won't right align.
'                chkWordWrap.Value = VBRUN.vbChecked
'            End If
'            rtb(1).SelAlignment = rtfRight
'        End If
'    Else
'        If rtb(1).SelAlignment <> rtfLeft Then
'            rtb(1).SelAlignment = rtfLeft
'        End If
'    End If
'    rtb(1).SelLength = 0: rtb(1).SelStart = 0
'    ' -------------------------------------------------------------------
'    Screen.MousePointer = iMPointer
'    ' -------------------------------------------------------------------
End Sub

Private Sub chkBoldRTF_Click()
    mBoldRTF = chkBoldRTF.Value = vbChecked
    lblClassName.FontBold = mBoldRTF
    lblDesc.FontBold = mBoldRTF
End Sub

Private Sub chkColour_Click()
' ... Turn coloured code On and Off.
' ... Hidden from use.
Dim iMPointer As MousePointerConstants
    
    On Error Resume Next    ' ... ensure can restore mouse pointer.
    ' -------------------------------------------------------------------
'    mColoured = Not mColoured
    iMPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    ' -------------------------------------------------------------------
    If chkColour.Value = VBRUN.vbChecked Then
        pColouriseText
    Else
        rtb(1).SelStart = 0
        rtb(1).SelLength = Len(rtb(1).TextRTF)
        rtb(1).SelColor = &H0
        rtb(1).SelLength = 0
    End If
    ' -------------------------------------------------------------------
    Screen.MousePointer = iMPointer
    ' -------------------------------------------------------------------
End Sub

Private Sub chkLineNos_Click()
' ... toggle line numbers flag.
    mLineNos = Not mLineNos
End Sub

Private Sub chkWordWrap_Click()
' ... Set Word Wrap On / Off; took ages to find this code, thanks the unknown coder.
Dim bValue As Boolean
    bValue = chkWordWrap.Value And vbChecked
    WordWrapRTFBox rtb(0).hwnd, bValue
    WordWrapRTFBox rtb(1).hwnd, bValue
'    SendMessageLong rtb(0).hwnd, EM_SETTARGETDEVICE, 0, IIf(chkWordWrap.Value = VBRUN.vbChecked, 0, 1)
'    SendMessageLong rtb(1).hwnd, EM_SETTARGETDEVICE, 0, IIf(chkWordWrap.Value = VBRUN.vbChecked, 0, 1)
End Sub

Private Sub cmdFind_Click()

' ... Helper:   Find Text in RichText Box (uses Statics).
' ... Note:     Only looks in rtb(1).
' ...           Uses the Rich Text Box find method.

Dim sFind As String
Dim lngFound As Long
Dim lngOptions As Long
Dim lngStart As Long
Dim lngLine As Long
' ... Statics, Yuck!
Static stLastPos As Long
Static stLastText As String
Static stPosIndex As Long

Dim saResults As StringArray
Dim lngResCount As Long
Dim lngResLoop As Long
    
    On Error GoTo ErrHan:
    
    If Len(txtFind.Text) Then
                
        sFind = txtFind.Text
        lngStart = 1
        If sFind = stLastText Then
            If stLastPos > -1 Then
                lngStart = stLastPos
            End If
        Else
            ' -------------------------------------------------------------------
            ' ... v8, try to find total count.
            If Not moCodeReader Is Nothing Then
                Set saResults = moCodeReader.GetAllMatches(sFind, , , , lngResCount)
                lblFind(0).Caption = lngResCount & " matches, " & saResults.Count & " unique line instances"
            End If
            
        End If
        
        If chkFWholeWord.Value = VBRUN.vbChecked Then
            lngOptions = lngOptions + rtfWholeWord
        End If
        
        If chkFMatchCase.Value = VBRUN.vbChecked Then
            lngOptions = lngOptions + rtfMatchCase
        End If
        
        If chkFHighlight.Value = VBRUN.vbUnchecked Then
            lngOptions = lngOptions + rtfNoHighlight
        End If
        ' -------------------------------------------------------------------
        lngFound = rtb(1).Find(sFind, lngStart, , lngOptions)
        ' -------------------------------------------------------------------
        If lngFound > 0 Then
            stPosIndex = stPosIndex + 1
            lngLine = rtb(1).GetLineFromChar(lngFound) + 1 ' ... else starts with line zero.
            stLastText = sFind
            stLastPos = lngFound + 1 ' ... shift start beyond this find for next.
            lblFind(1).Caption = "Item: " & Format$(stPosIndex, cNumFormat) & ".  Line: " & Format$(lngLine, cNumFormat) & ".  Char. Pos: " & Format$(lngFound, cNumFormat)
        Else
            stPosIndex = 0
            stLastText = vbNullString
            stLastPos = -1
            lblFind(1).Caption = "Not Found"
            rtb(1).SelLength = 0
            rtb(1).SelStart = 0
        End If
                
        If lngFound > 0 Then
            On Error GoTo 0
'            rtb(1).SetFocus
        End If
                
    End If

Exit Sub
ErrHan:
    lblFind(1).Caption = "Error on Find: " & Err.Description
    lblFind(0).Caption = vbNullString
    Debug.Print "frmViewer.cmdFind_Click.Error: " & Err.Number & "; " & Err.Description
End Sub

Private Sub pPrint()
' ... Helper: Show Print Form.
    If Len(rtb(1).TextRTF) Then
        If Len(rtb(1).Text) Then
            frmPrint.RTFText = rtb(1).TextRTF
        End If
    End If
    ' -------------------------------------------------------------------
    frmPrint.Show vbModal
    ' -------------------------------------------------------------------
End Sub

Private Sub cmdSaveRTF_Click()
' ... helper: Save Current RTF Text.
' ... note:   selection ignored / whole text or nothing.
Dim sRTFFileName As String
Dim sMsg As String
Dim sTitle As String
Dim lngMsgIcon As Long
Dim xFileInfo As FileNameInfo

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    sMsg = "Haven't anything to Save yet."
    ' -------------------------------------------------------------------
    If Len(rtb(1).Text) Then
        If Len(rtb(1).Text) Then
            
            If Len(mLastSaveFolder) Then
                cdgSave.InitDir = mLastSaveFolder
            End If
            
            cdgSave.ShowSave
            
            modFileName.ParseFileNameEx cdgSave.FileName, xFileInfo
            sRTFFileName = xFileInfo.PathAndName ' cdgSave.Filename
            mLastSaveFolder = xFileInfo.Path
            
            rtb(1).SaveFile sRTFFileName
            sMsg = vbNullString
        Else    ' ... what dude?
        End If
    End If
    ' -------------------------------------------------------------------
ResumeError:
    ' -------------------------------------------------------------------
    If Len(sMsg) Then
        sMsg = "Nothing was saved:" & vbNewLine & sMsg
        lngMsgIcon = vbExclamation
        sTitle = "Not Saved"
    Else
        sMsg = "The File :" & vbNewLine & sRTFFileName & vbNewLine & "was Saved"
        sTitle = "Saved"
        lngMsgIcon = vbInformation
    End If
    ' -------------------------------------------------------------------
    lngMsgIcon = lngMsgIcon + vbOKOnly
    ' -------------------------------------------------------------------
    MsgBox sMsg, lngMsgIcon, sTitle
    ' -------------------------------------------------------------------
Exit Sub
ErrHan:
    If Err.Number = cDlgCancelErr Then
        sMsg = "Save Dialog was Cancelled"
    Else
        Debug.Print "frmMain.cmdSaveRTF_Click.Error: " & Err.Number & "; " & Err.Description
        sMsg = Err.Description
    End If
    Resume ResumeError:
End Sub

Private Sub pShowHelp()
'    modGeneral.ShowHelp
End Sub

Private Sub Form_Activate()
    pResize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'   Cheap way to tell if Control or Shift keys are pressed when doing some process.
'   Tried using GetASyncKeyState API but not successfully; it needed to catch up as it were.
'   Not entirely satisfactory solution; if app lostfocus goes before keyup then keyup not processed.
    mControl = Shift And 2
    mShift = Shift And 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    pRelease
    ' -------------------------------------------------------------------
    ' ... just in case we are an mdi child and want to tell our parent
    ' ... we done unloaded.
    mdiMain.ChildUnloaded ChildIndex ', mProjectName
'    If Not Parent Is Nothing Then
'    RaiseEvent UnLoaded(ChildIndex)
'    End If
    ClearMemory
End Sub

Private Sub pProcessViewerMenu()
' ... show the viewer pop-up menu.
    PMenu.ShowViewerMenu mHistArray, mWithSyntaxColours, tvProj.Visible, tvMembers.Visible, picTB.Visible, picSB.Visible
End Sub

Private Sub lblClassName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then pProcessViewerMenu
End Sub

Private Sub lblDesc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then pProcessViewerMenu
End Sub

Private Sub picInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then pProcessViewerMenu
End Sub

Private Sub PMenu_MenuItemClick(Caption As String, Menu As Long, Item As Long)
' ... pop-up menu item from the ucMenus user control clicked.
    Select Case Menu
        Case 100
            Select Case Item
                Case 5 ' ... Search Project
                    pSearchProject ' v6
                Case 6 ' ... refresh
                    pRefreshProject ' v5/6
                Case 10 To 14 ' ... open a file in an external program or a new window.
                    pOpenSomething Item
                Case 22 ' ... project report
                    pCreateProjectReport
                Case 23 ' ... api report v5/6
                    pCreateAPIReport
                Case 24
                    pCreateAPIReport 1
                Case 30 ' ... copy project
                    If Not moVBPInfo Is Nothing Then
                        If moVBPInfo.Initialised Then
                            frmCopyProject.LoadVBP moVBPInfo.FileNameAndPath
                            frmCopyProject.Show vbModal
                        End If
                    End If
                Case 31 ' ... copy file
                    pCopyFile
                
                Case 36, 37, 38, 39 ' ... APIs, COnstants, Types & Enumerators.
                    pLoadAPITypeInfo Item
                Case 60 ' ... member dictionary
                    pMemberDictionary
                Case 70
                    pGenerateHTMLMemberHelp
                Case 80 ' ... compile manifest resource
                    If Not moVBPInfo Is Nothing Then
                        If moVBPInfo.Initialised Then
                            If moVBPInfo.HasResource Then
                                MsgBox "Project, " & moVBPInfo.Title & ", already has a resource file" & vbNewLine & _
                                    moVBPInfo.ResFileNameAndPath & vbNewLine & _
                                    "Code Browser will not attempt to overwrite this resource reference", vbExclamation, Caption
                            Else
                                frmCompileResource.SetUpVBP moVBPInfo.FileNameAndPath
                                frmCompileResource.Show vbModal
                            End If
                        End If
                    End If
                Case 101    ' ... load project
                    cmdOpenVBP_Click
                Case 102    ' ... close
                    Unload Me
            End Select
        Case 200
            Select Case Item
                Case 1  ' ... copy sig.
                    pCopySignature
                Case 2 ' ... copy method.
                    pCopyMethod
                Case 21 ' ... quick report.
                    pCreateQuickCodeReport
                Case 3  ' ... refresh class
                    ' ... Busking, v1.
'Dim sFileName As String
'Dim xClassTree As VBClassTree
'Dim xNode As Node
'Dim lIndex As Long
'
'                    If Not moCodeReader Is Nothing Then
'                        If moCodeReader.Initialised Then
'                            sFileName = moCodeReader.FileName
'                            If Len(sFileName) > 0 Then
'                                Set moCodeReader = New CodeInfo
'                                moCodeReader.ReadCodeFile sFileName
'                                If moCodeReader.Initialised Then
'                                    lIndex = -1
'                                    Set xNode = tvMembers.SelectedItem
'                                    If Not xNode Is Nothing Then
'                                        lIndex = xNode.Index
'                                    End If
'                                    Set xClassTree = New VBClassTree
'                                    xClassTree.Init moCodeReader, tvMembers, mdiMain.liMember
'                                    Set xClassTree = Nothing
'                                    If lIndex > -1 Then
'                                        Set xNode = tvMembers.Nodes(lIndex)
'                                        If Not xNode Is Nothing Then
'                                            tvMembers_NodeClick xNode
'                                            xNode.Selected = True
'                                        End If
'                                    End If
'                                    Set xNode = Nothing
'                                End If
'                            End If
'                        End If
'                    End If
'                    MsgBox "Refresh Class"
            End Select
        Case 300
            Select Case Item
                Case 1: cmdViewProj_Click ' show / hide proj.
                Case 2: cmdViewMember_Click ' show / hide memb.
                Case 3: cmdViewToolbar_Click ' show / hide toolbar.
                Case 4: cmdViewStatus_Click ' show / hide status bar.
                Case 5, 6 ' ... interface generation request.
                    pGenerateInterface Item
                Case 33
                    mWithSyntaxColours = Not mWithSyntaxColours
                Case Is > 100 ' ... visited method history.
                    pRevisit (Item - 100)
                
            End Select
    End Select

End Sub

Private Sub pMemberDictionary()
'Dim xDev As DevHelpGen
'
'    If moVBPInfo Is Nothing Then Exit Sub
'    If moVBPInfo.Initialised = False Then Exit Sub
'
'    If Dir$(moVBPInfo.FileNameAndPath) = "" Then Exit Sub
'
'    Set xDev = New DevHelpGen
'    xDev.Init moVBPInfo
'
'    Set xDev = Nothing
    
'Dim oForm As frmProject
'
'    Set oForm = New frmProject
'
'    Set oForm.VBPInfo = moVBPInfo
'
'    oForm.Show vbModal

    pLoadAPITypeInfo 42     ' ... 42, the answer to the meaning of life, the universe and everything :)
    
End Sub

Private Sub pGenerateHTMLMemberHelp()

Dim xDev As DevHelpGen

    If moVBPInfo Is Nothing Then Exit Sub
    If moVBPInfo.Initialised = False Then Exit Sub

    If Dir$(moVBPInfo.FileNameAndPath) = "" Then Exit Sub

    Set xDev = New DevHelpGen
    xDev.Init moVBPInfo

    Set xDev = Nothing

End Sub

Private Sub pLoadAPITypeInfo(ByVal pIndex As Long)
    
' ... v6
' ... load API, Constant, Type or Enum declarations within project.
' ... using an Interface to help things along code wise.

Dim xInterface As IReportForm
Dim sInsert As String

    On Error GoTo ErrHan:
    
    If moVBPInfo Is Nothing Then Exit Sub
    
    Select Case pIndex
        Case 36
            Set xInterface = frmAPIReport: sInsert = "API"
        Case 37
            Set xInterface = frmConstReport: sInsert = "Constant"
        Case 39
            Set xInterface = frmTypesReport: sInsert = "Type"
        Case 38
            Set xInterface = frmEnumsReport: sInsert = "Enum"
        Case 42
'            Set xInterface = frmDictionary: sInsert = "Member"
            Set xInterface = frmMembers: sInsert = "Member"
    End Select
    
    If Not xInterface Is Nothing Then
    
        xInterface.Init moVBPInfo
        
        If xInterface.ItemCount = 0 Then
            
            MsgBox "There were no " & sInsert & " Declarations found in " & moVBPInfo.Title & ".", vbInformation, sInsert & ": " & moVBPInfo.ProjectName
            
            Unload xInterface
        
        Else
        
            xInterface.ZOrder
        
        End If
    
        On Error GoTo 0
        
        Set xInterface = Nothing
    
    End If

ResumeError:
    
    sInsert = vbNullString

Exit Sub

ErrHan:

    Debug.Print "frmViewer.pLoadAPITypeInfo.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub pRefreshProject()
' v5
    If Not moVBPInfo Is Nothing Then
        If moVBPInfo.Initialised Then
            pOpenVBP moVBPInfo.FileNameAndPath
        End If
    End If

End Sub

Private Sub pSearchProject()
' v6
    If Not moVBPInfo Is Nothing Then
        
        If moVBPInfo.Initialised Then
            frmSearchVBProject.LoadVBP moVBPInfo.FileNameAndPath, txtFind.Text
        End If
    
        frmSearchVBProject.Show
        frmSearchVBProject.ZOrder
        
    End If

End Sub

Private Sub pCopyFile()

Dim xItem As Node
Dim sSource As String
Dim sTarget As String
Dim sTmpTag As String
Dim xFileInfo As FileNameInfo
Dim iAnswer As VbMsgBoxResult
Dim sTmp As String
Dim sFolder As String

    On Error GoTo ErrHan:
    
    If tvProj.Nodes.Count = 0 Then Exit Sub
    
    Set xItem = tvProj.SelectedItem
    
    If xItem Is Nothing Then Exit Sub
    
    sTmpTag = xItem.Tag
    
    If sTmpTag <> "" Then
    
        If sTmpTag = cFileSig Then
            
            sSource = xItem.Key
            sFolder = modDialog.GetFolder("Select Target Folder for File Copy" & IIf(Len(mLastCopyFolder), ", last destination was " & mLastCopyFolder, ""), mLastCopyFolder)
            
            If Len(sFolder) Then
                
                mLastCopyFolder = sFolder
                
                If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
                
                modFileName.ParseFileNameEx sSource, xFileInfo
                
                sTarget = sFolder & xFileInfo.File
                
                If Dir$(sTarget, vbNormal) <> "" Then
                
                    iAnswer = MsgBox("The Target File " & vbNewLine & sTarget & vbNewLine & "already Exists." & vbNewLine & "Overwrite this file?", vbQuestion + vbYesNo, Caption)
                    If iAnswer = vbNo Then GoTo ResumeError:
                
                End If
                
                
                FileCopy sSource, sTarget
                
                MsgBox xFileInfo.File & " copied to " & sFolder, vbInformation, Caption
                
                sTmp = UCase$(xFileInfo.Extension)
                
                If sTmp = "FRM" Then
                    sSource = xFileInfo.Path & "\" & xFileInfo.FileName & ".frx"
                    sTarget = sFolder & xFileInfo.FileName & ".frx"
                
                ElseIf sTmp = "CTL" Then
                    sSource = xFileInfo.Path & "\" & xFileInfo.FileName & ".ctx"
                    sTarget = sFolder & xFileInfo.FileName & ".ctx"
                
                ElseIf sTmp = "PAG" Then ' ... v6, add support for property pages
                    ' ... note: would do well to look inside a user control to see
                    ' ... if it has a property page, uc binary file will hold info
                    ' ... about property pages.
                    sSource = xFileInfo.Path & "\" & xFileInfo.FileName & ".pgx"
                    sTarget = sFolder & xFileInfo.FileName & ".pgx"
                
                ElseIf sTmp = "DSR" Then ' ... data designer, data environment / data report.
                    ' ... two binaries with dsrs, dsx and dca
                    ' ... manual override first... dsx.
                    sSource = xFileInfo.Path & "\" & xFileInfo.FileName & ".dsx"
                    sTarget = sFolder & xFileInfo.FileName & ".dsx"
                    If Dir$(sSource, vbNormal) <> "" Then
                        FileCopy sSource, sTarget
                    End If
                    ' ... and fall into second... dca.
                    sSource = xFileInfo.Path & "\" & xFileInfo.FileName & ".dca"
                    sTarget = sFolder & xFileInfo.FileName & ".dca"
                    
                Else
                    GoTo ResumeError:
                    
                End If
                
                If Dir$(sSource, vbNormal) <> "" Then
                    FileCopy sSource, sTarget
                End If
                
            End If
        End If
    End If
    

ResumeError:

    sTarget = vbNullString
    sSource = vbNullString
    sTmpTag = vbNullString
    
    If Not xItem Is Nothing Then
        Set xItem = Nothing
    End If
    
Exit Sub

ErrHan:

    Debug.Print "frmViewer.pCopyFile.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub pCopyMethod()
' ... copy selected method to clipboard
Dim oNode As Node
Dim stext As String
Dim lngKey As Long

    If moCodeReader Is Nothing Then Exit Sub
    If moCodeReader.MemberCount = 0 Then Exit Sub
    
    If Not tvMembers.SelectedItem Is Nothing Then
    
        Set oNode = tvMembers.SelectedItem
        If Right$(oNode.Key, 1) = "x" Then                                ' ... member code.
            lngKey = CLng(Val(oNode.Key))
            stext = moCodeReader.GetMemberCodeLines(lngKey)
            Clipboard.Clear
            Clipboard.SetText stext
            stext = vbNullString
        End If
    End If
    
End Sub

Private Sub pCopySignature()
' ... copy the signature of a method with its attributes
' ... ignoring any code within.

Dim xNode As ComctlLib.Node
Dim tQInfo As QuickMemberInfo
Dim sTmp As String
Dim sAttributes As String
Dim sAttribs() As String
Dim lngAttsLoop As Long
Dim sName As String
Dim lngCount As Long

    On Error GoTo ErrHan:
    
    If moCodeReader Is Nothing Then Exit Sub
    If tvMembers.Nodes.Count = 0 Then Exit Sub
    If tvMembers.SelectedItem Is Nothing Then Exit Sub
    If tvMembers.SelectedItem.Parent Is Nothing Then Exit Sub
                
    Set xNode = tvMembers.SelectedItem
                    
    Select Case xNode.Parent.Key
        ' ... attempt to ensure we have a method.
        Case cSubsNodeKey, cFuncNodeKey, cPropNodeKey
            If Not moCodeReader Is Nothing Then
                ' ... read the quick member info for the it.
                tQInfo = moCodeReader.QuickMember(CLng(Val(xNode.Key)))
                With tQInfo
                    ' ... build the return signature string.
                    sTmp = .Declaration
                    sTmp = sTmp
                    
                    sAttributes = tQInfo.Attribute
                    If Len(sAttributes) Then
                        ' ... build up the attributes string.
                        modStringArrays.SplitString sAttributes, sAttribs, "|" '":"
                        For lngAttsLoop = 0 To UBound(sAttribs)
                            sTmp = sTmp & vbNewLine & "Attribute " & tQInfo.Name & "." & sAttribs(lngAttsLoop)
                        Next lngAttsLoop
                    End If

                    ' ... courtesy line between dec and end.
                    sTmp = sTmp & vbNewLine & vbNewLine
                    
                    ' ... need to derive the end method signature from method type.
                    If .Type = 1 Then
                        sTmp = sTmp & "End Sub"
                    ElseIf .Type = 2 Then
                        sTmp = sTmp & "End Function"
                    ElseIf .Type = 3 Then
                        sTmp = sTmp & "End Property"
                    Else
                        GoTo ResumeError
                    End If
                    
                End With
                
            End If
        Case cTypsNodeKey
        
            ' ... build a Type definition from the information in the node's tag.
            sTmp = xNode.Tag
            If Len(sTmp) Then
                modStrings.SplitStringPair sTmp, ":", sName, sAttributes, True, True
                If Len(sAttributes) Then
                    modStringArrays.SplitString sAttributes, sAttribs, ";", lngCount
                    If lngCount = 0 Then GoTo ResumeError:
                    sTmp = "Type " & sName
                    For lngAttsLoop = 0 To lngCount - 1
                        sTmp = sTmp & vbNewLine & "    " & sAttribs(lngAttsLoop)
                    Next lngAttsLoop
                    sTmp = sTmp & vbNewLine & "End Type" & vbNewLine
                End If
            End If
            
        Case cEnusNodeKey
            
            sTmp = xNode.Key
            If Len(sTmp) Then
                modStrings.SplitStringPair sTmp, ":", sName, sAttributes, True, True
                If Len(sAttributes) Then
                    modStringArrays.SplitString sAttributes, sAttribs, ";", lngCount
                    If lngCount = 0 Then GoTo ResumeError:
                    sTmp = "Enum " & sName
                    For lngAttsLoop = 0 To lngCount - 1
                        sTmp = sTmp & vbNewLine & "    " & sAttribs(lngAttsLoop)
                    Next lngAttsLoop
                    sTmp = sTmp & vbNewLine & "End Enum" & vbNewLine
                End If
            End If
            
        Case cAPIsNodeKey
            sTmp = xNode.Tag
            
    End Select

ResumeError:
                            
    If Len(sTmp) Then
        ' ... done, copy to clip board.
        Clipboard.Clear
        Clipboard.SetText sTmp
    End If
    
    On Error GoTo 0
    If Not xNode Is Nothing Then
        Set xNode = Nothing
    End If
    sTmp = vbNullString
Exit Sub

ErrHan:

    Debug.Print "frmViewer.pCopySignature.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub pOpenSomething(ByVal pRequest As Long, Optional ByVal pTheFileName As String = vbNullString)

Dim xNode As ComctlLib.Node
Dim tFileInfo As FileNameInfo
Dim lngRet As Long
Dim sApp As String
Dim xOptions As cOptions
Dim sFileName As String
Dim sKey As String

' Note:     See Sub: picSB_DblClick for switch options to add to ide instruction.
'           e.g. run, compile ...

    On Error GoTo ErrHan:

    
    If Len(pTheFileName) Then
        
        sFileName = pTheFileName
        
    Else
    
        If tvProj.Nodes.Count > 0 Then
            If Not tvProj.SelectedItem Is Nothing Then
                If pRequest = 10 Then   ' ... open in new window, vbp only.
                                        ' ... need to expand this to accept files within vbp.
                    sKey = tvProj.SelectedItem.Key ' ... v3/4.
                    Set xNode = tvProj.Nodes(1)
                Else
                    Set xNode = tvProj.SelectedItem
                End If
                sFileName = xNode.Key
            End If
        End If
    
    End If
    
    If Dir$(sFileName, vbNormal) = "" Then
        MsgBox "The File:" & vbNewLine & sFileName & vbNewLine & "could not be found.", vbInformation, "Project: Open File"
        GoTo ResumeError:
    End If
    
    
    modFileName.ParseFileNameEx sFileName, tFileInfo
    
    If Len(tFileInfo.Path) Then
    
        Select Case pRequest
            Case 10 ' ... new window.
                    ' ... a bit more expansion required...
                    ' ... need a method to load a vbp and then select an item in the project explorer.
                mdiMain.LoadFile tFileInfo.PathAndName, sKey
            Case 11, 12, 14:
            
                Set xOptions = New cOptions
                xOptions.Read
            
                If pRequest = 11 Then
                    sApp = xOptions.PathToTextEditor ' ... notepad.
                ElseIf pRequest = 12 Then
                    sApp = xOptions.PathToVB6 ' ... ide 1.
                ElseIf pRequest = 14 Then
                    sApp = xOptions.PathToVB5 ' ... ide 2.
                End If
                If Dir$(sApp, vbNormal) <> "" Then
                    lngRet = Shell(sApp & " " & Chr$(34) & tFileInfo.PathAndName & Chr$(34), vbNormalFocus)
                End If
            Case 13 ' ... containing folder.
                lngRet = ShellExecute(0&, vbNullString, tFileInfo.Path & "\", vbNullString, vbNullString, vbNormalFocus)
    
        End Select
    
    End If
    
ResumeError:
    
    On Error GoTo 0
    
    If Not xOptions Is Nothing Then
        Set xOptions = Nothing
    End If
    
Exit Sub

ErrHan:

    Debug.Print "frmViewer.pOpenSomething.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub pGenerateInterface(pIndex As Long)
' ... generate and display the interface to the current code file.
Dim cReport As CodeReport
Dim stext As String

    If Not moCodeReader Is Nothing Then
        
        lblClassName.Caption = lblClassName.Tag & " | Member Interface"
        lblDesc.Caption = moCodeReader.Description
        
        Set cReport = New CodeReport
        Set cReport.CodeReader = moCodeReader
        
        stext = cReport.GenerateInterface(, IIf(pIndex = 6, True, False))
        
        Clipboard.Clear
        Clipboard.SetText stext, vbCFText
        
        stext = modEncode.BuildRTFString(stext, cboFont(1).Text, , cboFont(0).Text, mRTFColourTable, mLineNos, mShowAttributes, mBoldRTF)
        
        rtb(1).TextRTF = stext
        rtb(0).TextRTF = stext
        
    End If
    
    If Not cReport Is Nothing Then
        Set cReport = Nothing
    End If
    
    stext = vbNullString

End Sub

Private Sub rtb_Click(Index As Integer)

' v6, scrolling text member recognition.

' ... when viewing all text, attempt to name the
' ... member where the text is clicked, e.g. from the line clicked.

Dim lngLine As Long
Dim lngStart As Long
Dim sAttributes As String
Dim sName As String

    On Error GoTo ErrHan:
    
    If Not moCodeReader Is Nothing Then
        If tvMembers.Nodes.Count > 0 Then
            ' ... only bother if first node is selected indicating
            ' ... all text is loaded into viewer (minus header section of course).
            If tvMembers.Nodes(1).Selected = False Then Exit Sub
        End If
        ' -------------------------------------------------------------------
        ' ... get line number.
        lngStart = rtb(Index).SelStart
        lngLine = rtb(Index).GetLineFromChar(lngStart)
        lngLine = lngLine + 1
        ' -------------------------------------------------------------------
        ' ... get name of member and any attributes.
        sName = moCodeReader.GetMemberFromLineNo(lngLine, sAttributes)
        ' -------------------------------------------------------------------
        ' ... tidy up raw attributes.
        If Len(sAttributes) > 0 Then
            modStrings.ReplaceChars sAttributes, "VB_Description = ", ""
            modStrings.RemoveQuotes sAttributes
        End If
        ' -------------------------------------------------------------------
        ' ... update info panel descriptive captions.
        lblClassName.Caption = lblClassName.Tag & " | " & sName
        lblDesc.Caption = sAttributes
    End If
    
ResumeError:

    lngLine = 0&
    lngStart = 0&
    
    sName = vbNullString
    sAttributes = vbNullString

Exit Sub

ErrHan:

    Debug.Print "frmViewer.rtb_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
        
End Sub

Private Sub tmrPNode_Timer()
    
' ... this is here to reselect a project node after a new window
' ... has been opened because I couldn't find a quick way to
' ... reset the selected node in the node click event.
    
    tmrPNode.Enabled = False
    If Not moLastProjNode Is Nothing Then
        tvProj.Nodes(moLastProjNode.Index).Selected = True
    End If

End Sub

Private Sub tvMembers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

' Helper: capture right mouse button (see tvMembers_NodeClick).
    
    If Button = 2 Then mRMouseDown = True
    If Button = 2 Then
        PMenu.ShowClassMenu
    End If
    
End Sub

Private Sub tvMembers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mRMouseDown = False
'    If Button = 2 Then
'        PMenu.ShowClassMenu
'    End If
End Sub

Private Sub tvMembers_NodeClick(ByVal Node As ComctlLib.Node)
' ... process class explorer node click.
Dim lngKey As Long
Dim q As QuickMemberInfo
Dim stext As String
Dim lngLineCount As Long
Dim lngTotalLines As Long
Dim lngCommentedLines As Long
Dim iMPointer As VBRUN.MousePointerConstants
Dim bRTBVis(1) As Boolean

    On Error GoTo ErrHan:
    
    Select Case Node.Key
        Case cSubsNodeKey, cPropNodeKey, cFuncNodeKey
            Exit Sub
    End Select
    
    If moCodeReader Is Nothing Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------
    ' ... v7, will this help speed up the rt loading?
    bRTBVis(0) = rtb(0).Visible
    bRTBVis(1) = rtb(1).Visible
    
    If bRTBVis(0) = True Then rtb(0).Visible = False
    If bRTBVis(1) = True Then rtb(1).Visible = False
    
    If mLoadingClass = False Then
'        lblDesc.Caption = ""
        ' -------------------------------------------------------------------
        iMPointer = Screen.MousePointer
        If Not iMPointer = VBRUN.MousePointerConstants.vbHourglass Then Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
        ' -------------------------------------------------------------------
        lngTotalLines = moCodeReader.GetLineCount
        lngCommentedLines = moCodeReader.CountCommentLines
        ' -------------------------------------------------------------------
        If Node.Key = cMainNodeKey Then             ' ... reads entire code amd declarations.
                                                    
            stext = moCodeReader.GetDecsAndCode
            lngLineCount = moCodeReader.GetLineCount
            lblClassName.Caption = lblClassName.Tag & " | Declarations and Members"
            lblDesc.Caption = moCodeReader.Description
            
        ElseIf Node.Key = cDecsNodeKey Then         ' ... reads declarations only.
        
            stext = moCodeReader.GetDeclarations
            lngLineCount = moCodeReader.GetDeclarationsLineCount
            lblClassName.Caption = lblClassName.Tag & " | Declarations"
            lblDesc.Caption = "Declarative Section of Source File."
            
        ElseIf Node.Key = cHeadNodeKey Then
            ' -------------------------------------------------------------------
            ' v6, Source Header.
            ' ... include header section, not coloured by default.
            stext = moCodeReader.GetHeader
            lngLineCount = moCodeReader.GetHeaderLineCount ' v8
            
            lblClassName.Caption = lblClassName.Tag & " | Header"
            lblDesc.Caption = "Header Section of Source File."
            
            rtb(0).Text = stext: rtb(0).Font.Size = Val(cboFont(0).Text): rtb(0).SelStart = 0: rtb(0).SelLength = Len(stext): rtb(0).SelColor = vbBlack: rtb(0).SelLength = 0: rtb(0).SelStart = 0
            rtb(1).Text = stext: rtb(1).Font.Size = Val(cboFont(0).Text): rtb(1).SelStart = 0: rtb(1).SelLength = Len(stext): rtb(1).SelColor = vbBlack: rtb(1).SelLength = 0: rtb(1).SelStart = 0
            
            GoTo SkipEncode:
            
        ElseIf Right$(Node.Key, 1) = "x" Then   ' ... reads member code.
        
            pAddHistory Node    ' ... add item to history.
            lngKey = CLng(Val(Node.Key))
            stext = moCodeReader.GetMemberCodeLines(lngKey)
        
        Else
            GoTo ResumeError:
        End If
        
        ' -------------------------------------------------------------------
        ' v6, attempt control rtf encoding, re: TonyYong, Windows(Chinese).
        ' ... if no encoding, is it ok with Chinese characters?
        If mWithSyntaxColours Then
            stext = modEncode.BuildRTFString(stext, cboFont(1).Text, , cboFont(0).Text, mRTFColourTable, mLineNos, mShowAttributes, mBoldRTF)
            ' -------------------------------------------------------------------
            If mRMouseDown Then
                ' ... when right btn load into back rt box.
                rtb(0).TextRTF = stext: rtb(0).SelStart = 0
            Else
                rtb(1).TextRTF = stext: rtb(1).SelStart = 0
                rtb(0).TextRTF = stext: rtb(0).SelStart = 0
            End If
        
        Else
            If mRMouseDown Then
                ' ... when right btn load into back rt box.
                rtb(0).Text = stext: rtb(0).Font.Size = Val(cboFont(0).Text): rtb(0).SelStart = 0: rtb(0).SelLength = Len(stext): rtb(0).SelColor = vbBlack: rtb(0).SelLength = 0: rtb(0).SelStart = 0
            Else
                rtb(0).Text = stext: rtb(0).Font.Size = Val(cboFont(0).Text): rtb(0).SelStart = 0: rtb(0).SelLength = Len(stext): rtb(0).SelColor = vbBlack: rtb(0).SelLength = 0: rtb(0).SelStart = 0
                rtb(1).Text = stext: rtb(1).Font.Size = Val(cboFont(0).Text): rtb(1).SelStart = 0: rtb(1).SelLength = Len(stext): rtb(1).SelColor = vbBlack: rtb(1).SelLength = 0: rtb(1).SelStart = 0
            End If
        
        End If
        ' -------------------------------------------------------------------
'        chkAlign_Click ' ... k, so need a way to avoid this unnecessarily.
        ' -------------------------------------------------------------------
    
    End If
    
SkipEncode:

    If Right$(Node.Key, 1) = "x" Then       ' ... Member. note: gonna hit on the tag in future.
                
        lngKey = CLng(Val(Node.Key))
        q = moCodeReader.QuickMember(lngKey) ' ... get the quick member for the the method selected.
        
        lngLineCount = q.LineCount
        
        If Len(Node.Key) And mLoadingClass = False Then
            ' ... Info Panel: describing location in code file.
            lblClassName.Caption = lblClassName.Tag & " | " & Choose(q.Type, "Sub", "Function", "Property") & ": " & Node.Text
            stext = q.Attribute
            modStrings.ReplaceChars stext, "VB_Description = ", vbNullString
            modStrings.RemoveQuotes stext
            lblDesc.Caption = stext
        End If
        
    End If
ResumeError:
    ' -------------------------------------------------------------------
    If Node.Key <> cHeadNodeKey Then ' v8 condition added, displaying header line count
        If lngTotalLines = lngLineCount Then
            lblLineCount.Caption = "Total Lines: " & Format$(lngTotalLines, cNumFormat) ' & " ~ [ " & Format$(lngTotalLines - lngCommentedLines, cNumFormat) & " not commented ]"
        Else
            lblLineCount.Caption = "Lines: " & Format$(lngLineCount, cNumFormat) & " of " & Format$(lngTotalLines, cNumFormat)
        End If
    Else
        lblLineCount.Caption = "Hdr Lines: " & Format$(lngLineCount, cNumFormat)
    End If
    
    ' -------------------------------------------------------------------
    ' ... v7
    If bRTBVis(0) = True Then rtb(0).Visible = True
    If bRTBVis(1) = True Then rtb(1).Visible = True
    
    ClearMemory
    
    ' -------------------------------------------------------------------
    If Not iMPointer = Screen.MousePointer Then Screen.MousePointer = iMPointer ' VBRUN.MousePointerConstants.vbHourglass
    ' -------------------------------------------------------------------
Exit Sub
ErrHan:
    Debug.Print "frmViewer.tvMembers_NodeClick.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

Private Sub pUpdateCaption()

' ... note:
' ... this was an after thought to adding multiple child windows
' ... to indicate that they are in use or not.
' ... quick solution, v3/4; didn't know what else to do.

    If mUseChildWindows Then
        If InStr(1, Caption, " +") = 0 Then
            Caption = Caption & " +"
        End If
    End If

End Sub

Private Sub pProcessClass(Optional ByVal Node As ComctlLib.Node = Nothing, Optional pProcessRTF As Boolean = False)
' -------------------------------------------------------------------
Dim sToolTip As String
Dim sKey As String
Dim oMArray As StringArray
Dim otmpArray As StringArray
Dim sFile As String
Dim srtfText As String
Dim bDo As Boolean
Dim iMPointer As VBRUN.MousePointerConstants
Dim sMainNodeCaption As String
Dim sClassName As String
Dim sErr As String
Dim sPType As String ' ... parent node key.

Dim xCT As VBClassTree

' ... Class Explorer: Load/Reload Member Nodes from Project item or plain text in viewer.
    
    On Error GoTo ErrHan:
    mLoadingClass = True
    ' -------------------------------------------------------------------
    lblClassName.Caption = "":              lblClassName.Tag = ""
    m_CurrentTreeGUID = vbNullString:       m_CurrentTreeFileName = vbNullString
    m_CurrentTreeText = "Unknown Source":   sClassName = m_CurrentTreeText
    ' -------------------------------------------------------------------
    If Not moCodeReader Is Nothing Then
        Set moCodeReader = Nothing
    End If
    ' -------------------------------------------------------------------
    If Not Node Is Nothing Then
        sKey = Node.Key
        If Node.Tag = cFileSig Then
            ' -------------------------------------------------------------------
            If Not Node.Parent Is Nothing Then
                sPType = Node.Parent.Key
            End If
            ' -------------------------------------------------------------------
            sFile = sKey
            m_CurrentTreeFileName = sFile
            ' -------------------------------------------------------------------
            sMainNodeCaption = Node.Text
            m_CurrentTreeText = Node.Text
            ' -------------------------------------------------------------------
            If tvProj.Nodes.Count > 0 Then
                sClassName = tvProj.Nodes(1).Text
            End If
            sClassName = sClassName & " | " & Trim$(sPType) & " | " & Node.Text
            Caption = sClassName
            pUpdateCaption
            bDo = True
        Else
            sErr = "Process Class: Unknown Token, unable to process request." ' ... no recognised file descriptor.
        End If
    Else
        ' ... no node.
        If pProcessRTF = True Then
            ' ... problem if text is already rtf encoded.
            srtfText = rtb(1).Text
            If Len(srtfText) > 0 Then
                sMainNodeCaption = "Current Viewer Text"
                bDo = True
            Else
                sErr = "Process Class: No Raw Text to Parse, unable to process request." ' ... no text.
            End If
        Else
            sErr = "Process Class: No known Parser requested, unable to process request." ' ... no instruction.
        End If
    End If
    
    ' -------------------------------------------------------------------
    lblClassName.Tag = sClassName
    sClassName = sClassName & " | Declarations"
    lblClassName.Caption = sClassName
    
    ' -------------------------------------------------------------------
    If bDo = True Then
        ' -------------------------------------------------------------------
        
        mCodeFileName = vbNullString
        
        Set moCodeReader = New CodeInfo
        
        If Len(sFile) > 0 Then
            mCodeFileName = sFile
            moCodeReader.ReadCodeFile sFile
            lblDesc.Caption = moCodeReader.Description
        ElseIf Len(srtfText) > 0 Then
            moCodeReader.ReadCodeString srtfText
        End If
        moCodeReader.Declarations   ' ... need to call this for now to process declarations section of code file.
'        If Len(moCodeReader.MenuStructure) Then
'            MsgBox moCodeReader.MenuStructure, vbInformation, "Menu Found"
'            MsgBox moCodeReader.MenuMethods
'            Debug.Print moCodeReader.MenuMethods
'        End If
        ' -------------------------------------------------------------------
        ' ... v7/8 test, print variables declared.
'        Debug.Print moCodeReader.VarsString
        ' -------------------------------------------------------------------
        
        Set xCT = New VBClassTree
        xCT.Init moCodeReader, tvMembers, mdiMain.liMember, mShowClassHeadCount
        Set xCT = Nothing
        
        ' -------------------------------------------------------------------
        iMPointer = Screen.MousePointer
        If iMPointer <> VBRUN.MousePointerConstants.vbHourglass Then
            Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
        End If
    Else
        If Len(sErr) > 0 Then
            MsgBox sErr, vbInformation, "Request not processed."
        End If
    End If
ResumeErr:
    On Error GoTo 0
    picSB.ToolTipText = sToolTip
    picSB_Paint
    
    mLoadingClass = False
    
    ' v5 fix, presumed we had a declarations node :|
    If tvMembers.Nodes.Count > 0 Then
        If mCodeLoadAll = True Then
            If Not tvMembers.Nodes(1) Is Nothing Then
                tvMembers_NodeClick tvMembers.Nodes(1)
            End If
        Else
            If Not tvMembers.Nodes(cDecsNodeKey) Is Nothing Then
                tvMembers_NodeClick tvMembers.Nodes(cDecsNodeKey)
            End If
        End If
    End If
    
    Set oMArray = Nothing
    If Not otmpArray Is Nothing Then Set otmpArray = Nothing
    If iMPointer <> Screen.MousePointer Then
        Screen.MousePointer = iMPointer
    End If
Exit Sub
ErrHan:
    Debug.Print "frmViewer.pProcessClass.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeErr:
'    Resume
    
End Sub ' ... pProcessClass:

Private Sub tvMembers_OLECompleteDrag(Effect As Long)
    mMemberDrag = False ' v8
End Sub

Private Sub tvProj_DblClick()

Dim xNode As Node
Dim lngFound As Long
Dim sTmp As String
Dim sPNodeTag As String
Dim dShell As Double

    Set xNode = tvProj.SelectedItem
    
    If Not xNode Is Nothing Then
        ' -------------------------------------------------------------------
        ' v8, if proj is type exe and there's an exe file in vbp folder then try to run it :)
        If xNode.Key = cExeNodeKey Then
            If Not moVBPInfo Is Nothing Then
                If moVBPInfo.IsExe Then
                    sTmp = moVBPInfo.FilePath & "\" & moVBPInfo.ExeName32
                    If modFileName.FileExists(sTmp) Then
                        On Error Resume Next
                        dShell = Shell(sTmp, vbNormalFocus)
                        If Err.Number <> 0 Then
                            Err.Clear
                            MsgBox "Was not able to run ' " & sTmp & " '", vbInformation, "Run Project Exe"
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
        ' -------------------------------------------------------------------
        If Not xNode.Parent Is Nothing Then
            
            sPNodeTag = Trim$(xNode.Parent.Tag)
            
            If sPNodeTag = "Components" Or sPNodeTag = "References" Then
            
                sTmp = xNode.Key    ' ... key = GUID + " " + File Name / File Name and Path
                
                lngFound = InStr(1, sTmp, " ")
                If lngFound > 0 Then
                    sTmp = Mid$(sTmp, lngFound + 1)
                End If
                
            End If
            
            If Len(sTmp) Then
                frmRefViewer.LoadRef sTmp
                frmRefViewer.Show
                frmRefViewer.ZOrder
            
            End If
            
        End If
    
    End If

    If Not xNode Is Nothing Then Set xNode = Nothing

End Sub

Private Sub tvProj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PMenu.ShowProjectMenu
    End If

End Sub

Private Sub tvProj_NodeClick(ByVal Node As ComctlLib.Node)

Dim sToolTip As String
Dim sKey As String
Dim oMArray As StringArray
Dim otmpArray As StringArray
Dim xOptions As cOptions
Dim bHideProject As Boolean
Dim bHideToolbar As Boolean

' ... Project Node: Load/Reload Member Nodes from the Project item.
    
    On Error GoTo ErrHan:
'    Debug.Print Node.Text & ": " & Node.Key & ": " & Node.Tag
    ' ... k, so need to resolve what to do when a new file is clicked
    ' ... if we have a file open then might be polite to open up a new window
    ' ... keeping current one intact without losing it's history.
    ' ... we can glean whether the code reader is instanced and initialised
    ' ... and, now, we can check the last node that was clicked (moLastProjNode).
    ' ... cFileSig tells us that the node represents a file and if the file isn't
    ' ... a .res then the program can load it.
    Set xOptions = New cOptions
    xOptions.Read
    If mUseChildWindows Then
        If Node.Tag = cFileSig Then
            If Not moCodeReader Is Nothing Then
                If moCodeReader.Initialised = True Then
                    If Node.Text <> moCodeReader.Name Then
                        If mUseChildWindows Then
                            bHideProject = xOptions.HideChildProject
                            bHideToolbar = xOptions.HideChildToolbar
                        End If
                        mdiMain.LoadFile moVBPInfo.FileNameAndPath, Node.Key, bHideProject, bHideToolbar
                        If Not moLastProjNode Is Nothing Then
                            ' ... restore appearance of last selected node.
                            ' ... need to quit this method before being able to
                            ' ... reset the selected node other than this one.
                            tmrPNode.Enabled = True
                        End If
                        GoTo SkipUseChild:
                    End If
                End If
            End If
        End If
    End If
    
    lblClassName.Caption = mProjectName
    lblClassName.Tag = mProjectName
    lblDesc.Caption = mProjectDesc

    sKey = Node.Key
    
    m_CurrentTreeGUID = vbNullString
    m_CurrentTreeFileName = vbNullString
    m_CurrentTreeText = Node.Text
    
    If Node.Tag = cFileSig Then
        m_CurrentTreeFileName = sKey
        If Not Node.Parent Is Nothing Then
            If Node.Parent.Key = cRDocNodeKey Then
                GoTo ResumeErr:
            End If
        End If
        sToolTip = "Double-Click to open file in NotePad or Ctrl+ Double-Click to open in VB6 IDE (if possible, check program paths)"
        If Right$(LCase$(sKey), 3) <> "res" Then
            Set moLastProjNode = Node
            pProcessClass Node
            pClearHistory
        End If
    ElseIf Node.Tag = cGUIDSig Then
        m_CurrentTreeGUID = sKey
        sToolTip = "Double-Click to send info to an InputBox for copying"
    ' -------------------------------------------------------------------
    ' v5
    ElseIf Node.Tag = cMissingFileSig Then
        m_CurrentTreeFileName = sKey
        m_CurrentTreeGUID = sKey
    Else
        If Node.Index = 1 Then
            m_CurrentTreeFileName = Node.Key
        Else
            m_CurrentTreeFileName = tvProj.Nodes(1).Key
        End If
    End If
    
SkipUseChild:

ResumeErr:
    On Error GoTo 0
    Set xOptions = Nothing
    picSB.ToolTipText = sToolTip
    picSB_Paint
    Set oMArray = Nothing
    If Not otmpArray Is Nothing Then otmpArray.Clear:
    Set otmpArray = Nothing
    ClearMemory
    Screen.MousePointer = vbDefault
Exit Sub
ErrHan:
    Debug.Print "frmViewer.tvProj_NodeClick.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeErr:
    Resume
End Sub

Private Sub cmdOpenVBP_Click()
    pOpenVBP
End Sub

Public Function pStripEndString(TheString As String, TheDelimiter As String) As String
Dim sReturn As String
Dim sChar As String
Dim lngLoop As Long
Dim lngLength As Long
' -------------------------------------------------------------------
' ... helper to strip text from the end of a string following the last instance of a delimiter.
' -------------------------------------------------------------------
    lngLength = Len(TheString)
    sReturn = TheString
    For lngLoop = lngLength To 1 Step -1
        sChar = Mid$(TheString, lngLoop, 1)
        If sChar = TheDelimiter Then
            sReturn = Mid$(TheString, lngLoop + 1, lngLength - lngLoop + 1)
            Exit For
        End If
    Next lngLoop
    pStripEndString = sReturn
End Function ' ... pStripEndString: String

Private Sub pOpenVBP(Optional pFileName As String = vbNullString, _
                     Optional pKey As String = vbNullString, _
                     Optional pHideProject As Boolean = False, _
                     Optional pHideToolbar As Boolean = False, _
                     Optional ByVal pPSCReadMeText As String = vbNullString)

Dim xTree As VBProjTree
Dim sFileName As String
Dim sFilter As String
Dim sDialogTitle As String
' ... v3/4, load vbp & select node.
Dim oNode As Node
Dim xFileInfo As FileNameInfo

Dim xProj As VBProject

'Dim bVisible As Boolean
' -------------------------------------------------------------------
' ... Helper: Load Open Dialog / or use pFileName and parse vbp adding stuff to explorer tree view.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    
    mProjectName = ""
    mProjectDesc = ""
    
    lblFind(0).Caption = ""
    lblFind(1).Caption = ""
    
    ' -------------------------------------------------------------------
    If Len(Trim$(pFileName)) = 0 Then
        ' -------------------------------------------------------------------
        sDialogTitle = "Select Visual Basic Project (5 or 6)"
        sFilter = modDialog.MakeDialogFilter("VB Project", "*", "vbp")
'        sFilter = sFilter & "|" & modDialog.MakeDialogFilter("VB Project Group", "*", "vbg")
'        sFilter = sFilter & "|" & modDialog.MakeDialogMultiFilter("VB Files", "frm", "cls", "bas", "ctl")
        sFileName = ""
        
        sFileName = modDialog.GetOpenFileName(, , sFilter, , mLastOpenFolder, sDialogTitle)
        
        If Len(sFileName) = 0 Then
            Err.Raise cDlgCancelErr
        End If
        
        modFileName.ParseFileNameEx sFileName, xFileInfo
        mLastOpenFolder = xFileInfo.Path
        
    Else
        sFileName = pFileName
    
    End If
    ' -------------------------------------------------------------------
    mF2LastNode = vbNullString   ' ... helper field used with Shift+F2 keypress.
    ' -------------------------------------------------------------------
    
    mPSCReadMeText = pPSCReadMeText
    
    Set moVBPInfo = New VBPInfo
    moVBPInfo.ReadVBP sFileName
    
    Set xTree = New VBProjTree
    xTree.Init moVBPInfo, tvProj, mShowClassHeadCount
    
    ' -------------------------------------------------------------------
    Caption = moVBPInfo.ProjectName '& "." & sFileExt & ": " & moVBPInfo.Description '"Code Viewer: " & v.ProjectName & "." & sFileExt & " # " & v.Description
    mProjectName = moVBPInfo.ProjectName
    mProjectDesc = moVBPInfo.Description
    ' -------------------------------------------------------------------
    If Len(pKey) Then   ' ... v3/4, select child node.
        ' ... parent window using multiple child windows
        ' ... so inherit attribute.
        mUseChildWindows = True
        Set oNode = tvProj.Nodes(pKey)
        If Not oNode Is Nothing Then
            If Not oNode.Parent Is Nothing Then
                oNode.Parent.Expanded = True
            End If
            tvProj_NodeClick oNode
            oNode.Selected = True
            oNode.Expanded = True
            ' ... if its a code file then hide the project tree.
            tvProj.Visible = Not pHideProject
            picTB.Visible = Not pHideToolbar
            On Error Resume Next
            If pHideProject = False Then
                tvMembers.SetFocus
            End If
        End If
    End If
    
    pUpdateCaption
    
    If Len(mPSCReadMeText) = 0 Then pCheck4PSCReadMe (sFileName)
    
    imgBtn(14).Visible = Len(mPSCReadMeText)
    
    If tvProj.Nodes.Count > 0 Then tvProj_NodeClick tvProj.Nodes(1)
    ' -------------------------------------------------------------------
    ' v 5
    If moVBPInfo.MissingCount > 0 Then
        MsgBox "Some VBP Member Files were not found where expected:" & _
                vbNewLine & "Count: " & Format$(moVBPInfo.MissingCount, cNumFormat) & _
                vbNewLine & "Check VBP and Folders, references may just be wrong." & _
                vbNewLine & _
                vbNewLine & moVBPInfo.MissingFiles.ToString("", "", , True), vbExclamation, "Unaccounted Project Files"
    End If
    
'    Set xProj = New VBProject
'    xProj.Init moVBPInfo
    mMemberDrag = False ' v8
ResumeError:
Exit Sub
ErrHan:
    
    If Err.Number <> cDlgCancelErr Then
        Debug.Print "frmMain.cmdOpenVBP_Click.Error: " & Err.Number & "; " & Err.Description
    End If
    Resume ResumeError:

End Sub ' ... pOpenVBP:

Private Sub pCheck4PSCReadMe(ByVal pVBPFileName As String)
Dim x As FileNameInfo
Dim s As String

    If Len(pVBPFileName) = 0 Then Exit Sub
    
    modFileName.ParseFileNameEx pVBPFileName, x
    
    s = x.Path & "\@PSC_ReadMe*.txt"
    
    If Dir$(s, vbNormal) <> "" Then
        s = Dir$(s, vbNormal)
        s = modReader.ReadFile(s)
        mPSCReadMeText = s
    End If
    
    s = vbNullString
    
End Sub

Private Sub pRelease()
' -------------------------------------------------------------------
' Helper:   Release all current resources
' -------------------------------------------------------------------
    On Error GoTo 0
    tmrPNode.Enabled = False
    If Not moVBPInfo Is Nothing Then
        Set moVBPInfo = Nothing
    End If
    If Not moCodeReader Is Nothing Then
        Set moCodeReader = Nothing
    End If
    If Not mHistArray Is Nothing Then
        Set mHistArray = Nothing
    End If
    If Not moLastProjNode Is Nothing Then
        Set moLastProjNode = Nothing
    End If
    
    mLastCopyFolder = vbNullString
    mLastSaveFolder = vbNullString
    
End Sub
' -------------------------------------------------------------------
' ... General GUI.
' ... (Image) Buttons.
Private Sub imgBtn_Click(Index As Integer)
' -------------------------------------------------------------------
' ... Toolbar Image: Capture and respond to button click request.
' -------------------------------------------------------------------
    Select Case Index
        Case 0: cmdOpenVBP_Click
        Case 1: cmdViewProj_Click
        Case 2: cmdViewMember_Click
        Case 3: cmdViewToolbar_Click
        Case 4: cmdViewStatus_Click
        Case 5: pPrint
        Case 6: pColouriseText
        Case 7: cmdSaveRTF_Click
        Case 8: pSaveToHTML
        Case 9: cmdFind_Click
        Case 10: pShowHelp
        Case 11: pCreateQuickCodeReport
        Case 12: pHistBack
        Case 13: pHistFwd
        Case 14: pShowPSCReadMe
    End Select
End Sub

Private Sub pShowPSCReadMe()
    If Len(mPSCReadMeText) Then
        frmPSCReadMe.LoadReadMe mPSCReadMeText
        frmPSCReadMe.Show vbModal
    End If
End Sub

Private Sub pRevisit(pIndex As Long, Optional pFirstLine As Long = 0, Optional pFirstChar As Long = 0, Optional pSelLength As Long = 0)
' ... force node click on visited node.
Dim oNode As Node
    On Error GoTo ErrHan:
    If tvMembers.Nodes.Count > 0 Then
        If Not tvMembers.Nodes(pIndex) Is Nothing Then
            Set oNode = tvMembers.Nodes(pIndex)
            If Not tvMembers.SelectedItem Is Nothing Then
                If tvMembers.SelectedItem.Index = oNode.Index Then
                    GoTo ResErr:    ' ... oNode is Current Node.
                End If
                mMovingBack = True
                tvMembers_NodeClick oNode
                oNode.Selected = True
                If pFirstLine > 0 Then
'                    SendMessageLong rtb(1).hwnd, EM_LINESCROLL, 0&, pFirstLine
                    modGeneral.ScrollRTFBox rtb(1).hwnd, pFirstLine ' v6 added, swapped with above.
                End If
                If pSelLength > 0 Then
                    rtb(1).SelLength = pSelLength
                End If
                If pFirstChar > 0 Then
                    rtb(1).SelStart = pFirstChar
                    rtb(1).SetFocus
                End If
                mMovingBack = False
            End If
        End If
    End If
ResErr:
    If Not oNode Is Nothing Then Set oNode = Nothing
Exit Sub
ErrHan:
    Debug.Print "frmViewer.pRevisit.Error: " & Err.Description
    Err.Clear
    Resume ResErr:
End Sub

Private Sub pHistFwd()
' ... move forward in history.
Dim sTmpA As StringArray
Dim lngIndex As Long
Dim lngCount As Long
Dim lngFirstLine As Long
Dim lngFirstChar As Long
Dim lngSelLen As Long

    If Not mHistArray Is Nothing Then
        lngCount = mHistArray.Count
        If mHistIndex < lngCount Then
            mHistIndex = mHistIndex + 1
            Set sTmpA = mHistArray.ItemAsStringArray(mHistIndex, Chr$(0))
            lngIndex = CLng(sTmpA.ItemAsNumberValue(1))
            lngFirstLine = CLng(sTmpA.ItemAsNumberValue(3))
            lngFirstChar = CLng(sTmpA.ItemAsNumberValue(4))
            lngSelLen = CLng(sTmpA.ItemAsNumberValue(5))
            pRevisit lngIndex, lngFirstLine, lngFirstChar
        End If
    End If
    
    pUpdateHistBtns lngCount

End Sub

Private Sub pUpdateHistBtns(ByVal pCount As Long)
' ... toggle the visible state of the history buttons.
    imgBtn(12).Visible = pCount > 0 And mHistIndex > 1
    imgBtn(13).Visible = mHistIndex > 0 And mHistIndex < pCount

End Sub

Private Sub pHistBack()
' ... move bacward in history.
Dim sTmpA As StringArray
Dim lngIndex As Long
Dim lngCount As Long
Dim lngFirstLine As Long
Dim lngFirstChar As Long
Dim lngSelLen As Long

    If Not mHistArray Is Nothing Then
        lngCount = mHistArray.Count
        If mHistIndex > 1 Then
            mHistIndex = mHistIndex - 1
            Set sTmpA = mHistArray.ItemAsStringArray(mHistIndex, Chr$(0))
            lngIndex = CLng(sTmpA.ItemAsNumberValue(1))
            lngFirstLine = CLng(sTmpA.ItemAsNumberValue(3))
            lngFirstChar = CLng(sTmpA.ItemAsNumberValue(4))
            lngSelLen = CLng(sTmpA.ItemAsNumberValue(5))
            pRevisit lngIndex, lngFirstLine, lngFirstChar
        End If
    End If
    pUpdateHistBtns lngCount
    
End Sub

Private Sub pCreateAPIReport(Optional pType As Long = 0)

Dim cPRep As ProjectReport
Dim stext As String

    On Error GoTo ErrHan:
    
    Screen.MousePointer = vbHourglass
    
    If Not moVBPInfo Is Nothing Then
        
        Set cPRep = New ProjectReport
        cPRep.Init moVBPInfo

        cPRep.GenerateAPIReport stext, pType

        If Len(stext) Then

            If pType = 1 Then
                stext = "Distinct API Report for " & moVBPInfo.Title & vbNewLine & vbNewLine & stext
            ElseIf pType = 0 Then
                stext = "Project API Report for " & moVBPInfo.Title & vbNewLine & vbNewLine & stext
            End If
            stext = modEncode.BuildRTFString(stext, cboFont(1).Text, , cboFont(0).Text, mRTFColourTable, mLineNos, mShowAttributes, mBoldRTF)
            rtb(1).TextRTF = stext
            rtb(0).TextRTF = stext
        End If
        
'        frmAPIReport.Show: frmAPIReport.ZOrder: DoEvents
'        frmAPIReport.Init moVBPInfo
'
'        frmConstReport.Show: frmConstReport.ZOrder: DoEvents
'        frmConstReport.Init moVBPInfo
'
'        frmTypesReport.Show: frmTypesReport.ZOrder: DoEvents
'        frmTypesReport.Init moVBPInfo
'
'        frmEnumsReport.Show: frmEnumsReport.ZOrder: DoEvents
'        frmEnumsReport.Init moVBPInfo
        
    End If

ResumeError:
    
    If Not cPRep Is Nothing Then
        
        Set cPRep = Nothing
    
    End If

    Screen.MousePointer = vbDefault
    
Exit Sub

ErrHan:

    Debug.Print "frmViewer.pCreateAPIReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub pCreateProjectReport()

Dim cPRep As ProjectReport
Dim stext As String
Dim sColourTable As String

    On Error GoTo ErrHan:
    
    Screen.MousePointer = vbHourglass
    
    If Not moVBPInfo Is Nothing Then
        
        Set cPRep = New ProjectReport
        cPRep.Init moVBPInfo
        
        cPRep.GenerateVBPReport stext
        
        If Len(stext) Then
        
            sColourTable = modEncode.AllBlackFontColours
            stext = modEncode.BuildRTFString(stext, cboFont(1).Text, , cboFont(0).Text, sColourTable, mLineNos, mShowAttributes, mBoldRTF)
            rtb(1).TextRTF = stext
            rtb(0).TextRTF = stext
        
        End If
        
    End If

ResumeError:
    
    If Not cPRep Is Nothing Then
        
        Set cPRep = Nothing
    
    End If

    Screen.MousePointer = vbDefault
    
Exit Sub

ErrHan:

    Debug.Print "frmViewer.pCreateProjectReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Sub

Private Sub pCreateQuickCodeReport()
' ... generate and display the quick report for the cuirrent code file.
Dim cReport As CodeReport
Dim stext As String
Dim sColourTable As String
    
    If Not moCodeReader Is Nothing Then
        
        lblClassName.Caption = lblClassName.Tag & " | Member Summary"
        lblDesc.Caption = moCodeReader.Description
        
        Set cReport = New CodeReport
        Set cReport.CodeReader = moCodeReader
        
        sColourTable = modEncode.AllBlackFontColours 'v8
        stext = cReport.ReportString
        stext = modEncode.BuildRTFString(stext, cboFont(1).Text, , cboFont(0).Text, sColourTable, mLineNos, mShowAttributes, mBoldRTF)
        
        rtb(1).TextRTF = stext
        rtb(0).TextRTF = stext
        
    End If
    
    If Not cReport Is Nothing Then
        Set cReport = Nothing
    End If

End Sub

Private Sub pSaveToHTML()
' ... encode and save viewer text to html.
Dim stext As String
Dim xCR As CodeInfo
Dim sX As StringWorker
Dim bOK As Boolean
Dim sErrMsg As String
Dim sMsg As String
Dim xFileInfo As FileNameInfo

' -------------------------------------------------------------------
' Helper:   Save current viewer text or selected viewer text as HTML.
' Note:     Not complete, no choice over saved file name / location.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    
    If Len(rtb(1).Text) = 0 Then
        MsgBox "Not Saved." & vbNewLine & "Nothing to Save yet.", vbExclamation, "Not Saved"
        Exit Sub
    End If
    
    With cdgHTM
        
        .DialogTitle = "Save As HTML"
        .Filter = "HTML (*.htm) | *.htm"
        
        If Len(mLastSaveFolder) Then
            .InitDir = mLastSaveFolder
        End If
        
        .ShowSave
        
        modFileName.ParseFileNameEx .FileName, xFileInfo
        mLastSaveFolder = xFileInfo.Path
        
    End With
    
    If rtb(1).SelLength > 0 Then
        ' ... just save the selection not the whole text.
        stext = Mid$(rtb(1).Text, rtb(1).SelStart + 1, rtb(1).SelLength)
    Else
        stext = rtb(1).Text
    End If
    
    If Len(stext) Then
    
        Set xCR = New CodeInfo
        Set sX = New StringWorker
        sX = modEncode.BuildHTMLString(stext, cboFont(1).Text, , cboFont(0).Text)
        
'        sX.ToFile cdgHTM.Filename, , , True, bOK, sErrMsg
        sX.ToFile xFileInfo.PathAndName, , , True, bOK, sErrMsg
        
    End If
    

ResumeError:
        
    On Error GoTo 0
    
    If Not xCR Is Nothing Then
        Set xCR = Nothing
    End If
    If Not sX Is Nothing Then
        Set sX = Nothing
    End If
    stext = vbNullString
    
    If bOK = True Then
        sMsg = "File Saved" & vbNewLine & cdgHTM.FileName
    Else
        sMsg = "File Not Saved:" & vbCrLf & sErrMsg
    End If
    
    MsgBox sMsg, vbOKOnly + vbInformation, "Saved?"

Exit Sub

ErrHan:
    If Err.Number <> cDlgCancelErr Then
        Debug.Print "frmViewer.pSaveToHTML.Error: " & Err.Number & "; " & Err.Description
    End If
    GoTo ResumeError:

End Sub

Private Sub pColouriseText()
' -------------------------------------------------------------------
' Helper:   Instruction to digest viewer text to CodeInfo.
' -------------------------------------------------------------------
    pProcessClass , True
End Sub

Private Sub imgBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' -------------------------------------------------------------------
' ... Toolbar Image: Shift Right and Down.
' -------------------------------------------------------------------
    If mLMouseDown Then Exit Sub
    imgBtn(Index).Move imgBtn(Index).Left + 15, imgBtn(Index).Top + 15
    mLMouseDown = True
End Sub

Private Sub imgBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' -------------------------------------------------------------------
' ... Toolbar Image: Shift Left and Up.
' -------------------------------------------------------------------
    If mLMouseDown = False Then Exit Sub
    imgBtn(Index).Move imgBtn(Index).Left - 15, imgBtn(Index).Top - 15
    mLMouseDown = False
End Sub

Private Sub picSB_DblClick()

Dim lngOpenType As Long

' -------------------------------------------------------------------
' ... Status Bar: If Project File item Selected try load it into NotePad, if Project Ref. Selected try load its GUID into InputBox for copying.
' -------------------------------------------------------------------
    On Error Resume Next
    
'    If m_CurrentTreeFileName <> vbNullString Then
    If m_CurrentTreeFileName <> vbNullString And m_CurrentTreeGUID = vbNullString Then  ' v5 ... hack job.
                                                                                                ' if missing file send file name to input box.
                                                                                                ' if tree file name and guid have values then indicates missing file.
        If Dir$(m_CurrentTreeFileName, vbNormal) <> vbNullString Then
            If mControl = False Then ' ... is Control key pressed, No = False, Yes = True.
                lngOpenType = 11
            Else
'                If Dir$(cAppVB6IDE, vbNormal) <> vbNullString Then
'                    ' ... a few switch options against the VB6.exe, note hard coded path to my vb6.exe only, change as required.
'                    ' ... comment / uncomment to suit, adding the /sdi switch is for single document interface.
'                    ' lngRet = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe /run " & stmpFileName, vbNormalFocus) ' ... for vbp, run vb6.
'                    ' lngRet = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe /runexit" & stmpFileName, vbNormalFocus) ' ... for vbp, run vb6, exit vb6.
'                    ' lngRet = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe /make " & stmpFileName, vbNormalFocus) ' ... for standard exe vbp, make exe.
'                    ' lngRet = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe /makedll " & stmpFileName, vbNormalFocus) ' ... for dll project vbp, make dll
'                    'lngRet = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe " & stmpFileName, vbNormalFocus) ' ... for vb files frm, cls, mod, ctl, vbp (vbg? probably)
'                    lngRet = Shell(cAppVB6IDE & " " & stmpFileName, vbNormalFocus) ' ... for vb files frm, cls, mod, ctl, vbp (vbg? probably)
'                End If
                lngOpenType = 12
            End If
            pOpenSomething lngOpenType, m_CurrentTreeFileName
        End If
    ElseIf m_CurrentTreeGUID <> vbNullString Then
        ' ... what about that folder opening code as well, eg read the ref path and then allow ref folder to be opened when ctrl also pressed.
        ' -------------------------------------------------------------------
        ' v5, make missing file's name open in input box.
        If Len(m_CurrentTreeFileName) = 0 Then
            InputBox "The GUID for the Selected Project Explorer Object:" & vbNewLine & m_CurrentTreeText, "Project Item", m_CurrentTreeGUID
        Else
            InputBox "The Selected Project Explorer File is not available:" & vbNewLine & m_CurrentTreeText, "Project Item", m_CurrentTreeFileName
        End If
    End If
    ' -------------------------------------------------------------------
    ' ... control and shift key pressed hack.
    ' ... failure to capture correct key state after running once.
    Form_KeyDown 0, 0
    ' -------------------------------------------------------------------
End Sub

Private Sub picSB_Paint()
Dim stext As String
' -------------------------------------------------------------------
' ... Status Bar: Write Selected Project item File Name or GUID to the SB Picture Box.
' -------------------------------------------------------------------
    picSB.Cls
    If m_CurrentTreeFileName <> vbNullString Then
        stext = m_CurrentTreeFileName
    ElseIf m_CurrentTreeGUID <> vbNullString Then
        stext = m_CurrentTreeGUID
    End If
    If Len(stext) Then
        picSB.CurrentX = 3 * Screen.TwipsPerPixelX
        picSB.CurrentY = 1 * Screen.TwipsPerPixelY \ 2
        picSB.Print stext
    End If
End Sub

Private Sub tvProj_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
' ... something dropped onto project explorer, process any handled instruction.
Dim sFile As String
Dim sExt As String
'Dim xString As StringWorker
Debug.Print "frmViewer.tvProj_OLEDragDrop.IN.tv proj drop: "

    On Error GoTo ErrHan:   ' ... v6, added error handler
    '                       ' ... If Dir$(sFile, vbNormal) <> "" Then caused error when not a file name.
    If Data.GetFormat(vbCFFiles) Then
        ' ... if it's an ole file...
        sFile = Data.Files(1)
        
    ElseIf Data.GetFormat(vbCFText) Then
        ' ... if it's text, process *.vbp.
        sFile = Data.GetData(1)
        
    End If
    ' -------------------------------------------------------------------
    If Len(sFile) = 0 Then Exit Sub
    ' -------------------------------------------------------------------
    If Dir$(sFile, vbNormal) <> "" Then
        sExt = Right$(LCase$(sFile), 4)
        If sExt = ".vbp" Then
            ' ... if it's a vbp then load it, dismissing anything currently displayed.
            pOpenVBP sFile
        Else
            ' ... commented, for now, only because need to do stuff elsewhere
            ' ... and haven't thought about it.
            ' ... the idea would be to process vbp files individually.
            
'            Select Case sExt
'                Case ".cls", ".bas", ".frm", ".ctl", ".vbs"
'                    ' ... currently handles the following
'                    ' ... a class, form, module, user control, vb script file
'                    ' ... by opening the file into the class explorer.
'                    If Not moVBPInfo Is Nothing Then
'                        Set moVBPInfo = Nothing
'                    End If
'                    If tvProj.Nodes.Count > 0 Then
'                        ' ... clear project explorer.
'                        tvProj.Nodes(1).Selected = True
'                        tvProj.Nodes.Clear
'                    End If
'                    ' ... load the text from the file into the top rt box.
'                    Set xString = New StringWorker
'                    xString.FromFile sFile
'                    rtb(1).Text = xString
'                    ' ... run the process class method to parse the text
'                    ' ... into the class explorer.
'                    pProcessClass , True
'                    Set xString = Nothing
'            End Select
        End If
    End If
    

ResumeError:

Exit Sub

ErrHan:

    Debug.Print "frmViewer.tvProj_OLEDragDrop.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Sub

' -------------------------------------------------------------------
' ... Drag/Drop from Project Tree View to Find Text Box.
Private Sub tvProj_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
Dim stext As String
    On Error GoTo ErrHan:
    If Not tvProj.SelectedItem Is Nothing Then
        Data.Clear
        If tvProj.SelectedItem.Index = 1 Then
            stext = tvProj.SelectedItem.Key
        Else
            If tvProj.SelectedItem.Tag = cFileSig Then ' v5.
                stext = tvProj.SelectedItem.Key
            Else
                stext = tvProj.SelectedItem.Text
            End If
        End If
        Data.SetData stext
        AllowedEffects = 1
    End If

ResumeError:

Exit Sub

ErrHan:

    Debug.Print "frmViewer.tvProj_OLEStartDrag.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub tvMembers_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
' ... start ole drag from class explorer.
Dim stext As String
Dim lngFound As Long
Dim sKey As String
Dim xNode As ComctlLib.Node
'Dim iIndex As Long
'Dim q As QuickMemberInfo
    ' -------------------------------------------------------------------
    mMemberDrag = True ' v8
    ' -------------------------------------------------------------------
    If Not tvMembers.SelectedItem Is Nothing Then
        
        Set xNode = tvMembers.SelectedItem
        
'        If xNode.Tag = cMembSig Then
'            Debug.Print "Begin Drag: " & xNode.Text, xNode.Key
'            iIndex = CLng(Val(xNode.Key))
'            q = moCodeReader.QuickMember(iIndex)
'            Debug.Print q.Declaration
'            stext = moCodeReader.GetMemberCodeLines(q.Index)
'            Debug.Print stext
'            Data.SetData stext
'            GoTo Quit:
''            Exit Sub
'        End If
        
        stext = xNode.Text
        ' ... remove anything after a colon.
        lngFound = InStr(stext, ":")
        If lngFound > 0 Then stext = Left$(stext, lngFound - 1)
        ' ... remove anything after an As ___.
        lngFound = InStr(stext, " As ")
        If lngFound > 0 Then stext = Left$(stext, lngFound - 1)
                
        ' ... single line select case!
        Select Case Asc(Left$(stext, 1)): Case 43, 126, 35: stext = Mid$(stext, 2): End Select
        
        Data.SetData stext
        
        AllowedEffects = 1
        
        If Not xNode.Parent Is Nothing Then
            ' ... was the node an child of the enumerators node?
            If xNode.Parent.Key = cEnusNodeKey Then
                ' -------------------------------------------------------------------
                ' ... generate enum select case.
                ' -------------------------------------------------------------------
                ' ... note: enums key is single string in following format...
                ' ... name : mem1; mem2; mem3, ... etc.
                sKey = xNode.Key
                stext = modGeneral.SelectCaseFromEnum(sKey)
                
                Data.SetData stext
            
            Else
                
                If xNode.Parent.Key = cAPIsNodeKey Then
                    
                    ' ... if it's an api, capture its declaration for copy.
                    stext = xNode.Tag
                    
                    Data.SetData stext
                    
                End If
                
                If xNode.Key = cMembSig Then
                                
                    Debug.Print xNode.Text
                                
                End If
                
            End If
            
        End If
        
    End If
Quit:
    If Len(stext) Then
        Clipboard.Clear
        Clipboard.SetText stext
    End If
    
End Sub
' ... Drop Target: txtFind.
Private Sub txtFind_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim stext As String
    If Len(Data.GetData(1)) Then
        stext = Data.GetData(1)
        txtFind.Text = stext
        ' -------------------------------------------------------------------
'        cmdFind_Click   ' ... what do you think?
        ' -------------------------------------------------------------------
    End If
End Sub

Private Sub pFindMethod(pMethodName As String)
' ... finds a method in the class explorer tree view by name
' ... and forces it to click.
Dim sTmp As String
Dim sFound As String
Dim lngFound As Long
Dim lngMethodIndex As Long
Dim x As StringArray
Dim n As Node
    
    On Error GoTo ErrHan:
    lngMethodIndex = -1
    sTmp = Trim$(pMethodName)
    If Len(sTmp) Then
        If sTmp = tvMembers.Nodes(1).Text Then
            mF2LastNode = tvMembers.Nodes(1).Text
            sTmp = cMainNodeKey
            GoTo Jump:
        End If
        lngFound = InStr(1, sTmp, " ")
        If lngFound > 0 Then
            sTmp = Left$(sTmp, lngFound - 1)
        End If
        
        If Not moCodeReader Is Nothing Then
            If Not moCodeReader.MembersStringArray Is Nothing Then
                Set x = moCodeReader.MembersStringArray
                If x.Count > 0 Then
                    ' ... search the sorted members string array for the item.
                    lngMethodIndex = x.FindClosestItem(sTmp, sFound, True)
                    If lngMethodIndex = -1 Then
                        ' something's wrong with findclosest, backup
                        For lngFound = 1 To x.Count
                            If InStrB(1, x.Item(lngFound), sTmp) > 0 Then
'                                Debug.Print "Method: " & sTmp & " Found... " & x.Item(lngFound)
                                sTmp = Mid$(x.Item(lngFound), Len(sTmp) + 2)
                                lngMethodIndex = CLng(Val(sTmp))
                                Exit For
'                            Else
'                                Debug.Print "Method: " & sTmp & " Not Found... " & x.Item(lngFound)
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End If
    
    If lngMethodIndex > -1 Then
        ' ... derive the tvMembers Key from the info
        ' ... returned in sFound (name : index : .... ) am after index (2nd item in string).
        ' ... and then want to add an x on the end.
        sFound = Mid$(sFound, Len(sTmp) + 2)
        lngFound = InStr(1, sFound, ":")
        lngFound = CLng(Val(sFound)) ' ... v7/8 update to member string info   ' InStr(1, sFound, ":")
        If lngFound >= 0 Then
'            sTmp = Left$(sFound, lngFound - 1) & "x" ' ... v7/8 commented
            sTmp = CStr(lngFound) & "x" ' ... v7/8 update to member string info
            If tvMembers.Nodes.Count > 0 Then
Jump:
                If Not tvMembers.SelectedItem Is Nothing Then
                    mF2LastNode = tvMembers.SelectedItem.Text
                End If
                Set n = tvMembers.Nodes(sTmp)
                If Not n Is Nothing Then
                    ' ... select and click the node if found.
                    tvMembers.Nodes(n.Index).Selected = True
                    tvMembers_NodeClick n
                End If
            End If
        End If
    End If

ResumeError:
    On Error GoTo 0
    Set x = Nothing
    sFound = vbNullString: sTmp = vbNullString
    lngFound = 0: lngMethodIndex = 0
    
Exit Sub

ErrHan:

    Debug.Print "frmViewer.pFindMethod.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:

    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
' -------------------------------------------------------------------
' ... KeyPress: Capture and process relevent function.
' ... notice hack at end compenstaing for poor shift / control key awareness.
' -------------------------------------------------------------------
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyF1 Then
        pShowHelp
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF2 Then
        If Shift Then
            If Len(mF2LastNode) > 0 Then
                pFindMethod mF2LastNode
            End If
        Else
            If Len(rtb(1).Text) > 0 Then
                If rtb(1).SelLength > 0 Then
                    pFindMethod rtb(1).SelText
                End If
            End If
        End If
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF3 Then
        cmdFind_Click
' -------------------------------------------------------------------
' v5, changed key codes, open now f4, refresh (added) = f5
'     also, ctrl+o for open project as in ide.

'    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF5 Then
'        cmdOpenVBP_Click

    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF4 Then
        cmdOpenVBP_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF5 Then
        pRefreshProject ' v5/6
'        If Not moVBPInfo Is Nothing Then
'            If moVBPInfo.Initialised Then
'                pOpenVBP moVBPInfo.FileNameAndPath
'            End If
'        End If
' -------------------------------------------------------------------

    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF6 Then
        cmdViewProj_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF7 Then
        cmdViewMember_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF8 Then
        cmdViewToolbar_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF9 Then
        cmdViewStatus_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF10 Then
        mUseChildWindows = Not mUseChildWindows
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF11 Then
        pHistBack
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF12 Then
        pHistFwd
    ElseIf Shift = 2 Then ' ... Ctrl pressed.
        If KeyCode = VBRUN.KeyCodeConstants.vbKeyP Then
            pPrint
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyI Then
            ' ... project search
            pSearchProject ' v6
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyO Then
            cmdOpenVBP_Click ' v5 amendment
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyS Then
            cmdSaveRTF_Click
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF Then
            If Len(rtb(1).Text) Then
                If rtb(1).SelLength > 0 Then
                    txtFind.Text = rtb(1).SelText
                End If
            End If
            cmdFind_Click
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyR Then
            cmdViewProj_Click
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyW Then
            chkWordWrap.Value = IIf(chkWordWrap.Value = VBRUN.vbChecked, VBRUN.vbUnchecked, VBRUN.vbChecked)
            ' ... v3/4 below, ctrl+b for bold, ctrl+l for line nos.
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyB Then
            chkBoldRTF.Value = IIf(chkBoldRTF.Value = VBRUN.vbChecked, VBRUN.vbUnchecked, VBRUN.vbChecked)
        ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyL Then
            chkLineNos.Value = IIf(chkLineNos.Value = VBRUN.vbChecked, VBRUN.vbUnchecked, VBRUN.vbChecked)
        End If
    End If
    ' -------------------------------------------------------------------
    Form_KeyDown 0, 0   ' ... turns Shift and Control flags False.
    ' -------------------------------------------------------------------
End Sub

Private Sub pLoadTextStrings()
' ... load text values from resource for multiple language support.
    On Error GoTo ErrHan:
    
    tvProj.ToolTipText = LoadResString(101)
    tvMembers.ToolTipText = LoadResString(102)
    imgBtn(0).ToolTipText = LoadResString(103)
    imgBtn(1).ToolTipText = LoadResString(104)
    imgBtn(2).ToolTipText = LoadResString(105)
    imgBtn(3).ToolTipText = LoadResString(106)
    imgBtn(4).ToolTipText = LoadResString(107)
    imgBtn(5).ToolTipText = LoadResString(108)
    imgBtn(6).ToolTipText = LoadResString(109)
    imgBtn(7).ToolTipText = LoadResString(110)
    imgBtn(8).ToolTipText = LoadResString(111)
    imgBtn(10).ToolTipText = LoadResString(112)
    imgBtn(11).ToolTipText = LoadResString(132)
    
    imgBtn(12).ToolTipText = LoadResString(136)
    imgBtn(13).ToolTipText = LoadResString(137)
    
    
    imgBtn(9).ToolTipText = LoadResString(113)
'    lblFind(0).Caption = LoadResString(113) ' ... v8 commented.
    txtFind.ToolTipText = LoadResString(114)
    
    chkFWholeWord.Caption = LoadResString(115)
    chkFWholeWord.ToolTipText = LoadResString(116)
    
    lblFont(1).Caption = LoadResString(119)
    cboFont(0).ToolTipText = LoadResString(120)
    cboFont(1).ToolTipText = LoadResString(121)
    
    chkWordWrap.Caption = LoadResString(122)
    chkWordWrap.ToolTipText = LoadResString(123)
    
    chkColour.Caption = LoadResString(124)
    chkColour.ToolTipText = LoadResString(125)
    
    chkFMatchCase.Caption = LoadResString(117)
    chkFMatchCase.ToolTipText = LoadResString(118)
    
    chkAlign.Caption = LoadResString(126)
    chkAlign.ToolTipText = LoadResString(127)

Exit Sub
ErrHan:
    Debug.Print "frmViewer.pLoadTextStrings.Error: " & Err.Number & "; " & Err.Description
    Err.Clear
    Resume Next

End Sub

Private Sub Form_Load()
Dim lngLoop As Long
Dim lngFCount As Long
Dim sFont As String
Dim sUserFont As String
Dim sFontSize As String
Dim sUserFontSize As String
Dim lngCFontIndex As Long
Dim lResult As Long
Dim xOptions As cOptions

' -------------------------------------------------------------------
' ... Form Load: Set up Font Combos, Initial Sizing and execute design time word wrap option.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    m_loading = True
    
    mRTFColourTable = vbNullString ' v4.
    
    mProjectName = "Code Browser"
    Hide
    
    
    ' v5 ... hand cursor for buttons
    pLoadHandCursor
    
    ' v4, ... options.
    Set xOptions = New cOptions
    xOptions.Read
    sUserFont = xOptions.FontName
    sUserFontSize = xOptions.FontSize
    mShowAttributes = xOptions.ShowAttributes
    mShowClassHeadCount = xOptions.ShowClassHeadCount
    mUseChildWindows = xOptions.UseChildWindows
    
    ' v6 ... syntax colouring
    mWithSyntaxColours = xOptions.AutoRTFEncoding
    mCodeLoadAll = xOptions.AutoLoadAllCode
    ' -------------------------------------------------------------------
    
    mColoured = IIf(chkColour.Value = VBRUN.vbChecked, True, False) ' v3
    pLoadTextStrings
    
    ' ... font sizes followed by names.
    lngCFontIndex = -1
    For lngLoop = 7 To 20
        sFontSize = Format$(lngLoop, "00")
        cboFont(0).AddItem sFontSize
        If sFontSize = sUserFontSize Then
            lngCFontIndex = cboFont(0).ListCount - 1
        End If
    Next lngLoop
    
    If lngCFontIndex > -1 Then
        cboFont(0).ListIndex = lngCFontIndex
    End If
    
    lngCFontIndex = -1
    lngFCount = VB.Screen.FontCount
    
    For lngLoop = 0 To lngFCount - 1
        sFont = VB.Screen.Fonts(lngLoop)
        If sFont = sUserFont Then
            lngCFontIndex = lngLoop
        End If
        cboFont(1).AddItem sFont
    Next lngLoop
    
    If lngCFontIndex > -1 Then
        cboFont(1).Text = sUserFont
    Else
        cboFont(1).Text = "Courier New" ' ... force a click if found courier new (default font) font.
    
    End If
    
    chkBoldRTF.Value = IIf(xOptions.FontBold, vbChecked, vbUnchecked)
    chkLineNos.Value = IIf(xOptions.LineNumbers, vbChecked, vbUnchecked)
    
    If xOptions.UseOwnColours = True Then
        rtb(0).BackColor = xOptions.ViewerBackColour
        rtb(1).BackColor = xOptions.ViewerBackColour
        With xOptions
            mRTFColourTable = modEncode.BuildRTFColourTable(.NormalTextColour, .KeywordTextColour, _
                                                            .CommentTextColour, .AttributeTextColour, _
                                                            .LineNoTextColour)
            picInfo.BackColor = .ViewerBackColour
            picInfo.ForeColor = .AttributeTextColour
            
        End With
    End If
    
    PMenu.ShowTextEditor = xOptions.ShowTextEditor
    PMenu.ShowVB5IDE = xOptions.ShowVB5IDE
    PMenu.ShowVB6IDE = xOptions.ShowVB6IDE
    
    ' ... resize screen controls.
    picCodeCanvas.BorderStyle = 0 ' ... border in designer so can see it's there else not required.
    picSplitMain.Left = 3000 '2800
    picSplitProj.Left = 2500
    
    ' ... ok to resize now.
    m_loading = False
    
    Width = 10000
    Height = 8000
    
    ' -------------------------------------------------------------------
    ' ... if image list is on this form then it is re-created in memory
    ' ... each time a new instance is loaded increasing GDI count.
    Set tvProj.ImageList = mdiMain.ilProject
    Set tvMembers.ImageList = mdiMain.liMember
    ' -------------------------------------------------------------------
    ' ... set up cue banner to find text box.
    Let lResult = SendMessageLongW(txtFind.hwnd, EM_SETCUEBANNER, 0, StrPtr("Enter Search Text"))
    
    picSplitRTB.Top = picInfo.Height + cSplitterHeight
    
    chkWordWrap_Click   ' ... force rtb word wrap property.
    cboFont_Click 0     ' ... force rtb text font size

'''    mdiMain.ChildLoaded ChildIndex, mProjectName ' Caption
    
    ' -------------------------------------------------------------------
    
    ClearMemory
    
    Show
    
Exit Sub
ErrHan:

    Debug.Print "frmViewer.Form_Load.Error: " & Err.Number & "; " & Err.Description
    Err.Clear: Resume Next

End Sub

' -------------------------------------------------------------------
' ... The following section is the resizing source only.
Private Sub pResize()
' ... main control resizing done in here.
Dim xp As CoDim, xs As CoDim
    On Error Resume Next
    xs.Height = ScaleHeight: xs.Width = ScaleWidth
    If picTB.Visible Then ' ... picTB > Toolbar.
        xs.Height = xs.Height - picTB.Height
        xs.Top = picTB.Height
    End If
    If picSB.Visible Then xs.Height = xs.Height - picSB.Height ' ... picSB > Status Bar.
    xp.Top = xs.Top
    xp.Height = xs.Height: xp.Width = xs.Width
    If tvProj.Visible Then ' ... tvProject > Project Explorer.
        xp.Left = picSplitProj.Left
        xp.Width = cSplitterWidth
        If mSizeEditorOnly = False Then
    '        Debug.Print "frmViewer.pResize.IN.resize ever thing: "
            picSplitProj.Move xp.Left, xp.Top, xp.Width, xp.Height
            imgSplitProj.Move xp.Left, xp.Top, xp.Width, xp.Height
            tvProj.Move 0, xp.Top, xp.Left, xp.Height
        End If
        xp.Width = xs.Width - picSplitProj.Left - cSplitterWidth
        xp.Left = xp.Left + cHBorder
    End If
    If mSizeEditorOnly = False Then picMain.Move xp.Left, xp.Top, xp.Width, xp.Height ' ... picMain > Canvas to Class Explorer and Code Viewer.
    imgSplitProj.Visible = tvProj.Visible
    xp.Top = 0: xp.Left = 0
    xp.Height = picMain.Height: xp.Width = picMain.Width
    If tvMembers.Visible Then ' ... tvMembers > Class Explorer.
        xp.Left = picSplitMain.Left
        xp.Width = xp.Left
        If mSizeEditorOnly = False Then
            picSplitMain.Move xp.Left, xp.Top, cSplitterWidth, xp.Height
            imgSplitMain.Move xp.Left, xp.Top, cSplitterWidth, xp.Height
            tvMembers.Move 0, xp.Top, xp.Width, xp.Height
        End If
        xp.Left = xp.Left + cSplitterWidth
        xp.Width = picMain.Width - xp.Left
    End If
    If mSizeEditorOnly = False Then picCodeCanvas.Move xp.Left, xp.Top, xp.Width, xp.Height ' ... picCodeCanvas > Canvas to Code Viewers.
    imgSplitMain.Visible = tvMembers.Visible
    xp.Left = 0: xp.Top = 0
    picInfo.Move xp.Left, xp.Top, xp.Width
    
    xp.Top = picSplitRTB.Top
    picSplitRTB.Move xp.Left, xp.Top, xp.Width, (2 * cSplitterHeight)
    imgSplitRTB.Move xp.Left, xp.Top, xp.Width, (2 * cSplitterHeight)
    xp.Height = xp.Top - picInfo.Height
    xp.Top = picInfo.Height
    rtb(0).Move xp.Left, xp.Top, xp.Width, xp.Height
    xp.Top = picSplitRTB.Top + (2 * cSplitterHeight)
    xp.Height = picMain.Height - xp.Top
    rtb(1).Move xp.Left, xp.Top, xp.Width, xp.Height
    
End Sub
Private Sub cmdViewProj_Click()
    tvProj.Visible = Not tvProj.Visible
    pResize
End Sub
Private Sub cmdViewMember_Click()
    tvMembers.Visible = Not tvMembers.Visible
    pResize
End Sub
Private Sub cmdViewToolbar_Click()
    picTB.Visible = Not picTB.Visible
    pResize
End Sub
Private Sub cmdViewStatus_Click()
    picSB.Visible = Not picSB.Visible
    pResize
End Sub
Private Sub Form_Resize()
    If m_loading Then Exit Sub
    pResize
End Sub
Private Sub imgSplitMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLMouseDown = True
    picSplitMain.Visible = True
End Sub
Private Sub imgSplitMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' ... class / viewer resizer is moveing.
Dim xcd As CoDim, xmm As MinMax
    If mLMouseDown Then
        xcd.Left = imgSplitMain.Left + x
        xmm.Min = cSplitLimit
        xmm.max = picMain.Width - (2 * cSplitLimit)
        If xcd.Left > xmm.Min And xcd.Left < xmm.max Then
            picSplitMain.Move xcd.Left
            pResize
        End If
    End If
End Sub
Private Sub imgSplitMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLMouseDown = False
    picSplitMain.Visible = False
    pResize
End Sub
Private Sub imgSplitProj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLMouseDown = True
    picSplitProj.Visible = True
End Sub
Private Sub imgSplitProj_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' ... project resizer is moving.
Dim xcd As CoDim, xmm As MinMax
    If mLMouseDown Then
        xcd.Left = imgSplitProj.Left + x
        xmm.Min = cSplitLimit
        xmm.max = ScaleWidth - (3 * cSplitLimit)
        If xcd.Left > xmm.Min And xcd.Left < xmm.max Then
            picSplitProj.Move xcd.Left
            pResize
        End If
    End If
End Sub
Private Sub imgSplitProj_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLMouseDown = False
    picSplitProj.Visible = False
    pResize
End Sub
Private Sub imgSplitRTB_DblClick()
' ... double click on the viewer splitter bar, hide or show accordingly.
    mSizeEditorOnly = True
    If imgSplitRTB.Top > 100 * Screen.TwipsPerPixelY Then
        picSplitRTB.Top = picInfo.Height + 1 * Screen.TwipsPerPixelY
    Else
        picSplitRTB.Top = picCodeCanvas.Height \ 2
    End If
End Sub
Private Sub imgSplitRTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Debug.Print "frmViewer.imgSplitRTB_MouseDown.IN. : "
    mSizeEditorOnly = True
    mLMouseDown = True
    picSplitRTB.Visible = True
End Sub
Private Sub imgSplitRTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' ... viewer's horizontal splitter bar is moving.
Dim xcd As CoDim, xmm As MinMax
    If mLMouseDown Then
        xcd.Top = imgSplitRTB.Top + y
        xmm.Min = picInfo.Height
        xmm.max = picMain.Height - (45 * Screen.TwipsPerPixelY)
        If xcd.Top > xmm.Min And xcd.Top < xmm.max Then
            picSplitRTB.Top = xcd.Top
            pResize
        End If
    End If
End Sub
Private Sub imgSplitRTB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLMouseDown = False
    picSplitRTB.Visible = False
    pResize
    mSizeEditorOnly = False
End Sub
' ... End of resizing source.
' -------------------------------------------------------------------

