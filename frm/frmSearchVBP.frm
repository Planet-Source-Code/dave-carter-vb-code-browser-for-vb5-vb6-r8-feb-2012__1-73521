VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSearchVBProject 
   Caption         =   "Project Search"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   43
   Icon            =   "frmSearchVBP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   10260
   Begin VB.PictureBox picSplitProj 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   2610
      ScaleHeight     =   1245
      ScaleWidth      =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   10200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4920
      Width           =   10260
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   210
         Left            =   60
         TabIndex        =   19
         Top             =   30
         Width           =   60
      End
   End
   Begin VB.PictureBox picTB 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10260
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "465"
      Top             =   0
      Width           =   10260
      Begin VB.CheckBox chkFWholeWord 
         Caption         =   "Whole Word"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5220
         TabIndex        =   2
         Top             =   90
         Width           =   1245
      End
      Begin VB.CheckBox chkFMatchCase 
         Caption         =   "Text Compare"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   6720
         TabIndex        =   3
         Top             =   90
         Width           =   1425
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         ToolTipText     =   "Enter Search String"
         Top             =   90
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(r & d draft)"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   8280
         TabIndex        =   4
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   765
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   9
         Left            =   2280
         Picture         =   "frmSearchVBP.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         Picture         =   "frmSearchVBP.frx":0B14
         Stretch         =   -1  'True
         Top             =   90
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   8520
         Picture         =   "frmSearchVBP.frx":109E
         Stretch         =   -1  'True
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   1
         Left            =   690
         Picture         =   "frmSearchVBP.frx":1628
         Stretch         =   -1  'True
         Top             =   90
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   8220
         Picture         =   "frmSearchVBP.frx":1BB2
         Stretch         =   -1  'True
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   3
         Left            =   1020
         Picture         =   "frmSearchVBP.frx":213C
         Stretch         =   -1  'True
         Top             =   90
         Width           =   285
      End
      Begin VB.Image imgBtn 
         Height          =   285
         Index           =   4
         Left            =   1350
         Picture         =   "frmSearchVBP.frx":26C6
         Stretch         =   -1  'True
         Top             =   90
         Width           =   285
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   3600
      ScaleHeight     =   3390
      ScaleWidth      =   5175
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5175
      Begin VB.PictureBox picSplitMain 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   2220
         ScaleHeight     =   1455
         ScaleWidth      =   75
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.PictureBox picCodeCanvas 
         BorderStyle     =   0  'None
         Height          =   3105
         Left            =   2460
         ScaleHeight     =   3105
         ScaleWidth      =   2535
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   2535
         Begin VB.PictureBox picSplitRTB 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Height          =   90
            Left            =   900
            ScaleHeight     =   90
            ScaleWidth      =   705
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2700
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   60
            ScaleHeight     =   840
            ScaleWidth      =   2265
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   120
            Width           =   2325
            Begin VB.Label lblMember 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
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
               Height          =   240
               Left            =   60
               TabIndex        =   13
               Top             =   480
               Width           =   75
            End
            Begin VB.Label lblFind 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Search Result"
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
               Height          =   210
               Left            =   60
               TabIndex        =   12
               Top             =   30
               Width           =   10890
               WordWrap        =   -1  'True
            End
         End
         Begin RichTextLib.RichTextBox rtb 
            Height          =   555
            Left            =   120
            TabIndex        =   17
            Top             =   2220
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   979
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            TextRTF         =   $"frmSearchVBP.frx":2C50
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
         Begin ComctlLib.ListView lv 
            Height          =   615
            Left            =   1860
            TabIndex        =   14
            Top             =   2220
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Source"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Index"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Line No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Line Pos."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Text Pos."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Line Text"
               Object.Width           =   21167
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   "LineSortVal"
               Text            =   "LineSortVal"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "KeyField"
               Object.Width           =   0
            EndProperty
         End
         Begin ComctlLib.ListView lvFilter 
            Height          =   615
            Left            =   1860
            TabIndex        =   15
            Top             =   1380
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Source"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Index"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Line No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Line Pos."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Text Pos."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Line Text"
               Object.Width           =   21167
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   "LineSortVal"
               Text            =   "LineSortVal"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "KeyField"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Image imgSplitRTB 
            Height          =   165
            Left            =   900
            MousePointer    =   7  'Size N S
            ToolTipText     =   "Resize Me"
            Top             =   2400
            Width           =   645
         End
      End
      Begin VB.PictureBox picMembers 
         BackColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   90
         ScaleHeight     =   2955
         ScaleWidth      =   1185
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Image imgSplitMain 
         Height          =   1425
         Left            =   1590
         MousePointer    =   9  'Size W E
         ToolTipText     =   "Resize Me"
         Top             =   180
         Width           =   105
      End
   End
   Begin ComctlLib.TreeView tvProj 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6165
      _Version        =   327682
      Indentation     =   176
      Style           =   7
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
   End
   Begin VB.Image imgSplitProj 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   2820
      MousePointer    =   9  'Size W E
      ToolTipText     =   "Resize Me"
      Top             =   1110
      Width           =   165
   End
End
Attribute VB_Name = "frmSearchVBProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface to searching a project for a string."

' what?
'  a user interface for searching a (classic) vb project
'  for a substring.
' why?
'  just because, it's the kind of thing you might like to do.
' when?
'  you want to search for something(?)
' how?
'  select a project, enter a query string and click find button.
'  press escape to cancel loading the search results.
'  click a node to filter the results to either a group
'  such as Forms or Classes or a source file or the top node
'  to show all results.
' who?
'  d.c.

' Notes:
'       Search results are naturally filtered to one per line
'       so if the substring is on a line more than once the results
'       won't show it.
'
'       Heap big improvements to be made to the efficiency of code... still researching,
'       busker's 2nd attempt.

Option Explicit

Private Const cDlgCancelErr As Long = 32755

Private mBusy As Boolean
Private mCancel As Boolean

Private mFileNameInfo As FileNameInfo
Private moVBPInfo As VBPInfo

Private mLoading As Boolean

Private mLMouseDown As Boolean
Private mSizeEditorOnly As Boolean
Private mShift As Boolean
Private mControl As Boolean

Private Const cHBorder As Long = 60
Private Const cSplitLimit As Long = 1200 '900 '660
Private Const cSplitterHeight As Long = 30
Private Const cSplitterWidth As Long = 45 '60

Private Const MF_CHECKED = &H8&
Private Const TPM_LEFTALIGN = &H0&
Private Const MF_STRING = &H0&
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_RIGHTBUTTON = &H2&

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' -------------------------------------------------------------------
' ... this lot is for pParseClass
' ... bit of a repetition.

Private Const cm_len_Sub As Long = 4
Private Const cm_len_Function As Long = 9
Private Const cm_len_Property As Long = 9
 
Private Const cm_word_Sub As String = "Sub "
Private Const cm_word_Function As String = "Function "
Private Const cm_word_Property As String = "Property "

Private Const cm_len_Public As Long = 7
Private Const cm_len_Private As Long = 8
Private Const cm_len_Friend As Long = 7

Private Const cm_word_Public As String = "Public "
Private Const cm_word_Private As String = "Private "
Private Const cm_word_Friend As String = "Friend "
' -------------------------------------------------------------------
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub LoadVBP(pTheVBPFileName As String, Optional pSearchString As String = vbNullString)
    psOpen pTheVBPFileName
    If Len(pSearchString) Then
        txtFind.Text = pSearchString
        cmdFind_Click
    End If
End Sub

Private Sub psOpen(Optional ByVal pTheFileName As String = vbNullString)

' ... Open the VBP and load project explorer tree view.

Dim sFileName As String
Dim sCDFilter As String
Dim oVBPTree As VBProjTree

    On Error GoTo ErrHan:
    
    If Len(pTheFileName) Then
        sFileName = pTheFileName
    Else
        sCDFilter = modDialog.MakeDialogFilter("Visual Basic Project", , "vbp")
        sFileName = modDialog.GetOpenFileName(, , sCDFilter)
    End If
    
    If Len(sFileName) Then
        
        pInit
        
        mBusy = True
        
        modFileName.ParseFileNameEx sFileName, mFileNameInfo
        
        With mFileNameInfo
        
            If LCase$(.Extension) = "vbp" Then
                
                Set moVBPInfo = New VBPInfo
                moVBPInfo.ReadVBP .PathAndName
                
                Set tvProj.ImageList = mdiMain.ilProject
                Set oVBPTree = New VBProjTree
                
                oVBPTree.Init moVBPInfo, tvProj, , eSourceFiles ' ... filter nodes to source files only.
                
                Caption = "Search: " & moVBPInfo.Title
                
            End If
            
        End With
    
    End If
        
ResErr:
    mBusy = False
    
Exit Sub
ErrHan:
    If Err.Number <> cDlgCancelErr Then
        Debug.Print "frmSearch.psOpen.Error: ", Err.Number, Err.Description
    End If
    Resume ResErr:
    
End Sub


Private Sub cmdFind_Click()

Dim oVBPSearch As VBPSearch

Dim oFSA As StringArray
Dim sTmpArray As StringArray

Dim lngFoundPositions() As Long
Dim lngFileLines() As Long
Dim lngLines() As Long
Dim lngLens() As Long
Dim sMembers() As String
Dim lMembers() As Long

Dim c As Long
Dim k As Long
Dim p As Long
Dim lIndex As Long
Dim lFound As Long
Dim lngLoop As Long
Dim lngLength As Long
Dim lngDelLen As Long
Dim lngLastLine As Long
Dim lngLineCount As Long
Dim lngFilesSearched As Long
Dim lngFilesFound As Long
Dim x As Long
Dim y As Long
Dim a As Long
Dim b As Long

Dim sFileName As String
Dim sFName As String
Dim sFileText As String
Dim sTheString As String
Dim sSearch As String
Dim sDel As String
Dim sTag As String
Dim s As String
Dim zz As String

Dim xTime As zTimer

Dim fFileInfo As FileNameInfo

Dim lngComp As VbCompareMethod

Dim xItem As ListItem

Dim xNode As Node
Dim oFCount As LongNumber
Dim oCCount As LongNumber
Dim oMCount As LongNumber
Dim oUCount As LongNumber

    ' ... Note:
    ' ... Bit of a mess with commented lines, changed from adding a key to list view items
    ' ... to having no key because performance is significantly improved
    ' ... diddled with the idea of only allowing 32000 items in list view (integer bound)
    ' ... but commented this.
    ' ... There's a bug in deriving line lengths or something so there's a hack as well.

' WARNING:
'   This code may mostly work but i made it up as i went along and didn't get back
'   to writing it properly, its way over the top.
'   I can rewrite this with the new VBProject Class, or
'   I can keep hold of files read rather than reading them twice with each search

'   The VBProject Class reads and parses all project source code files
'   before it is ready to do anything whereas, currently, this method
'   parses source code files as required (it does this to derive the
'   name of the method to which the result belongs).
'   I want to have another go at this
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    ' ... obvious thing to start with, will retain last query results exiting here.
    sSearch = txtFind.Text
    If Len(Trim$(sSearch)) = 0 Then
        MsgBox "Please provide a sub string to query.", vbInformation, Caption
        Exit Sub
    End If
    ' -------------------------------------------------------------------
    Screen.MousePointer = vbHourglass
    ' -------------------------------------------------------------------
    
    mBusy = True
    mCancel = False
    
    ' ... general tidy up of screen for new search.
    lv.ListItems.Clear
    lvFilter.ListItems.Clear
    
    lv.Sorted = False
    lvFilter.Sorted = False
    
    rtb.Text = ""
    rtb.LoadFile ""
    
'    lv.Visible = False
    lv.ZOrder
    lblFind.Caption = ""
    lblMember.Caption = ""
    ' -------------------------------------------------------------------
    ' ... clear project node result numbering.
    If tvProj.Nodes.Count > 0 Then
        tvProj.Nodes(1).Text = moVBPInfo.ProjectName
        For Each xNode In tvProj.Nodes
            If InStr(1, xNode.Text, ";") > 0 Then
                xNode.Text = Left$(xNode.Text, InStr(1, xNode.Text, ";") - 1)
            End If
        Next xNode
    End If
    
    Set oFCount = New LongNumber
    Set oCCount = New LongNumber
    Set oMCount = New LongNumber
    Set oUCount = New LongNumber
    
    DoEvents ' ... allow list views to catch up
    ' -------------------------------------------------------------------
    
    Set oVBPSearch = New VBPSearch
    oVBPSearch.Init moVBPInfo
    
    lngComp = IIf(chkFMatchCase.Value And vbChecked, VbCompareMethod.vbTextCompare, VbCompareMethod.vbBinaryCompare)
    
    lblFind.Caption = "Searching for ... " & sSearch
    lblFind.Refresh
    
    lngFilesSearched = moVBPInfo.FilesData.Count
    
    Set xTime = New zTimer
    xTime.StartTiming False
    ' ... initial/primary search to provide count.
    lFound = oVBPSearch.SearchFiles(sSearch, oFSA, , lngComp, IIf(chkFWholeWord.Value And vbChecked, True, False))
    xTime.StopTiming
    
    lblFind.Caption = "Found word [ " & sSearch & " ]: " & Format$(lFound, cNumFormat)
    lblFind.Refresh
    
    If Not oFSA Is Nothing Then
    
        ' ... oFSA is a StringArray of files with a matching substring returned from initial search above.
        If oFSA.Count > 0 Then
        
            lngFilesFound = oFSA.Count
        
            ' ... secondary search, this time grabbing details from files listed in oFSA.
            c = 0
                        
            For lngLoop = 1 To oFSA.Count
                
                Set sTmpArray = oFSA.ItemAsStringArray(lngLoop, "|")
                sFileName = sTmpArray(4)
                modFileName.ParseFileNameEx sFileName, fFileInfo
                sFName = fFileInfo.FileName
                
                sTag = UCase$(fFileInfo.Extension)
                ' ... read the file & get its Text, File Lines Array and Line Deliiter.
                modReader.ReadTextFile sFileName, sFileText, lngFileLines, sDel
                lngDelLen = Len(sDel)
                
                ' -------------------------------------------------------------------
                ' ... read the source code member names and positions.
'                If lngLineCount < 32000 Then
                    ' ... skip finding the member names because we aren't adding any more to the list view.
                    pParseClass sFileText, sDel, sMembers, lMembers
'                    DoEvents
'                End If
                ' -------------------------------------------------------------------
                
                ' ... search the file & get an Array of SubString Start Character Positions.
                modStrings.FindAllMatches sFileText, sSearch, lngFoundPositions, , lngComp, IIf(chkFWholeWord.Value And vbChecked, True, False)
                ' ... derive line numbers from search results; using Line Delimiter Length, File Lines and Char Pos Arrays, derive line nos. for substring finds.
                modReader.DeriveLineNumbers lngFileLines, lngFoundPositions, lngLines, lngDelLen 'Len(sDel)
                ' ... derive line lengths from search results and line numbers.
                modReader.DeriveLineLengths lngFileLines, lngLines, lngLens, lngDelLen 'Len(sDel)
                
                On Error Resume Next
                
                a = 0: b = 0
                lngLastLine = 0
                y = 0
                
                For k = 0 To UBound(lngLines)
                    
                    ' ... NOTE:
                    ' ... at present, line numbering begins in the source file's header section not the start of its declarative section.
                    ' ... extract lines where the substring is found.
                    c = c + 1
                    
                    lIndex = lngLines(k)
                    If lIndex < 1 And k > 0 Then
                        lIndex = UBound(lngFileLines) - 1
                    Else
                        p = lngFoundPositions(k) - lngFileLines(lIndex)
                    End If
                    
                    If p < 1 Then p = 1
                    
                    If lIndex <> lngLastLine Or lngLastLine = 0 Then
                    
                        lngLineCount = lngLineCount + 1
                        
'                        If lngLineCount < 32000 Then
                        
                            lngLastLine = lIndex
                            lngLength = lngLens(k)
                            
    ' -------------------------------------------------------------------
                            ' ... extract current line from file text.
                            ' ... there's a bug calculating line lengths (or something)
                            ' ... when substring is the first thing in the file e.g. VERSION
                            ' ... we read an extra character in the line.
                            
                            s = Mid$(sFileText, lngFileLines(lIndex), lngLength)
                            
                            If p = 1 Then
                                ' ... hack bug.
                                If Right$(s, 1) = vbCr Then
                                    s = Left$(s, Len(s) - 1)
                                End If
                            End If
    ' -------------------------------------------------------------------
                            zz = sMembers(UBound(sMembers)) ' ... default to last because loop never gets there.
                            ' ... try finding the name of the method or member.
                            For x = y To UBound(lMembers) - 1
                                
                                a = lMembers(x)
                                b = lMembers(x + 1) - 1
                                
                                If lIndex >= a And lIndex < b Then
                                    zz = sMembers(y)
                                    Exit For
                                ElseIf lIndex > b Then
                                    y = y + 1
                                End If
                                
                            Next x
                            
                            If Len(s) Then
                                
'                                Set xItem = lv.ListItems.Add(, sFileName & "|" & CStr(lIndex) & "|" & CStr(lngFoundPositions(k)), sFName & ": " & zz)         ' ... source name.
                                ' ... try adding without a key, see if quicker
                                ' ... key less means lv doesn't have to test key exists
                                ' ... key less is so much faster than keyed
                                Set xItem = lv.ListItems.Add(, , sFName & ": " & zz)        ' ... source name.

                                xItem.SubItems(1) = Format$(c, cNumFormat)         ' ... search index.
                                xItem.SubItems(2) = Format$(lIndex, cNumFormat)    ' ... line no.
                                xItem.SubItems(3) = Format$(p, cNumFormat)         ' ... line pos.
                                xItem.SubItems(4) = Format$(lngFoundPositions(k), cNumFormat)        ' ... text pos.
                                xItem.SubItems(5) = s
                                xItem.SubItems(6) = Format$(lIndex, "000000000")   ' ... line no.
                                xItem.SubItems(7) = sFileName & "|" & CStr(lIndex) & "|" & CStr(lngFoundPositions(k)) ' ... key field.
                                
                                xItem.Tag = sTag
                                
                                s = vbNullString
'                                zz = vbNullString
                                
                            End If
                                                
'                        Else
'                            Debug.Print "Nearing Max List View Items"
'                        End If
                        
                    End If
                                                            
                Next k
                
'                If lngLineCount > 31999 Then
'                    lblMember.Caption = "Too many items to display, press Esc to cancel or allow to complete to see unique line count."
'                End If
                
                Select Case sTag
                    Case "FRM": oFCount.Increment k
                    Case "CLS": oCCount.Increment k
                    Case "BAS": oMCount.Increment k
                    Case "CTL": oUCount.Increment k
                End Select
                
                sFileText = vbNullString ' ... release the File's Text resource.
                
                If tvProj.Nodes.Count Then
                    ' -------------------------------------------------------------------
                    ' ... update the source file's node with no. of items found.
                    Set xNode = tvProj.Nodes(fFileInfo.PathAndName)
                    If Not xNode Is Nothing Then
                        xNode.Text = xNode.Text & "; " & Format$(k, cNumFormat)
                    End If
                                        
                End If
                
                If lFound > 0 Then
                    lblFind.Caption = "Found word [ " & sSearch & " ]: Item Count = " & Format$(lFound, cNumFormat) & ": Unique Line Count = " & Format$(lngLineCount, cNumFormat) & " >>> Loading... ."
                End If

                lblFind.Caption = lblFind.Caption & vbNewLine & "Files Searched = " & Format$(lngLoop, cNumFormat) & " of " & Format$(lngFilesSearched, cNumFormat) & ": Search String found in " & Format$(lngFilesFound, cNumFormat) & " out of " & Format$(lngFilesSearched, cNumFormat)
                lblFind.Refresh
                
                If mCancel Then GoTo ResumeError:
                
                DoEvents
                Sleep 5
                
            Next lngLoop
            
        End If
        
    End If
    
ResumeError:
    
    On Error GoTo 0
    
    If tvProj.Nodes.Count Then
        ' -------------------------------------------------------------------
        ' ... update the tree view nodes with numbers of items found ...
        
        Set xNode = tvProj.Nodes(1) ' ... Top Node: Total found / unique lines.
        If Not xNode Is Nothing Then
            xNode.Text = moVBPInfo.ProjectName & "; " & Format$(lFound, cNumFormat) & " / " & Format$(lngLineCount, cNumFormat)
        End If
        If oFCount > 0 Then         ' ... Forms Node: total found in forms.
            Set xNode = tvProj.Nodes(cFormNodeKey)
            If Not xNode Is Nothing Then
                xNode.Text = RTrim$(cFormNodeKey) & "; " & Format$(oFCount, cNumFormat)
            End If
        End If
        If oCCount > 0 Then         ' ... Classes Node: total found in classes.
            Set xNode = tvProj.Nodes(cClasNodeKey)
            If Not xNode Is Nothing Then
                xNode.Text = RTrim$(cClasNodeKey) & "; " & Format$(oCCount, cNumFormat)
            End If
        End If
        If oMCount > 0 Then         ' ... Modules Node: total found in modules.
            Set xNode = tvProj.Nodes(cModsNodeKey)
            If Not xNode Is Nothing Then
                xNode.Text = RTrim$(cModsNodeKey) & "; " & Format$(oMCount, cNumFormat)
            End If
        End If
        If oUCount > 0 Then         ' ... User Controls Node: total found in user controls.
            Set xNode = tvProj.Nodes(cUCtlNodeKey)
            If Not xNode Is Nothing Then
                xNode.Text = RTrim$(cUCtlNodeKey) & "; " & Format$(oUCount, cNumFormat)
            End If
        End If
    End If


    If lFound > 0 Then
        lblFind.Caption = "Found word [ " & sSearch & " ]: Item Count = " & Format$(lFound, cNumFormat) & ": Unique Line Count = " & Format$(lngLineCount, cNumFormat) & IIf(mCancel, " >> Incomplete, Cancelled <<", "")
    End If
    
    lblFind.Caption = lblFind.Caption & vbNewLine & "Files Searched = " & Format$(lngFilesSearched, cNumFormat) & ": Search String found in " & Format$(lngFilesFound, cNumFormat) & " out of " & Format$(lngFilesSearched, cNumFormat)
    If Not xTime Is Nothing Then
        lblFind.Caption = lblFind.Caption & ". Initial Search Time: " & xTime.Duration & " secs"
    End If
    lblFind.Refresh
    
    If lv.ListItems.Count Then
        modGeneral.LVSizeColumn lv, 0
        lvFilter.ColumnHeaders(1).Width = lv.ColumnHeaders(1).Width
        lv.ListItems(1).Selected = True
        lv_ItemClick lv.ListItems(1)
    End If
    
    lv.Visible = True
    lv.ZOrder
    
    ' -------------------------------------------------------------------
    ' ... general release resources.
    ' ... I prefer to do this than leave it upto vb to dispose of garbage as such.
    ' ... I have noticed some improvement in performance for doing so.
    
    If Not oFCount Is Nothing Then
        Set oFCount = Nothing
    End If
    
    If Not oCCount Is Nothing Then
        Set oCCount = Nothing
    End If

    If Not oMCount Is Nothing Then
        Set oMCount = Nothing
    End If

    If Not oUCount Is Nothing Then
        Set oUCount = Nothing
    End If
    
    If Not oFSA Is Nothing Then
        Set oFSA = Nothing
    End If
    
    If Not sTmpArray Is Nothing Then
        Set sTmpArray = Nothing
    End If
    
    If Not oVBPSearch Is Nothing Then
        Set oVBPSearch = Nothing
    End If
    
    If Not xNode Is Nothing Then
        Set xNode = Nothing
    End If
    
    If Not xTime Is Nothing Then
        Set xTime = Nothing
    End If
    
    Erase lngFileLines
    Erase lngFoundPositions
    Erase lngLines
    Erase lngLens
    Erase sMembers
    Erase lMembers
    
    sFileName = vbNullString
    sFName = vbNullString
    sFileText = vbNullString
    sTheString = vbNullString
    sSearch = vbNullString
    sDel = vbNullString
    sTag = vbNullString
    s = vbNullString
    zz = vbNullString
    
    picSplitMain.ZOrder
    
    mBusy = False
    
    Screen.MousePointer = vbDefault
    
    DoEvents

Exit Sub

ErrHan:

    Debug.Print "frmSearchVBP.cmdFind_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub cmdOpenFile_Click()
    psOpen
End Sub

Private Sub Form_Load()

    mLoading = True
    
    ' default scale mode for resizing to stand a chance of working.
    ScaleMode = vbTwips
    
    pInit
    
    ' force far right top object on top.
    picSplitRTB.Top = 3 * picInfo.Height
    pResize
            
    ' ... set up FullRowSelect on the List Views.
    modGeneral.LVFullRowSelect lv.hwnd
    modGeneral.LVFullRowSelect lvFilter.hwnd
    ' ... turn off rich text box auto word wrapping.
    modGeneral.WordWrapRTFBox rtb.hwnd
    
    lv.ZOrder
    picSplitMain.ZOrder
    
    mLoading = False
    
    ClearMemory

End Sub

Private Function pShowPopUp() As Long

Dim Pt As POINTAPI
Dim ret As Long
Dim lFlags As Long
Dim hMenu As Long

Const cFlags As Long = MF_STRING

    hMenu = CreatePopupMenu()
    
    lFlags = cFlags + IIf(tvProj.Visible, MF_CHECKED, 0)
    AppendMenu hMenu, lFlags, 1, "Project Explorer"
    
    lFlags = cFlags + IIf(picTB.Visible, MF_CHECKED, 0)
    AppendMenu hMenu, lFlags, 3, "Toolbar"
    
    lFlags = cFlags + IIf(picSB.Visible, MF_CHECKED, 0)
    AppendMenu hMenu, lFlags, 4, "Status Bar"
    
    GetCursorPos Pt
    
    ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.hwnd, ByVal 0&)
    
    DestroyMenu hMenu
    
    pShowPopUp = ret
    
    ret = 0&
    lFlags = 0&
    hMenu = 0&
    
End Function

Private Sub pDoPopUp()
Dim lRet As Long
Dim iBtn As Integer

    lRet = pShowPopUp
    
    If lRet > 0 Then
        iBtn = CInt(lRet)
        pButtonClick iBtn
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mBusy And UnloadMode = 0 Then
        Cancel = True
    End If
End Sub

Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picInfo_MouseUp Button, 0, 0, 0
End Sub

Private Sub lblMember_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picInfo_MouseUp Button, 0, 0, 0
End Sub

Private Sub picInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    pDoPopUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pRelease
    ClearMemory
End Sub

Private Sub pInit()
' -------------------------------------------------------------------
' Helper:   Initialise stuff for this form.
' -------------------------------------------------------------------
    pRelease
    ' ... .
    ' ... .
End Sub

Private Sub pRelease()
' -------------------------------------------------------------------
' Helper:   Release resources used by this form.
' -------------------------------------------------------------------

    On Error Resume Next
    
    ' -------------------------------------------------------------------
    ' ... clear list view data.
    ' ... I'm setting the first item as selected before calling Clear
    ' ... I think this comes from using the VB6 ListView which can
    ' ... go silly when cleared and setting the first item as selected
    ' ... seemed to stop this observed behaviour.
    If lv.ListItems.Count Then
        lv.ListItems(1).Selected = True
        lv.ListItems.Clear
    End If
    If lvFilter.ListItems.Count Then
        lvFilter.ListItems(1).Selected = True
        lvFilter.ListItems.Clear
    End If
    ' -------------------------------------------------------------------
    If Not moVBPInfo Is Nothing Then
        Set moVBPInfo = Nothing
    End If
    ' -------------------------------------------------------------------
    rtb.FileName = ""   ' ... a quick way to erase the text from an rt box?
    mCancel = False
    mBusy = False
    ClearMemory
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' -------------------------------------------------------------------
'   Cheap way to tell if Control or Shift keys are pressed when doing some process.
'   Tried using GetASyncKeyState API but not successfully; it needed to catch up as it were.
'   Not entirely satisfactory solution; if app lostfocus goes before keyup then keyup not processed.
' -------------------------------------------------------------------
    mControl = Shift And 2
    mShift = Shift And 1
' -------------------------------------------------------------------
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
' -------------------------------------------------------------------
' ... KeyPress: Capture and process relevent function.
' ... notice hack at end compenstaing for poor shift / control key awareness.
' -------------------------------------------------------------------
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        mCancel = True
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF3 Then
        cmdFind_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF4 Then
        cmdOpenFile_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF6 Then
        cmdViewProj_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF8 Then
        cmdViewToolbar_Click
    ElseIf KeyCode = VBRUN.KeyCodeConstants.vbKeyF9 Then
        cmdViewStatus_Click
    ElseIf KeyCode = 93 Then
        pDoPopUp
    ElseIf Shift = 2 Then
        If KeyCode = VBRUN.KeyCodeConstants.vbKeyR Then
            cmdViewProj_Click
        End If
    End If
    ' -------------------------------------------------------------------
    Form_KeyDown 0, 0   ' ... turns Shift and Control flags False.
    ' -------------------------------------------------------------------
End Sub

' -------------------------------------------------------------------
' ... The following section is the resizing source only.

Private Sub pResize()
Dim xp As CoDim
Dim xs As CoDim
' -------------------------------------------------------------------
' Helper:   Resize the controls on the form to suit its size.
' -------------------------------------------------------------------
    On Error Resume Next
    
    ' available height for the inner objects to the toolbar and status bar.
    xs.Height = ScaleHeight: xs.Width = ScaleWidth
    If picTB.Visible Then ' ... picTB > Toolbar.
        xs.Height = xs.Height - picTB.Height
        xs.Top = picTB.Height
    End If
    
    If picSB.Visible Then xs.Height = xs.Height - picSB.Height ' ... picSB > Status Bar.
    
    ' far left object width and height.
    xp.Top = xs.Top
    xp.Height = xs.Height: xp.Width = xs.Width
    
    If tvProj.Visible Then ' ... tvProject > Project Explorer.
        
        xp.Left = picSplitProj.Left
        xp.Width = cSplitterWidth
        
        If mSizeEditorOnly = False Then
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
    
    ' second left object width and height.
    If picMembers.Visible Then ' ... picMembers > Class Explorer.
        
        xp.Left = picSplitMain.Left
        xp.Width = xp.Left
        
        If mSizeEditorOnly = False Then
            picSplitMain.Move xp.Left, xp.Top, cSplitterWidth, xp.Height
            imgSplitMain.Move xp.Left, xp.Top, cSplitterWidth, xp.Height
            picMembers.Move 0, xp.Top, xp.Width, xp.Height
        End If
        
        xp.Left = xp.Left + cSplitterWidth
        xp.Width = picMain.Width - xp.Left
    
    End If
    
    If mSizeEditorOnly = False Then picCodeCanvas.Move xp.Left, xp.Top, xp.Width, xp.Height ' ... picCodeCanvas > Canvas to Code Viewers.
    
    imgSplitMain.Visible = picMembers.Visible
    
    xp.Left = 0: xp.Top = 0
    picInfo.Move xp.Left, xp.Top, xp.Width
    
    xp.Top = picSplitRTB.Top
    ' far right objects splitter position.
    picSplitRTB.Move xp.Left, xp.Top, xp.Width, (2 * cSplitterHeight)
    imgSplitRTB.Move xp.Left, xp.Top, xp.Width, (2 * cSplitterHeight)
    
    ' far right back object height and width.
    xp.Height = xp.Top - picInfo.Height
    xp.Top = picInfo.Height
    rtb.Move xp.Left, xp.Top, xp.Width, xp.Height
    
    ' far right top object height and width.
    xp.Top = picSplitRTB.Top + (2 * cSplitterHeight)
    xp.Height = picMain.Height - xp.Top
    lv.Move xp.Left, xp.Top, xp.Width, xp.Height
    lvFilter.Move xp.Left, xp.Top, xp.Width, xp.Height
    
End Sub

Private Sub cmdViewProj_Click()
    ' ... Hide / Show far left object.
    tvProj.Visible = Not tvProj.Visible
    pResize
End Sub

Private Sub cmdViewToolbar_Click()
    ' ... Hide / Show the toolbar at the top.
    picTB.Visible = Not picTB.Visible
    pResize
End Sub

Private Sub cmdViewStatus_Click()
    ' ... Hide / Show the status bar at the bottom.
    picSB.Visible = Not picSB.Visible
    pResize
End Sub

Private Sub Form_Resize()
    ' ... resize the form as required.
    If mLoading Then Exit Sub
    pResize
End Sub

Private Sub pButtonClick(Index As Integer)

    Select Case Index
        Case 0: cmdOpenFile_Click
        Case 1: cmdViewProj_Click
        Case 3: cmdViewToolbar_Click
        Case 4: cmdViewStatus_Click
        Case 9: cmdFind_Click
    End Select
    
    ClearMemory
    
End Sub

' -------------------------------------------------------------------
' ... General GUI.
' ... (Image) Buttons.
Private Sub imgBtn_Click(Index As Integer)
' -------------------------------------------------------------------
' ... Toolbar Image: Capture and respond to button click request.
' -------------------------------------------------------------------
    pButtonClick Index

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

Private Sub imgSplitMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLMouseDown = True
    picSplitMain.Visible = True
End Sub

Private Sub imgSplitMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    ' split the far right objects or restore top one to full height.
    mSizeEditorOnly = True
    If imgSplitRTB.Top > 100 * Screen.TwipsPerPixelY Then
        picSplitRTB.Top = picInfo.Height ' + 1 * Screen.TwipsPerPixelY
    Else
        picSplitRTB.Top = picCodeCanvas.Height \ 2
    End If
End Sub

Private Sub imgSplitRTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' start vertical split on far right object(s).
    mSizeEditorOnly = True
    mLMouseDown = True
    picSplitRTB.Visible = True
End Sub

Private Sub imgSplitRTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xcd As CoDim, xmm As MinMax
    ' resize the heights of the far left objects.
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
    ' turn off vertical split on far right object(s).
    mLMouseDown = False
    picSplitRTB.Visible = False
    pResize
    mSizeEditorOnly = False
End Sub
' ... End of resizing source.
' -------------------------------------------------------------------

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    pLVHeaderClick lv, ColumnHeader
End Sub

Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)
    pLVItemClick Item
End Sub

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    pLVHeaderClick lvFilter, ColumnHeader
End Sub

Private Sub pLVHeaderClick(pLV As ListView, pColumnHeader As ColumnHeader)

' ... sort one of the iist views (both have same header and use).

Dim lngSortIndex As Long
Dim lngSortOrder As Long

    On Error GoTo ErrHan:
    
    '... trap a invalid object error.
    lngSortOrder = pLV.SortOrder
    lngSortIndex = pColumnHeader.Index
        
    Select Case lngSortIndex
        
        ' hidden:   7 = line number.
        ' visible:  3 = line number.
        
        Case 1, 3, 6, 7 ' Source Name, Line No., Text Line with substring, sortable line number.
            ' ... defer nnumeric/date fields to their
            ' ... string counterparts for sorting.
            If lngSortIndex = 3 Then lngSortIndex = 7
            
            ' ... reverse current sort order.
            lngSortOrder = IIf(lngSortOrder = lvwAscending, lvwDescending, lvwAscending)
            
            pLV.SortOrder = lngSortOrder
            pLV.SortKey = lngSortIndex - 1
            pLV.Sorted = True
            
    End Select

ResumeError:

Exit Sub

ErrHan:

    Debug.Print "frmSearchVBP.pLVHeaderClick.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub


Private Sub pLVItemClick(pLVItem As ListItem)

' ... try to find the item from its key, load its file (if not already loaded)
' ... find and highlight the item in the RTF box setting it to the
' ... first visible line.

Dim sKey As String
Dim lngLine As Long
Dim lngFound As Long
Dim sFile As String
Dim lngScroll As Long
Dim lngPos As Long

    On Error GoTo ErrHan:
    
    If pLVItem Is Nothing Then Err.Raise vbObjectError + 1000, , "List Item is not Valid."
    
    sKey = pLVItem.SubItems(7) ' pLVItem.Key
    lblMember.Caption = pLVItem.Text
    
    If Len(sKey) Then
    
        lngFound = InStr(1, sKey, "|")
        If lngFound > 0 Then
                        
            sFile = Left$(sKey, lngFound - 1)
            lngLine = CLng(Val(Mid$(sKey, lngFound + 1)))
            
            lngFound = InStr(lngFound + 1, sKey, "|")
            If lngFound > 0 Then
                lngPos = CLng(Val(Mid$(sKey, lngFound + 1))) - 1
            End If
            
            If rtb.FileName <> sFile Then
                ' ... clear text and formatting.
                ' ... had issues with previous formatting surviving.
                rtb.Text = ""
                rtb.SelStart = 0
                rtb.SelLength = Len(rtb.Text)
                rtb.SelBold = False
                rtb.SelColor = vbBlack
                rtb.SelStart = 0
                
                rtb.FileName = sFile
                
                ' ... update the status bar caption with the file name.
                lblFileName.Caption = sFile
                
            Else
                If Len(rtb.Text) = 0 Then
                    ' ... presume we cleared the text in code
                    ' ... and try reloading file.
                    rtb.LoadFile sFile
                End If
            End If
            
            lngScroll = lngLine
            rtb.SelStart = 0
                        
            rtb.ZOrder
            picSplitMain.ZOrder
            
            ' -------------------------------------------------------------------
            ' ... scroll to the line in the text.
            If lngScroll > 0 Then
                ' ... scroll to the line number.
                modGeneral.ScrollRTFBox rtb.hwnd, lngScroll
            End If
            
            ' -------------------------------------------------------------------
            ' ... select the string matched.
            rtb.SelStart = lngPos
            rtb.SelLength = Len(txtFind.Text)
            rtb.SelBold = True
            rtb.SelColor = vbBlue
        
        End If
    
    End If

ResumeError:

Exit Sub

ErrHan:

    Debug.Print "frmSearchVBP.pLVItemClick.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub lvFilter_ItemClick(ByVal Item As ComctlLib.ListItem)
    pLVItemClick Item
End Sub

Private Sub tvProj_NodeClick(ByVal Node As ComctlLib.Node)

' ... Load search results from selected node.

Dim xItem As ListItem
Dim sFind As String
Dim zItem As ListItem
Dim sTag As String
Dim sKey As String

' Note:
'   I have added two list views for displaying the data.
'   the first will be loaded with all the results,
'   the second is filled from filtered results found in first list view.
'   a bit cheap perhaps but easy to implement, and why not?

    On Error GoTo ErrHan:
    
    lvFilter.ListItems.Clear
    lvFilter.Sorted = False
    lblFileName.Caption = ""
    lblMember.Caption = ""
    rtb.FileName = ""
    
    If Node Is Nothing Then Err.Raise vbObjectError + 1000, , "Project Node not Valid."
    If tvProj.Nodes.Count = 0 Then Err.Raise vbObjectError + 1000, "", "Project Tree Not Valid for raising Node Click events."
    
    Screen.MousePointer = vbHourglass
    
    If Node.Index = 1 Then
        ' ... this is the main node; all found items visible.
        lv.ZOrder
    Else
        ' ... node is either a group header or source file.
        ' ... group headers are recognised as having a tag value of
        ' ... Forms, Classes, Modules or UserControls.
        
        sTag = Trim$(Node.Tag)
        
        Select Case sTag
            Case "Forms": sFind = "FRM"
            Case "Classes": sFind = "CLS"
            Case "Modules": sFind = "BAS"
            Case "User Controls": sFind = "CTL"
        End Select
        
        If Len(sFind) Then ' ... e.g. it's a group header, filter by source file type.
            For Each xItem In lv.ListItems
                With xItem
                    If .Tag = sFind Then
                        'Set zItem = lvFilter.ListItems.Add(, .Key, .Text)
                        Set zItem = lvFilter.ListItems.Add(, , .Text)
                        zItem.SubItems(1) = .SubItems(1)
                        zItem.SubItems(2) = .SubItems(2)
                        zItem.SubItems(3) = .SubItems(3)
                        zItem.SubItems(4) = .SubItems(4)
                        zItem.SubItems(5) = .SubItems(5)
                        zItem.SubItems(6) = .SubItems(6)
                        zItem.SubItems(7) = .SubItems(7)
                        zItem.Tag = .Tag
                    End If
                End With
            Next xItem
        Else ' ... it's a source file, filter by file name.
            If sTag = cFileSig Then
                sFind = Node.Key
                For Each xItem In lv.ListItems
                    With xItem
                        sKey = xItem.SubItems(7)
                        'sTag = Left$(.Key, InStr(1, .Key, "|") - 1)
                        sTag = Left$(sKey, InStr(1, sKey, "|") - 1)
                        If sTag = sFind Then
                            'Set zItem = lvFilter.ListItems.Add(, .Key, .Text)
                            Set zItem = lvFilter.ListItems.Add(, , .Text)
                            zItem.SubItems(1) = .SubItems(1)
                            zItem.SubItems(2) = .SubItems(2)
                            zItem.SubItems(3) = .SubItems(3)
                            zItem.SubItems(4) = .SubItems(4)
                            zItem.SubItems(5) = .SubItems(5)
                            zItem.SubItems(6) = .SubItems(6)
                            zItem.SubItems(7) = .SubItems(7)
                            zItem.Tag = .Tag
                        End If
                    End With
                Next xItem
            End If
        End If
        lvFilter.ZOrder
    End If

ResumeError:
    
    picSplitMain.ZOrder
    Screen.MousePointer = vbDefault
    
Exit Sub

ErrHan:

    Debug.Print "frmSearchVBP.tvProj_NodeClick.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
'    Resume
End Sub

Private Sub pParseClass(pString As String, _
                        pDelimiter As String, _
                        pMembers() As String, _
                        pMemberIndexes() As Long)

Dim i As Long
Dim j As Long
Dim sItem As String
Dim bLineIsCommented As Boolean
Dim sFindEnd As String
Dim lngType As Long
Dim sDeclaration As String
Dim lngLineStart As Long
Dim lngFoundSpace As Long
Dim s() As String
Dim uB As Long
Dim m As Long
Dim lngDecEnd As Long

    ' -------------------------------------------------------------------
    On Error GoTo ErrHan:
            
    ' -------------------------------------------------------------------
    modStringArrays.SplitString pString, s, pDelimiter
    
    ReDim pMembers(1199)
    ReDim pMemberIndexes(1199)
    
    pMemberIndexes(0) = 0
    pMembers(0) = "Declarations"
    
    ' -------------------------------------------------------------------
                
        uB = UBound(s)
        
        Do While i < uB
    
            For i = 0 To uB
            
                sItem = LTrim$(s(i))  ' ... trim the line.
                bLineIsCommented = Left$(sItem, 1) = "'" Or Left$(sItem, 3) = "Rem"  ' ... check if line is already commented.
                                                                                     ' ... v2, added Rem to comment test.
                If bLineIsCommented = False And Len(sItem) Then
                    ' -------------------------------------------------------------------
                    ' ... strip accessor if exists.
                    If Left$(sItem, cm_len_Public) = cm_word_Public Then
                        sItem = Mid$(sItem, c_len_Public + 2)
                    ElseIf Left$(sItem, cm_len_Private) = cm_word_Private Then
                        sItem = Mid$(sItem, c_len_Private + 2)
                    ElseIf Left$(sItem, cm_len_Friend) = cm_word_Friend Then
                        sItem = Mid$(sItem, c_len_Friend + 2)
                    End If
                    ' -------------------------------------------------------------------
                    ' ... check for method type signature,
                    ' ... if found, generate end method signature to find next
                    ' ... and capture method name.
                    If Left$(sItem, cm_len_Sub) = cm_word_Sub Then ' v4 fix, make sure is single word and not part of word.
                        lngType = 1
                        sFindEnd = c_word_End & Space$(1) & c_word_Sub
                    ElseIf Left$(sItem, cm_len_Function) = cm_word_Function Then ' v4 fix, make sure is single word and not part of word.
                        lngType = 2
                        sFindEnd = c_word_End & Space$(1) & c_word_Function
                    ElseIf Left$(sItem, cm_len_Property) = cm_word_Property Then ' v4 fix, make sure is single word and not part of word.
                        lngType = 3
                        sFindEnd = c_word_End & Space$(1) & c_word_Property
                    End If
                    
                    If lngType > 0 Then
                    
                        If lngDecEnd = 0 Then
                            lngDecEnd = i - 1
                        End If
                        
                        ' -------------------------------------------------------------------
                        lngLineStart = i
                        ' -------------------------------------------------------------------
                        sDeclaration = sItem
                        ' -------------------------------------------------------------------
                        If Right$(sDeclaration, 1) = "_" Then
                            ' -------------------------------------------------------------------
                            ' ... concatenate the declaration's lines into a single line.
                            sDeclaration = Left$(sDeclaration, Len(sDeclaration) - 1)
                            For j = i + 1 To uB
                                ' -------------------------------------------------------------------
                                ' ... left trim spaces from next line.
                                sDeclaration = sDeclaration & LTrim$(s(j))
                                If Right$(sDeclaration, 1) <> "_" Then
                                    i = j
                                    Exit For
                                Else
                                    ' -------------------------------------------------------------------
                                    ' ... trim line extender.
                                    sDeclaration = Left$(sDeclaration, Len(sDeclaration) - 1)
                                End If
                            Next j
                        End If
                        
                        ' -------------------------------------------------------------------
                        ' ... find the end of the member.
                        For j = i + 1 To uB
                        
                            sItem = LTrim$(s(j))  ' ... trim the line.
                            bLineIsCommented = Left$(sItem, 1) = "'" Or Left$(sItem, 3) = "Rem"    ' ... check if line is already commented.
                            
                            If bLineIsCommented = False And Len(sItem) Then
                                ' -------------------------------------------------------------------
                                ' ... check for line numbering, added this after reading Qwerti60's SetPixel (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73412&lngWId=1)
                                ' ... project where End Sub can be preceded by a line number.
                                ' ... I've written some other code to handle specifically finding End Sub, Function and Property
                                ' ... to get line numbers, it's byte oriented and may be found in mod\modMembers.bas as GetMember_5
                                ' -------------------------------------------------------------------
                                If Val(sItem) > 0 Then
                                    lngFoundSpace = InStr(1, sItem, " ")
                                    If lngFoundSpace > 0 Then
                                        sItem = Mid$(sItem, lngFoundSpace + 1)
                                    End If
                                End If

                                ' ... check for end method signature.
                                If Left$(sItem, Len(sFindEnd)) = sFindEnd Then
                                    
                                    m = m + 1
                                    pMemberIndexes(m) = lngLineStart '+ 1 'j
                                    pMembers(m) = Left$(sDeclaration, InStr(1, sDeclaration, "(") - 1)

                                    i = j ' ... update i for next go.
                                    Exit For
                                    
                                End If
                                
                            End If
                            
                        Next j
                        
                    End If
                
                End If
                
                lngType = 0
                
            Next i
        
        Loop
        
        ReDim Preserve pMembers(m)
        ReDim Preserve pMemberIndexes(m)
        
ResumeError:

    On Error Resume Next
    
    Erase s
    m = 0&
    uB = 0&
    sDeclaration = vbNullString
    sFindEnd = vbNullString
    sItem = vbNullString
    i = 0&
    j = 0&
        
Exit Sub

ErrHan:

    Debug.Print "Code.pParseClass.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pParseClass:

