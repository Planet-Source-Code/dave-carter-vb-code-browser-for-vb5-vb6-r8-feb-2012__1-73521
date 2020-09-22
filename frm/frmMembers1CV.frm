VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMembers 
   Caption         =   "Methods"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   HelpContextID   =   47
   Icon            =   "frmMembers1CV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   5745
   Begin CodeViewer.ucMenus ucMenu 
      Left            =   2460
      Top             =   3360
      _ExtentX        =   1508
      _ExtentY        =   1931
   End
   Begin VB.PictureBox tb 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   5745
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5745
      Begin VB.CommandButton cmdAlpha 
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
         Index           =   0
         Left            =   480
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Clear Alpha Filter"
         Top             =   60
         Width           =   375
      End
      Begin ComctlLib.ProgressBar prBar 
         Height          =   315
         Left            =   480
         TabIndex        =   28
         Top             =   960
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin RichTextLib.RichTextBox rtbName 
         Height          =   855
         Left            =   480
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   420
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1508
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmMembers1CV.frx":058A
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
      Begin VB.Label lblAlpha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   -15
         Width           =   255
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   720
         Width           =   8415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picSplitRTB 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   3060
      ScaleHeight     =   90
      ScaleWidth      =   705
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picSplitMain 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   2220
      ScaleHeight     =   2595
      ScaleWidth      =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7410
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   915
      Left            =   3000
      TabIndex        =   17
      Top             =   3000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1614
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMembers1CV.frx":0605
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
      Height          =   795
      Left            =   3000
      TabIndex        =   15
      Top             =   1980
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1402
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
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Scope"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Parent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PT"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Attributes"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView lvFilter 
      Height          =   795
      Left            =   4320
      TabIndex        =   16
      Top             =   1980
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1402
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
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Scope"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Parent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PT"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Attributes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picNav 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   2415
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2415
      Begin VB.PictureBox picFilter 
         Height          =   5235
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   2175
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   0
         Width           =   2235
         Begin VB.CommandButton cmdClearFilter 
            Caption         =   "Clear"
            Height          =   495
            Left            =   1140
            TabIndex        =   6
            ToolTipText     =   "Click to clear user filter"
            Top             =   1200
            Width           =   915
         End
         Begin VB.ComboBox cboFilter 
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
            Index           =   4
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Tag             =   "Type"
            ToolTipText     =   "Filter by Parent Type"
            Top             =   2220
            Width           =   735
         End
         Begin VB.CommandButton cmdFilter 
            Caption         =   "Filter"
            Height          =   495
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Click for user filter"
            Top             =   1200
            Width           =   915
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "Description"
            Height          =   495
            Index           =   1
            Left            =   1020
            TabIndex        =   4
            ToolTipText     =   "filter on description / attributes"
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "Name"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "filter on name"
            Top             =   720
            Width           =   795
         End
         Begin VB.ListBox lstEvents 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "double-click to filter on selected item"
            Top             =   4320
            Width           =   1995
         End
         Begin VB.TextBox txtFind 
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
            TabIndex        =   2
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox cboFilter 
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
            Index           =   3
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "Parent"
            ToolTipText     =   "Filter by Parent"
            Top             =   3660
            Width           =   1995
         End
         Begin VB.ComboBox cboFilter 
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
            Index           =   2
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "Scope"
            ToolTipText     =   "Filter by Scope"
            Top             =   2940
            Width           =   1995
         End
         Begin VB.ComboBox cboFilter 
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
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "Type"
            ToolTipText     =   "Filter by Type"
            Top             =   2220
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Filter Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Parent"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Scope"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   2700
            Width           =   1875
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Method and Parent Types"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   1920
            Width           =   1935
         End
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "Names"
         ToolTipText     =   "Filter by Name"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Image imgSplitMain 
      Height          =   2715
      Left            =   2460
      MousePointer    =   9  'Size W E
      ToolTipText     =   "Click and Drag to Size or Dbl-Click: Expand/Shrink"
      Top             =   1500
      Width           =   120
   End
   Begin VB.Image imgSplitRTB 
      Height          =   180
      Left            =   3060
      MousePointer    =   7  'Size N S
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1500
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Code Browser Form"
Option Explicit

Private mInitialised As Boolean
Private mFilterChar As String

Private mCurrentMethodText As String

Private WithEvents moVBProject As VBProject
Attribute moVBProject.VB_VarHelpID = -1

Private mHaveVBPInfo As Boolean
Private moVBPInfo As VBPInfo
Private mDoing As Boolean
Private msaMembers As StringArray

Private moCodeInfoArray() As CodeInfo
Private mswCodeRTFArray() As StringWorker
Private mlngSourceFileCount As Long

Private miTopLV As Long

Private Const cWeight As Long = 90
Private Const cSplitMin As Long = 450   ' ... width of alpha button
Private mSplitMin As Long

Private mbMouseDown As Boolean

Implements IReportForm

Private Sub cmdClearFilter_Click()
    txtFind.Text = ""
    pFilterList
    pSetTopMostLV
End Sub

Private Sub cmdFilter_Click()
    pFilterList
    pSetTopMostLV
End Sub

Private Sub IReportForm_Init(pVBPInfo As VBPInfo, Optional pOK As Boolean = False, Optional pErrMsg As String = vbNullString)
    Init pVBPInfo
End Sub

Private Property Get IReportForm_ItemCount() As Long
    IReportForm_ItemCount = mlngSourceFileCount
End Property

Private Sub IReportForm_ZOrder(Optional pOrder As Long = 0&)
    Show
    ZOrder pOrder
End Sub

Public Sub Init(pVBPInfo As VBPInfo)
    
    On Error GoTo ErrHan:
    Screen.MousePointer = vbHourglass
    
    Me.Hide
    DoEvents
    
    pInit
    
    Me.Show
    DoEvents
    
    prBar.Visible = True
    DoEvents
    
    Set moVBPInfo = pVBPInfo
    
    Set moVBProject = New VBProject
    
    moVBProject.Init moVBPInfo
    mlngSourceFileCount = moVBProject.CountOfSourceFiles
    
    prBar.Visible = False
    DoEvents
    
    Me.Tag = moVBPInfo.ProjectName
    
    Set msaMembers = moVBProject.MembersData
    
    mInitialised = True
    
    pLoadMembers
    
    mHaveVBPInfo = True

ResumeError:
    
    On Error Resume Next
    lv.SetFocus
    
    Screen.MousePointer = vbDefault

Exit Sub

ErrHan:

    Debug.Print "frmDictionary.Init.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Public Property Set MembersData(pMembers As StringArray)

    pInit
    Set msaMembers = pMembers
    mInitialised = True
    pLoadMembers
    
End Property

Private Sub pLoadMembers()
    
Dim iCount As Long
Dim xItem As ListItem
Dim xMember As MemberInfo
Dim xFile As FileNameInfo

Dim i As Long
Dim j As Long

Dim bOK As Boolean

Dim saNames As StringArray
Dim saTypes As StringArray
Dim saAccessors As StringArray
Dim saParents As StringArray
Dim saPTypes As StringArray

Dim saTmp As StringArray

    On Error GoTo ErrHan:
    
    If mInitialised = False Then Err.Raise vbObjectError + 1000, , "Class not initialised"
    iCount = msaMembers.Count
    
    If iCount Then
        
        ' -------------------------------------------------------------------
        ' ... set up some string arrays to add unique filter data to the combo boxes
        For i = 1 To 5 '4
            Set saTmp = New StringArray
            saTmp.DuplicatesAllowed = False
            saTmp.Sortable = True
            Select Case i
                Case 1: Set saNames = saTmp
                Case 2: Set saTypes = saTmp
                Case 3: Set saAccessors = saTmp
                Case 4: Set saParents = saTmp
                Case 5: Set saPTypes = saTmp
            End Select
            cboFilter(i - 1).AddItem ""
        Next i
        ' -------------------------------------------------------------------
        ' ... load data into listview and string arrays
        For i = 1 To iCount
            
            ParseMemberInfoItem msaMembers, i, xMember
            
            With xMember
            
                Set xItem = lv.ListItems.Add(, , .Name)
                ' name type scope parent
                xItem.SubItems(1) = .TypeAsString
                xItem.SubItems(2) = .AccessorAsString
                xItem.SubItems(3) = .ParentName
                ' -------------------------------------------------------------------
                modFileName.ParseFileNameEx .ParentFileName, xFile
                xItem.SubItems(4) = xFile.Extension
                ' -------------------------------------------------------------------
                xItem.SubItems(5) = .MethodAttributes
                ' -------------------------------------------------------------------
                
                saNames.AddItemString .Name
                saTypes.AddItemString .TypeAsString
                saAccessors.AddItemString .AccessorAsString
                saParents.AddItemString .ParentName
                saPTypes.AddItemString xFile.Extension
                
                xItem.Tag = .Index
                
            End With
            
        Next i
        
        For i = 1 To 5 '4
            Select Case i
                Case 1: Set saTmp = saNames
                Case 2: Set saTmp = saTypes
                Case 3: Set saTmp = saAccessors
                Case 4: Set saTmp = saParents
                Case 5: Set saTmp = saPTypes
            End Select
            iCount = saTmp.Count
            If iCount Then
                saTmp.Sort
                For j = 1 To iCount
                    cboFilter(i - 1).AddItem saTmp(j)
                Next j
            End If
        Next i
        
    End If
    
    
ResumeError:
        
    If lv.ListItems.Count Then
        lv.ListItems(1).Selected = True
        lv_ItemClick lv.ListItems(1)
        LVSizeColumn lv, 5
    End If
    
    sb.SimpleText = "Number of Items: " & Format$(lv.ListItems.Count, cNumFormat)
    
Exit Sub

ErrHan:

    Debug.Print "frmDictionary.pLoadMembers.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

Private Sub lstEvents_DblClick()
    If lstEvents.ListCount Then
        If lstEvents.ListIndex > -1 Then
            txtFind.Text = lstEvents.Text
            cmdFilter_Click
        End If
    End If
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

Dim lngSortIndex As Long

    ' ... reverse current sort order.
    lv.SortOrder = IIf(lv.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    lngSortIndex = ColumnHeader.Index
        
    lv.SortKey = lngSortIndex - 1
    lv.Sorted = True
    
    lngSortIndex = 0
    
End Sub

Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)

Dim i As Long
Dim xCodeInfo As CodeInfo
Dim sPName As String
Dim stext As String
Dim j As Long
Dim sSyntax As String
Dim tQuickMem As QuickMemberInfo
    
    mCurrentMethodText = vbNullString
    lblName.Tag = ""
    lblName.Caption = Item.Text
    lblDesc.Caption = Item.SubItems(5)
    Caption = "Methods: " & Me.Tag & " | " & Item.SubItems(3) & " | " & Item.SubItems(1) & " | " & Item.Text
    rtbName.Text = ""
    sPName = Item.SubItems(3)
    j = CLng(Val(Item.Tag))
    For i = 0 To mlngSourceFileCount - 1
        Set xCodeInfo = moCodeInfoArray(i)
        If xCodeInfo.Name = sPName Then
            mCurrentMethodText = xCodeInfo.GetMemberCodeLines(j)
            stext = BuildRTFString(mCurrentMethodText, , , , , , True)
            rtb.TextRTF = stext
            tQuickMem = xCodeInfo.QuickMember(j)
            GetMemberSyntax tQuickMem.Declaration, sSyntax
            If Len(tQuickMem.ValueType) Then
                Caption = Caption & ": " & tQuickMem.ValueType
'                lblName.Caption = "[" & LCase$(Left$(Item.SubItems(1), 1)) & "]" & " " & Item.Text & ": " & tQuickMem.ValueType & " ( " & sSyntax & " ) "
                lblName.Caption = "[" & LCase$(Left$(Item.SubItems(1), 1)) & "]" & " " & sPName & "." & Item.Text & " = " & tQuickMem.ValueType & " ( " & sSyntax & " ) "
                
                lblName.Tag = "done"
            Else
                lblName.Caption = "[" & LCase$(Left$(Item.SubItems(1), 1)) & "]" & " " & sPName & "." & Item.Text & " ( " & sSyntax & " ) "
            End If
            rtbName.TextRTF = BuildRTFString(lblName.Caption & vbNewLine & lblDesc.Caption, "Tahoma", , "10", , , , True)
            Exit For
        End If
    Next i

    stext = vbNullString
    sPName = vbNullString
    
'    If lblName.Tag = "" Then
'        lblName.Caption = "[" & LCase$(Left$(Item.SubItems(1), 1)) & "]" & " " & Item.Text & ": ( " & sSyntax & " ) "
'    End If
    
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        pLVPopUpMenu lv
    End If
End Sub


Private Sub lvFilter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        pLVPopUpMenu lvFilter
    End If
End Sub

Private Sub pLVPopUpMenu(pLV As ListView)
    If pLV Is Nothing Then Exit Sub
    If pLV.ListItems.Count = 0 Then Exit Sub
    If pLV.SelectedItem Is Nothing Then Exit Sub
    ucMenu.ShowClassMenu
End Sub

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

Dim lngSortIndex As Long

    ' ... reverse current sort order.
    lvFilter.SortOrder = IIf(lvFilter.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    lngSortIndex = ColumnHeader.Index
        
    lvFilter.SortKey = lngSortIndex - 1
    lvFilter.Sorted = True
    
    lngSortIndex = 0
    

End Sub

Private Sub lvFilter_ItemClick(ByVal Item As ComctlLib.ListItem)
    
    lv_ItemClick Item
        
End Sub

Private Sub cboFilter_Click(Index As Integer)
    pFilterList
End Sub

Private Sub cmdAlpha_Click(Index As Integer)
    
    lblAlpha.Caption = vbNullString
    
    If Index = 0 Then
        mFilterChar = vbNullString: miTopLV = 0
    Else
        mFilterChar = cmdAlpha(Index).Caption
        lblAlpha.Caption = mFilterChar
        mFilterChar = mFilterChar & "*"
    End If
    
    pFilterList
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

Dim i As Long
Dim j As Long

    lblAlpha.Caption = vbNullString
    
    Select Case KeyAscii
        Case 65 To 90
            mFilterChar = Chr$(KeyAscii) ' & "*"
        Case 97 To 122
            mFilterChar = UCase$(Chr$(KeyAscii)) ' & "*"
        Case vbKeyEscape
            j = 1: mFilterChar = "": miTopLV = 0
        Case Else
            KeyAscii = 0
    End Select
    If KeyAscii Or j Then
        If j = 0 Then
            lblAlpha.Caption = mFilterChar
            mFilterChar = mFilterChar & "*"
        End If
        pFilterList
'        i = KeyAscii - 97 + 1
'        On Error Resume Next
'        cmdAlpha(i).SetFocus
    Else
        pSetTopMostLV
    End If
    
End Sub

Private Sub moVBProject_ProcessFile(ByVal ItemIndex As Long, ByVal MaxItems As Long, ByVal ItemName As String)

    lblDesc.Caption = ItemIndex & " of " & MaxItems & ": " & ItemName
    lblDesc.Refresh
    With prBar
        .Min = 0
        .max = MaxItems
        .Value = ItemIndex
    End With

End Sub

Private Sub pResize()

Dim iTBHeight As Long   ' ... toolbar height
Dim iSBHeight As Long   ' ... status bar height
Dim iLeft As Long       ' ... left
Dim iHeight As Long     ' ... available height
Dim iTop As Long        ' ... top
Dim iWidth As Long      ' ... width
Dim iHTop As Long       ' ... new top
Dim iHHeight As Long    ' ... new height

    On Error Resume Next
    ' -------------------------------------------------------------------
    If WindowState = vbMinimized Then Exit Sub                  ' ... quit if minimised
    ' -------------------------------------------------------------------
    iHeight = ScaleHeight                                       ' ... available height
    If tb.Visible Then
        iTBHeight = tb.Height                                   ' ... toolbar visible, reduce available height
        iTop = iTBHeight                                        ' ... top
    End If
    If sb.Visible Then
        iSBHeight = sb.Height                                   ' ... status bar visible, reduce available height
    End If
    iHeight = iHeight - iTBHeight - iSBHeight                   ' ... available height
    ' -------------------------------------------------------------------
    iWidth = picSplitMain.Left                                  ' ... tree view width
    ' -------------------------------------------------------------------
    picNav.Move 0, iTop, iWidth, iHeight                         ' ... tree view
    imgSplitMain.Move iWidth, iTop, cWeight, iHeight            ' ... v image splitter
    picSplitMain.Move iWidth, iTop, cWeight, iHeight            ' ... v pic splitter
    ' -------------------------------------------------------------------
    iLeft = iWidth + cWeight                                    ' ... new left
    iWidth = ScaleWidth - iLeft - (1 * Screen.TwipsPerPixelX)   ' ... new width
    iHTop = picSplitRTB.Top                                     ' ... h split top
    iHHeight = iHTop - iTop - (1 * Screen.TwipsPerPixelY)       ' ... new lv height
    ' -------------------------------------------------------------------
    lv.Move iLeft, iTop, iWidth, iHHeight                       ' ... list view
    lvFilter.Move iLeft, iTop, iWidth, iHHeight
    ' -------------------------------------------------------------------
    imgSplitRTB.Move iLeft, iHTop, iWidth                       ' ... h image splitter
    picSplitRTB.Move iLeft, iHTop, iWidth                       ' ... h pic splitter
    ' -------------------------------------------------------------------
    iHTop = iHTop + picSplitRTB.Height                                  ' ... new top
    iHHeight = iHeight - iHTop + (1 * Screen.TwipsPerPixelY) + iTop     ' ... new rtb height
    ' -------------------------------------------------------------------
    rtb.Move iLeft, iHTop, iWidth, iHHeight                             ' ... rich text box
    ' -------------------------------------------------------------------
    lblDesc.Width = ScaleWidth - lblDesc.Left 'tb.Width - (2 * lblDesc.Left)
    lblName.Width = ScaleWidth - lblName.Left ' tb.Width - (2 * lblName.Left)
    ' -------------------------------------------------------------------
    rtbName.Width = lblName.Width
    picFilter.Height = picNav.Height - (1 * Screen.TwipsPerPixelY)
    lstEvents.Height = picFilter.Height - lstEvents.Top - 60
    sb.ZOrder
    ' -------------------------------------------------------------------
    prBar.Width = ScaleWidth - prBar.Left - 90
    
End Sub

Private Sub Form_Load()
    pInit
    ClearMemory
End Sub

Private Sub Form_Resize()
    pResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pRelease
    ClearMemory
End Sub

Private Sub pFilterList()

Dim i As Long
Dim c As Long
Dim j As Long
Dim k As Long
Dim f As Long

Dim sData() As String
Dim sTmp As String
Dim xItem As ListItem
Dim zItem As ListItem
Dim bAdd As Boolean

Dim sFind As String
Dim bDoFind As Boolean
Dim bfindName As Boolean

    On Error GoTo ErrHan:
    
    sb.SimpleText = "Directory"
    ' -------------------------------------------------------------------
    rtb.TextRTF = ""
    rtbName.TextRTF = ""
    ' -------------------------------------------------------------------
    sFind = Trim$(txtFind.Text)
    
    bDoFind = Len(sFind)
    bfindName = optFilter(0).Value And True
    
    c = cboFilter.Count
    
    If c Then
        ReDim sData(c - 1)
        For i = 0 To c - 1
            If cboFilter(i).ListIndex > 0 Then
                sData(i) = cboFilter(i).Text
                j = j + 1
            End If
        Next i
    End If
    
    If j Or Len(mFilterChar) Or bDoFind Then
    
        lvFilter.ListItems.Clear
        lvFilter.Sorted = False
        
        ' ... filter is present, show filtered list view.
        For Each xItem In lv.ListItems
            k = 0
            f = 0
            For i = 0 To c - 1
                If i = 0 Then
                    sTmp = xItem.Text
                Else
                    sTmp = xItem.SubItems(i)
                End If
                If i = 0 Then
                    If Len(mFilterChar) Then
                        If UCase$(sTmp) Like mFilterChar Then
                            f = f + 1
                        End If
                    Else
                        If sTmp = sData(i) Then
                            k = k + 1
                        End If
                    End If
                Else
                    If sTmp = sData(i) Then
                        k = k + 1
                    End If
                End If
            Next i
            bAdd = k = j
            If Len(mFilterChar) Then
                If f = 0 Then
                    bAdd = False
                End If
            End If
            If bDoFind Then
                If bAdd Then
                    If bfindName Then
                        bAdd = InStrB(1, xItem.Text, sFind)
                    Else
                        bAdd = InStrB(1, xItem.SubItems(5), sFind)
                    End If
                End If
            End If
            If bAdd Then     ' ... where j is the number of filters added and k is the number of filters matched.
                ' ... add item to filter list view.
                If k And j Or f Or bDoFind Then
                    With xItem
                        Set zItem = lvFilter.ListItems.Add(, , .Text)
                        For i = 1 To lvFilter.ColumnHeaders.Count - 1
                            zItem.SubItems(i) = .SubItems(i)
                        Next i
                        zItem.Tag = xItem.Tag
                    End With
                End If
            End If
        Next xItem
        
        miTopLV = 1
        
        If lvFilter.ListItems.Count Then
            lvFilter.ListItems(1).Selected = True
            lvFilter_ItemClick lvFilter.ListItems(1)
            lvFilter.ListItems(1).EnsureVisible
            LVSizeColumn lvFilter, 5
        End If
        sb.SimpleText = "Number of Items Showing: " & Format$(lvFilter.ListItems.Count, cNumFormat) & " of " & Format$(lv.ListItems.Count, cNumFormat)
    Else
        ' ... no filter, show main list view.
        If lv.ListItems.Count Then
            lv.ListItems(1).Selected = True
            lv_ItemClick lv.ListItems(1)
            lv.ListItems(1).EnsureVisible
            LVSizeColumn lv, 5
        End If
        miTopLV = 0
        
        sb.SimpleText = "Number of Items: " & Format$(lv.ListItems.Count, cNumFormat)
        
    End If
    
ResumeError:
    
    Erase sData
    i = 0
    c = 0
    j = 0
    
    pSetTopMostLV
    
Exit Sub

ErrHan:

    Debug.Print "frmDictionary.pFilterList.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pFilterList:

Private Sub pLoadButtons()

Dim i As Long
Dim w As Long
Dim l As Long
Dim t As Long
Dim h As Long
    
    On Error Resume Next
    
    l = lblName.Left
    
    cmdAlpha(0).Left = l 'lblName.Left ' 0
    
    w = cmdAlpha(0).Width
    t = cmdAlpha(0).Top
    h = cmdAlpha(0).Height
    
    For i = 1 To 26
        l = l + w      ' ... horizontal buttons, left to right
'        t = t + h       ' ... vertical buttons, top to bottom
        Load cmdAlpha(i)
        With cmdAlpha(i)
            .Move l, t, w, h
            .Caption = Chr$(i + 64)
            .ToolTipText = "Alpha Filter on " & Chr$(i + 64)
            .Visible = True
        End With
    Next i

End Sub

Private Sub pInit()
    
Dim saList As StringArray
Dim i As Long
Dim c As Long

    mInitialised = False
    ' -------------------------------------------------------------------
    pRelease
    pLoadButtons
    mSplitMin = 30 ' cmdAlpha(0).Width + (1 * Screen.TwipsPerPixelX)
    ' -------------------------------------------------------------------
    ' ... set width of splitter controls
    picSplitMain.Width = cWeight
    imgSplitMain.Width = cWeight
    picSplitRTB.Height = cWeight
    imgSplitRTB.Height = cWeight
    ' -------------------------------------------------------------------
    
    lv.Sorted = False
    lvFilter.Sorted = False
    
    LVFullRowSelect lv.hwnd
    LVFullRowSelect lvFilter.hwnd
    
    lv.ZOrder
    ' -------------------------------------------------------------------
    
    WordWrapRTFBox rtb.hwnd
    WordWrapRTFBox rtbName.hwnd
    
    ' -------------------------------------------------------------------
    lstEvents.Clear
    ' -------------------------------------------------------------------
    Set saList = New StringArray
    saList.FromFile App.Path & "\dirmemb.txt", vbCrLf
    c = saList.Count
    If c Then
        saList.Sortable = True
        saList.Sort
        For i = 1 To c
            lstEvents.AddItem saList(i)
        Next i
    End If
    Set saList = Nothing
    ' -------------------------------------------------------------------
    optFilter(0).Value = True
    ' -------------------------------------------------------------------
End Sub

Private Sub pSetTopMostLV()
    
    On Error Resume Next
    
    If miTopLV Then
        lvFilter.ZOrder
        lvFilter.SetFocus
    Else
        lv.ZOrder
        lv.SetFocus
    End If
    
End Sub

Private Sub pRelease()
Dim i As Integer
    
    On Error Resume Next
    
    Set msaMembers = Nothing
    Set moVBPInfo = Nothing
    Set moVBProject = Nothing
    
    If lv.ListItems.Count Then
        lv.ListItems(1).Selected = True
        lv.ListItems.Clear
    End If
    If lvFilter.ListItems.Count Then
        lvFilter.ListItems(1).Selected = True
        lvFilter.ListItems.Clear
    End If
    
    For i = 0 To cboFilter.Count - 1
        cboFilter(i).Clear
    Next i
    
    mFilterChar = vbNullString
    
    mbMouseDown = False
    
    For i = 26 To 1 Step -1
        Unload cmdAlpha(i)
    Next i
    
    i = 0
    miTopLV = 0
    
End Sub

Private Sub pShowHideSplit(Optional ByVal pIndex As Long = 0)
    If pIndex = 0 Then
        picSplitMain.Visible = mbMouseDown
        If mbMouseDown Then picSplitMain.ZOrder
    ElseIf pIndex = 1 Then
        picSplitRTB.Visible = mbMouseDown
        If mbMouseDown Then picSplitRTB.ZOrder
    End If
End Sub ' ... show / hide pic splitter, mouse down = visible

Private Sub pButtonDown(Optional ByVal pIndex As Long = 0, Optional ByVal pOn As Boolean = True)
    mbMouseDown = pOn
    pShowHideSplit pIndex
End Sub

Private Sub pButtonUp(Optional ByVal pIndex As Long = 0, Optional ByVal pOn As Boolean = False)
    mbMouseDown = pOn
    pShowHideSplit pIndex
    pResize
End Sub

Private Sub imgSplitMain_DblClick()
Attribute imgSplitMain_DblClick.VB_Description = "Attempts to resize the rolodex and filters picture box to fit either buttons only or buttons and filters."

Dim iLeftA As Long
Dim iLeftB As Long
Dim x As Long
Dim ixA As Long
Dim ixB As Long

    'iLeftA = cmdAlpha(0).Width + cmdAlpha(0).Left + (1 * Screen.TwipsPerPixelX)
    iLeftA = 30 ' cmdAlpha(0).Width + (1 * Screen.TwipsPerPixelX)
    
'    iLeftB = cboFilter(0).Width + cboFilter(0).Left + (cboFilter(0).Left - cmdAlpha(0).Width) + (1 * Screen.TwipsPerPixelX)
    iLeftB = picFilter.Width ' cboFilter(0).Width + cboFilter(0).Left '+ (cboFilter(0).Left - cmdAlpha(0).Width) + (1 * Screen.TwipsPerPixelX)
    
    x = picSplitMain.Left - (1 * Screen.TwipsPerPixelX)
    
    If x <> iLeftA And x <> iLeftB Then
        ixA = x - iLeftA
        ixB = x - iLeftB
        If ixA > ixB Then
            If ixB < 1 Then
                x = iLeftB
            Else
                x = iLeftA
            End If
        Else
            x = iLeftB
        End If
    ElseIf x = iLeftA Then
        x = iLeftB
    Else
        x = iLeftA
    End If
    
    picSplitMain.Left = x
    pResize
        
    iLeftA = 0: iLeftB = 0: ixA = 0: ixB = 0: x = 0
        
End Sub

Private Sub imgSplitMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pButtonDown
End Sub

Private Sub imgSplitMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pButtonUp
End Sub

Private Sub imgSplitRTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pButtonDown 1
End Sub

Private Sub imgSplitRTB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pButtonUp 1
End Sub

Private Sub imgSplitMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim max As Long
Dim xx As Long
    ' -------------------------------------------------------------------
    If mbMouseDown = False Then Exit Sub
    ' -------------------------------------------------------------------
    xx = x + imgSplitMain.Left
    max = ScaleWidth - (2 * mSplitMin)
    If xx < mSplitMin Then
        xx = mSplitMin
    ElseIf xx > max Then
        xx = max
    End If
    ' -------------------------------------------------------------------
    picSplitMain.Left = xx
    ' -------------------------------------------------------------------
End Sub ' ... move v splitter

Private Sub imgSplitRTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim maxTop As Long
Dim maxBottom As Long
Dim yy As Long
    ' -------------------------------------------------------------------
    If mbMouseDown = False Then Exit Sub
    ' -------------------------------------------------------------------
    yy = y + imgSplitRTB.Top
    maxTop = cSplitMin
    maxBottom = ScaleHeight - cSplitMin
    If tb.Visible Then
        maxTop = maxTop + tb.Height
    End If
    If sb.Visible Then
        maxBottom = maxBottom - sb.Height
    End If
    ' -------------------------------------------------------------------
    If yy < maxTop Then
        yy = maxTop
    ElseIf yy > maxBottom Then
        yy = maxBottom
    End If
    ' -------------------------------------------------------------------
    picSplitRTB.Top = yy
    ' -------------------------------------------------------------------
End Sub ' ... move h splitter

Private Sub moVBProject_MembersAndRTFArrays(CodeMembers() As CodeInfo, CodeAsRTF() As StringWorker)
     moCodeInfoArray = CodeMembers
     mswCodeRTFArray = CodeAsRTF
End Sub

Private Sub txtFind_GotFocus()
    KeyPreview = False
End Sub

Private Sub txtFind_LostFocus()
    KeyPreview = True
End Sub

Private Sub ucMenu_MenuItemClick(Caption As String, Menu As Long, Item As Long)
'    Debug.Print Caption, Menu, Item
Dim sTmp As String
    Select Case Menu
        Case 200 ' class menu
            Select Case Item
                Case 1 ' copy signature
                    '
                Case 2 ' copy method
                    sTmp = mCurrentMethodText
                    
            End Select
    End Select
    If Len(sTmp) > 0 Then
        Clipboard.Clear
        Clipboard.SetText sTmp
    End If
End Sub
