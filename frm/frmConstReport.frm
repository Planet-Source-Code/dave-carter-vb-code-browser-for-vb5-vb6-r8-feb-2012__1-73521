VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmConstReport 
   Caption         =   "Project Constants"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   4
   Icon            =   "frmConstReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   6930
   Begin ComctlLib.ListView lv 
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "If items, double-click to copy selected item's declaration"
      Top             =   1260
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Scope"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6930
      TabIndex        =   3
      Top             =   0
      Width           =   6930
      Begin VB.ComboBox cboTypes 
         Height          =   315
         Left            =   2340
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Select Data Type Filter"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cboSource 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select Source File Filter"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblConstDec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Declaration"
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
         Left            =   120
         TabIndex        =   6
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   900
      End
      Begin VB.Label lblPInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Description: "
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
         Index           =   1
         Left            =   6420
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblPInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Title: "
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
         Index           =   0
         Left            =   4980
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6930
      TabIndex        =   1
      Top             =   3180
      Width           =   6930
      Begin VB.Label lblItemCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Coiunt:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   900
      End
   End
   Begin ComctlLib.ListView lvFilter 
      Height          =   1575
      Left            =   1980
      TabIndex        =   8
      ToolTipText     =   "If items, double-click to copy selected item's declaration"
      Top             =   1320
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Scope"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "frmConstReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface for viewing constants declared at module level throughout a project's source files."

' what?
'  a visual user interface to viewing Constants declared
'  in the declarative section of a source code files within a project.
' why?
'  to make this information more readily available to the developer.
' when?
'  there is a need to check out constants declared in a vbp.
' how?
'  this form is called from the main browser form (frmViewer), via its
'  project explorer pop-up menu.
'  Use the Source and Data Type Combo Boxes to filter the visible items.
'  Double-Click an item in a list view to copy its declaration to the clipboard.
'  Click a list view header to sort list contents by that field.
' who?
'  d.c.

Option Explicit

Private moVBPInfo As VBPInfo
Private mAttributeDelimiter As String
Private mInitialised As Boolean

Private moTypes As StringArray

' -------------------------------------------------------------------
' v6, report form interface.
Private mItemCount As Long

Implements IReportForm

' ... IReportForm Interface methods.

Private Sub IReportForm_Init(pVBPInfo As VBPInfo, Optional pOK As Boolean = False, Optional pErrMsg As String = vbNullString)
Attribute IReportForm_Init.VB_Description = "IReportForm Interface to Initialising this class."
    Init pVBPInfo, pOK, pErrMsg
End Sub ' ... IReportForm_Init:

Private Property Get IReportForm_ItemCount() As Long
Attribute IReportForm_ItemCount.VB_Description = "IReportForm Interface to returning the number of items in the report's data."
    IReportForm_ItemCount = ItemCount
End Property ' ... IReportForm_ItemCount: Long

Private Sub IReportForm_ZOrder(Optional pOrder As Long = 0&)
Attribute IReportForm_ZOrder.VB_Description = "IReportForm Interface to ordering this form on top of other windows."
    ZOrder pOrder
End Sub ' ... IReportForm_ZOrder:

' -------------------------------------------------------------------

Public Property Get ItemCount() As Long
Attribute ItemCount.VB_Description = "Returns the number of Constants declared in the VB Project (only reads from Declarations Section)."
    ItemCount = mItemCount
End Property ' ... ItemCount: Long

Public Sub Init(ByRef pVBPInfo As VBPInfo, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Init.VB_Description = "Initialises the class with a loaded VBPInfo instance and then tries to create the report's data."
    
'... Parameters.
'    R__ pVBPInfo: VBPInfo           ' ... A VBPInfo instance loaded with data.

Dim bOK As Boolean
Dim sErrMsg As String

    On Error GoTo ErrHan:
    Screen.MousePointer = vbHourglass
    bOK = Not pVBPInfo Is Nothing
    If bOK = False Then
        Err.Raise vbObjectError + 1000, , "VBP Info object not instanced."
    Else
        bOK = pVBPInfo.Initialised
        If bOK = False Then
            Err.Raise vbObjectError + 1000, , "VBP Info object not initialised."
        End If
    End If
    
    pInit
    
    lv.Visible = False
    
    Set moVBPInfo = pVBPInfo
    mAttributeDelimiter = moVBPInfo.AttributeDelimiter
    lblPInfo(0).Caption = moVBPInfo.Title
    lblPInfo(1).Caption = moVBPInfo.Description
    
    Let sErrMsg = vbNullString
    Let bOK = True

ErrResume:
    
    On Error GoTo 0
    
    Let pErrMsg = sErrMsg
    Let pOK = bOK
    
    If bOK Then
        pGenerateConstsReport
    End If
    
    Me.Caption = IIf(mItemCount, mItemCount & " ", "") & "Constants found in " & moVBPInfo.Title
    
    lblItemCount.Caption = "Item Count: " & Format$(lv.ListItems.Count, cNumFormat)
    
    If lv.ListItems.Count Then
        lv.ListItems(1).Selected = True
        pLVItemClick lv.ListItems(1)
    End If
    
    mInitialised = bOK
    
    lv.Visible = True
    DoEvents
    Screen.MousePointer = vbDefault

Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "frmConstReport.Init.Error: " & Err.Number & "; " & Err.Description
    Resume ErrResume:

End Sub ' ... Init:

Private Sub pRelease()
Attribute pRelease.VB_Description = "releases all resources used prior to termination / re-use (empties list views)."
    
    On Error Resume Next
    
    cboSource.Clear
    
    If lv.ListItems.Count > 0 Then
        lv.ListItems(1).Selected = True
        lv.ListItems.Clear
    End If
    
    If lvFilter.ListItems.Count > 0 Then
        lvFilter.ListItems(1).Selected = True
        lvFilter.ListItems.Clear
    End If
    
    If Not moVBPInfo Is Nothing Then
        Set moVBPInfo = Nothing
    End If
    
    If Not moTypes Is Nothing Then
        Set moTypes = Nothing
    End If
    
    mItemCount = 0&
    
    lblPInfo(0).Caption = ""
    lblPInfo(1).Caption = ""
    lblConstDec.Caption = ""
    
    mInitialised = False

End Sub ' ... pRelease:

Private Sub pInit()
Attribute pInit.VB_Description = "attempts to initialise the form for use/re-use."
    
    pRelease
    
    Set moTypes = New StringArray
    moTypes.DuplicatesAllowed = False
    moTypes.Sortable = True
    
    lv.Sorted = False
    lvFilter.Sorted = False
    
    cboSource.AddItem ""
    cboTypes.AddItem ""
    
    modGeneral.LVFullRowSelect lv.hwnd
    modGeneral.LVFullRowSelect lvFilter.hwnd
    
End Sub ' ... pInit:

Private Sub Form_Load()
Attribute Form_Load.VB_Description = "tries to load, size & position of form based upon previous size and position."
    LoadFormPosition Me, mdiMain.Height, mdiMain.Width
    ClearMemory
End Sub ' ... Form_Load:

Private Sub Form_Resize()
Attribute Form_Resize.VB_Description = "attempts to maximise the size of the list views given the size of the form."
    
Dim lWidth As Long
Dim lHeight As Long
Dim lTop As Long
Dim lLeft As Long
    
    On Error Resume Next
    
    lWidth = ScaleWidth
    lHeight = ScaleHeight - picSB.Height - picMain.Height
    lTop = picMain.Height
    
    lv.Move lLeft, lTop, lWidth, lHeight
    lvFilter.Move lLeft, lTop, lWidth, lHeight

End Sub ' ... Form_Resize:

Private Sub Form_Unload(Cancel As Integer)
Attribute Form_Unload.VB_Description = "releases all resources on termination."
    
    pRelease
    SaveFormPosition Me
    ClearMemory
    
End Sub ' ... Form_Unload:


Private Sub pGenerateConstsReport()
Attribute pGenerateConstsReport.VB_Description = "method to build data list of constants found."

Dim oTmpA As StringArray
Dim oTmpConsts As StringArray
Dim oCodeInfo As CodeInfo
Dim sTmp As String
Dim lngLoop As Long
Dim lngCount As Long
Dim lngTCount As Long
Dim lngAPILoop As Long
Dim tDataInfo As DataInfo
Dim lngFileLoop As Long
Dim xItem As ListItem
Dim lngLIIndex As Long
Dim tConstInfo As ConstInfo
Dim sDec As String

    On Error GoTo ErrHan:
        
    For lngFileLoop = 1 To 4
            
        Select Case lngFileLoop
            Case 1: sTmp = "Forms": Set oTmpA = moVBPInfo.FormsData
            Case 2: sTmp = "Modules": Set oTmpA = moVBPInfo.ModulesData
            Case 3: sTmp = "Classes": Set oTmpA = moVBPInfo.ClassesData
            Case 4: sTmp = "User Controls": Set oTmpA = moVBPInfo.UserControlsData
        End Select
        
        lngCount = oTmpA.Count
        
        If lngCount > 0 Then
            
            For lngLoop = 1 To lngCount
            
                modGeneral.ParseDataInfoItem oTmpA, lngLoop, tDataInfo, mAttributeDelimiter
                                
                Set oCodeInfo = New CodeInfo
                oCodeInfo.ReadCodeFile tDataInfo.ExtraInfo
                oCodeInfo.Declarations
                
                Set oTmpConsts = oCodeInfo.ConstantsStringArray
                lngTCount = oTmpConsts.Count
                
                If lngTCount > 0 Then
                    
                    cboSource.AddItem oCodeInfo.Name
                    
                    For lngAPILoop = 1 To lngTCount
                    
                        sDec = oTmpConsts(lngAPILoop)
                        modGeneral.ParseConstantsItem sDec, tConstInfo
                        
                        lngLIIndex = lngLIIndex + 1
                        
                        ' source, scope, name, type, value
                        Set xItem = lv.ListItems.Add(, sDec & "|" & lngLIIndex, oCodeInfo.Name) ' oTmpConsts(lngAPILoop))
                        xItem.SubItems(1) = tConstInfo.Scope
                        xItem.SubItems(2) = tConstInfo.Name
                        xItem.SubItems(3) = tConstInfo.Type
                        xItem.SubItems(4) = tConstInfo.Value
                        
                        moTypes.AddItemString tConstInfo.Type
                        
                    Next lngAPILoop
                    
                End If
                
                Set oCodeInfo = Nothing
                
            Next lngLoop
        End If
        
    Next lngFileLoop

ResumeError:
    
    lngCount = moTypes.Count
    If lngCount Then
        moTypes.Sort
        For lngLoop = 1 To lngCount
            cboTypes.AddItem moTypes(lngLoop)
        Next lngLoop
    End If
    
    If Not oTmpA Is Nothing Then
        Set oTmpA = Nothing
    End If
    
    If Not oTmpConsts Is Nothing Then
        Set oTmpConsts = Nothing
    End If
    
    If Not oCodeInfo Is Nothing Then
        Set oCodeInfo = Nothing
    End If
    
    mItemCount = lngLIIndex
    lngLIIndex = 0&
    
Exit Sub

ErrHan:

    Debug.Print "frmConstReport.pGenerateConstsReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pGenerateConstsReport:

Private Sub pLVItemClick(ByVal Item As ComctlLib.ListItem)
Attribute pLVItemClick.VB_Description = "Attempts to refresh the selected Constant's Declaration."

Dim tConstInfo As ConstInfo
Dim sKey As String
Dim sDec As String

    If Item Is Nothing Then Exit Sub
    
    lblConstDec.Caption = ""
    
    sKey = Item.Key
    
    modGeneral.ParseConstantsItem sKey, tConstInfo
    sDec = tConstInfo.Declaration
    
    If Len(sDec) Then
        lblConstDec.Caption = sDec
    End If
        
    lblConstDec.Refresh
    
    sKey = vbNullString
    sDec = vbNullString
    
End Sub ' ... pLVItemClick:

Private Sub pCopyDeclaration()
Attribute pCopyDeclaration.VB_Description = "Copy selected declaration to the clip board."

    If Len(lblConstDec.Caption) Then
        Clipboard.Clear
        Clipboard.SetText lblConstDec.Caption
    End If

End Sub ' ... pCopyDeclaration:

Private Sub pLVColumnClick(pLV As ListView, pColumnHeader As ComctlLib.ColumnHeader)
Attribute pLVColumnClick.VB_Description = "Sort a list view by column header and column values."

Dim lngSortIndex As Long

    If pLV Is Nothing Then Exit Sub
    If pColumnHeader Is Nothing Then Exit Sub
    
    ' ... reverse current sort order.
    pLV.SortOrder = IIf(pLV.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    lngSortIndex = pColumnHeader.Index
        
    pLV.SortKey = lngSortIndex - 1
    pLV.Sorted = True
    
End Sub ' ... pLVColumnClick:

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Attribute lv_ColumnClick.VB_Description = "main list view, sort by column."
    
    pLVColumnClick lv, ColumnHeader

End Sub ' ... lv_ColumnClick:

Private Sub lv_DblClick()
Attribute lv_DblClick.VB_Description = "main list view, copy selected constant declaration."
    
    pCopyDeclaration

End Sub ' ... lv_DblClick:

Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)
Attribute lv_ItemClick.VB_Description = "main list view, refresh selected constant."
    
    pLVItemClick Item

End Sub ' ... lv_ItemClick:

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Attribute lvFilter_ColumnClick.VB_Description = "filter list view, sort by column."

    pLVColumnClick lvFilter, ColumnHeader
    
End Sub ' ... lvFilter_ColumnClick:

Private Sub lvFilter_DblClick()
Attribute lvFilter_DblClick.VB_Description = "filter list view, copy selected constant declaration to the clipboard."
    
    pCopyDeclaration

End Sub

Private Sub lvFilter_ItemClick(ByVal Item As ComctlLib.ListItem)
Attribute lvFilter_ItemClick.VB_Description = "filter list view, refresh the selected constant declaration."
    
    pLVItemClick Item

End Sub ' ... lvFilter_ItemClick:

Private Sub cboSource_Click()
Attribute cboSource_Click.VB_Description = "filter the list data by source file."

    pFilterList

End Sub ' ... cboSource_Click:

Private Sub cboTypes_Click()
Attribute cboTypes_Click.VB_Description = "filter the list data by data type."

    pFilterList
    
End Sub ' ... cboTypes_Click:

Private Sub pFilterList()
Attribute pFilterList.VB_Description = "executes the selected filter on the data displayed and either displays the main or the filter list view."

Dim xItem As ListItem
Dim zItem As ListItem
Dim sSource As String
Dim sType As String
Dim sSourceFind As String
Dim sTypeFind As String
Dim bAddSource As Boolean
Dim bAddType As Boolean
Dim bAdd As Boolean

    On Error GoTo ErrHan:
    
    If cboSource.ListIndex > 0 Then
        sSourceFind = cboSource.Text
    End If
    
    If cboTypes.ListIndex > 0 Then
        sTypeFind = cboTypes.Text
    End If
    
    If Len(sSourceFind) = 0 And Len(sTypeFind) = 0 Then
        
        ' ... no filter.
        If Not lv.SelectedItem Is Nothing Then
            pLVItemClick lv.SelectedItem
        End If
        
        lblItemCount.Caption = "Item Count: " & Format$(lv.ListItems.Count, cNumFormat)
        lv.ZOrder
    
    Else
        
        If lvFilter.ListItems.Count Then
            lvFilter.ListItems(1).Selected = True
            lvFilter.ListItems.Clear
            DoEvents
        End If
        
        For Each xItem In lv.ListItems
            
            ' ... multiple statements!
            bAddSource = False: bAddType = False: bAdd = False
            
            If Len(sSourceFind) Then
                sSource = xItem.Text ' ... source filter.
                bAddSource = CBool(sSourceFind = sSource)
            End If
            
            If Len(sTypeFind) Then
                sType = xItem.SubItems(3) ' ... type filter.
                bAddType = CBool(sType = sTypeFind)
            End If
            
            If Len(sSourceFind) > 0 And Len(sTypeFind) > 0 Then
                bAdd = bAddSource And bAddType
            ElseIf Len(sSourceFind) > 0 And Len(sTypeFind) = 0 Then
                bAdd = bAddSource
            ElseIf Len(sSourceFind) = 0 And Len(sTypeFind) > 0 Then
                bAdd = bAddType
            End If
            
            If bAdd Then
                With xItem
                    Set zItem = lvFilter.ListItems.Add(, .Key, .Text)
                    zItem.SubItems(1) = .SubItems(1)
                    zItem.SubItems(2) = .SubItems(2)
                    zItem.SubItems(3) = .SubItems(3)
                    zItem.SubItems(4) = .SubItems(4)
                End With
            End If
            
        Next xItem
        
        On Error Resume Next
        
        lvFilter.ListItems(1).Selected = True
        pLVItemClick lvFilter.SelectedItem
        
        lblItemCount.Caption = "Item Count: showing " & Format$(lvFilter.ListItems.Count, cNumFormat) & " of " & Format$(lv.ListItems.Count, cNumFormat)
        
        lvFilter.ZOrder
        
    End If
    
ResumeError:
    
    On Error Resume Next
    
    sSource = vbNullString
    sSourceFind = vbNullString
    sType = vbNullString
    sTypeFind = vbNullString
    
    If Not zItem Is Nothing Then
        Set zItem = Nothing
    End If
    If Not xItem Is Nothing Then
        Set xItem = Nothing
    End If
    
Exit Sub

ErrHan:

    Debug.Print "frmConstReport.pFilterList.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pFilterList:
