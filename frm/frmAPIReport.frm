VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAPIReport 
   Caption         =   "Project API Viewer"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   44
   Icon            =   "frmAPIReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   10485
   Begin ComctlLib.ListView lvFilter 
      Height          =   1215
      Left            =   1680
      TabIndex        =   10
      ToolTipText     =   "If items, double-click to copy selected item's declaration"
      Top             =   2280
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
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
         Text            =   "Type"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lib"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Alias"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Parameters"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Return Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin ComctlLib.ListView lv 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "If items, double-click to copy selected item's declaration"
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
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
         Text            =   "Type"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lib"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Alias"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Parameters"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Return Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   10485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10485
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search PSC"
         Height          =   315
         Index           =   4
         Left            =   4440
         TabIndex        =   5
         ToolTipText     =   "Search Planet Source Code"
         Top             =   1080
         Width           =   2235
      End
      Begin VB.ComboBox cboLib 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Select Library Filter"
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search vbAccelerator"
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   4
         ToolTipText     =   "Search vbAccelerator"
         Top             =   1080
         Width           =   1995
      End
      Begin VB.CheckBox chkIncVB6Tag 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include VB6 Tag in search"
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         ToolTipText     =   "Add +vb6 to the search criteria"
         Top             =   1500
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Google"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Search Google"
         Top             =   1500
         Width           =   1935
      End
      Begin VB.ComboBox cboSource 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Source File Filter"
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search MSDN"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         ToolTipText     =   "Search MSDN"
         Top             =   1500
         Width           =   2235
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search All API dot Net"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Search All API dot Net"
         Top             =   1080
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.Label lblAPIDec 
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
         TabIndex        =   1
         ToolTipText     =   "If valid API Name, Click to try finding info on Web"
         Top             =   180
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
         Left            =   7080
         TabIndex        =   12
         Top             =   1560
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
         Left            =   4440
         TabIndex        =   11
         Top             =   660
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10485
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3900
      Width           =   10485
      Begin VB.Label lblItemCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Coiunt:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   90
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmAPIReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface for viewing unique APIs declared in a project.  All being well, one may search the internet for info on selected declaration."

' ------------------------------------------------------------
' Name:         frmAPIReport
' Purpose:      Provides a user interface for viewing APIs Declared in a VB Project.
' Author:       Dave Carter.
' Date:         Friday 17 June 2011
' ------------------------------------------------------------

Option Explicit

Private moVBPInfo As VBPInfo
Private mAttributeDelimiter As String
Private mInitialised As Boolean
Private mLMouseDown As Boolean
Private moAPILibs As StringArray
Private mItemCount As Long

' ... variables for the Hand Cursor.
Private myHandCursor As StdPicture
Private myHand_handle As Long

' -------------------------------------------------------------------
' v6, report form interface.
Implements IReportForm

' ... IReportForm Interface methods.

Private Sub IReportForm_Init(pVBPInfo As VBPInfo, Optional pOK As Boolean = False, Optional pErrMsg As String = vbNullString)
Attribute IReportForm_Init.VB_Description = "IReportForm Interface to Initialise this form."
    Init pVBPInfo, pOK, pErrMsg
End Sub ' ... IReportForm_Init:

Private Property Get IReportForm_ItemCount() As Long
Attribute IReportForm_ItemCount.VB_Description = "IReportForm Interface to the number of items in the report data."
    IReportForm_ItemCount = ItemCount
End Property ' ... IReportForm_ItemCount: Long

Private Sub IReportForm_ZOrder(Optional pOrder As Long = 0&)
Attribute IReportForm_ZOrder.VB_Description = "IReportForm Interface to ordering this form on top of other windows."
    ZOrder pOrder
End Sub ' ... IReportForm_ZOrder:

' -------------------------------------------------------------------

Public Property Get ItemCount() As Long
Attribute ItemCount.VB_Description = "Returns the number of APIs declared in the project being read."
    ItemCount = mItemCount
End Property ' ... ItemCount: Long

Public Sub Init(ByRef pVBPInfo As VBPInfo, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Init.VB_Description = "Attempts to Initialise this form with a VBPInfo instance that will provide access to the data required."

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
    ' -------------------------------------------------------------------
    Set moAPILibs = New StringArray
    moAPILibs.DuplicatesAllowed = False
    moAPILibs.Sortable = True
    ' -------------------------------------------------------------------
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
        pGenerateAPIReport
    End If
    
    Me.Caption = IIf(mItemCount, mItemCount & " ", "") & "APIs found in " & moVBPInfo.Title
    
    mInitialised = bOK
    
    lblItemCount.Caption = "Item Count: " & Format$(lv.ListItems.Count, cNumFormat)
    
    If lv.ListItems.Count Then
        lv.ListItems(1).Selected = True
        lv_ItemClick lv.ListItems(1)
    End If
    
    lv.Visible = True
    DoEvents
    Screen.MousePointer = vbDefault
    
Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "frmAPIReport.Init.Error: " & Err.Number & "; " & Err.Description
    Resume ErrResume:

End Sub ' ... Init:

Private Sub pRelease()
Attribute pRelease.VB_Description = "release all resources used by this class prior to termination / re-use."

    On Error Resume Next
    
    cboSource.Clear
    cboLib.Clear
    
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
    
    If Not moAPILibs Is Nothing Then
        Set moAPILibs = Nothing
    End If
    
    mItemCount = 0&
    
    lblPInfo(0).Caption = ""
    lblPInfo(1).Caption = ""
    lblAPIDec.Caption = ""
    
    mInitialised = False

End Sub ' ... pRelease:

Private Sub pInit()
Attribute pInit.VB_Description = "attempts to set up the form to read a new project."
    
    pRelease
    
    lv.Sorted = False
    lvFilter.Sorted = False
    
    cboSource.AddItem ""
    cboLib.AddItem ""
    
    modGeneral.LVFullRowSelect lv.hwnd
    modGeneral.LVFullRowSelect lvFilter.hwnd

End Sub ' ... pInit:

Private Sub Form_Load()
Attribute Form_Load.VB_Description = "tries to position and dimension form based upon its last size and position."
    
    LoadFormPosition Me, mdiMain.Height, mdiMain.Width
    pLoadHandCursor
    
End Sub ' ... Form_Load:

Private Sub Form_Resize()
Attribute Form_Resize.VB_Description = "tries to size the list views to the maximum size possible given the form's area."
    
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
Attribute Form_Unload.VB_Description = "releases all resources because the form is unloading."
    
    pRelease
    SaveFormPosition Me
    ClearMemory
    
End Sub ' ... Form_Unload:

Private Sub pGenerateAPIReport()
Attribute pGenerateAPIReport.VB_Description = "Attempts to build a list of unique API Declarations in the project."

Dim oTmpA As StringArray
Dim oTmpAPI As StringArray
Dim oCodeInfo As CodeInfo
Dim sTmp As String
Dim lngLoop As Long
Dim lngCount As Long
Dim lngTCount As Long
Dim lngAPILoop As Long
Dim tDataInfo As DataInfo
Dim lngFileLoop As Long
Dim xItem As ListItem
Dim tAPIInfo As APIInfo
Dim lngLIIndex As Long

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
                
                Set oTmpAPI = oCodeInfo.APIStringArray
                
                lngTCount = oTmpAPI.Count
                
                If lngTCount > 0 Then
                
                    cboSource.AddItem oCodeInfo.Name
                    
                    For lngAPILoop = 1 To lngTCount
                        
                        lngLIIndex = lngLIIndex + 1
                        
                        modGeneral.ParseAPIInfoItem oTmpAPI, lngAPILoop, tAPIInfo
                        
                        Set xItem = lv.ListItems.Add(, tAPIInfo.Declaration & "|" & lngLIIndex, oCodeInfo.Name)
                        ' source,type,name,lib,alias, parameters, return value
                        xItem.SubItems(1) = IIf(tAPIInfo.Type = 1, "Sub", "Function")
                        xItem.SubItems(2) = tAPIInfo.Name
                        xItem.SubItems(3) = tAPIInfo.Lib
                        xItem.SubItems(4) = tAPIInfo.Alias
                        xItem.SubItems(5) = tAPIInfo.Parameters
                        xItem.SubItems(6) = tAPIInfo.ReturnValue
                        
                        sTmp = UCase$(tAPIInfo.Lib)
                        If Right$(sTmp, 4) <> ".DLL" Then
                            sTmp = tAPIInfo.Lib & ".dll"
                        Else
                            sTmp = tAPIInfo.Lib
                        End If
                        
                        moAPILibs.AddItemString sTmp
                        
                    Next lngAPILoop
                    
                End If
                
                Set oCodeInfo = Nothing
                
            Next lngLoop
            
        End If
        
    Next lngFileLoop
    
    If Not moAPILibs Is Nothing Then
        lngCount = moAPILibs.Count
        If lngCount Then
            For lngLoop = 1 To lngCount
                sTmp = moAPILibs(lngLoop)
                cboLib.AddItem sTmp
            Next lngLoop
        End If
    End If

ResumeError:
    
    On Error GoTo 0
    
    If cboSource.ListCount Then
        cboSource.ListIndex = 0
    End If
    
    If Not oTmpA Is Nothing Then
        Set oTmpA = Nothing
    End If
    
    If Not oTmpAPI Is Nothing Then
        Set oTmpAPI = Nothing
    End If
    
    If Not oCodeInfo Is Nothing Then
        Set oCodeInfo = Nothing
    End If
    
    sTmp = vbNullString
    mItemCount = lngLIIndex
    lngLIIndex = 0&
    
Exit Sub

ErrHan:

    Debug.Print "frmAPIReport.pGenerateAPIReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pGenerateAPIReport:

Private Sub lblAPIDec_Click()
Attribute lblAPIDec_Click.VB_Description = "attempts to link to the web to search for information on the name of the selected API."

' ... try to search the internet for the named API.
' ... hard coding for now, could become configurable for easier and more robust use.
' ... current search pages include, All API dot Net, MSDN, vbAccelerator and Google.

Dim sName As String
Dim sPage As String
Dim xItem As ListItem
Dim bAddTag As Boolean

' ... of course, these hard coded web site addresses are bound to need updating...
Const cAPIPage As String = "http://allapi.mentalis.org/apilist/" 'apilist.php"
Const cAPIDefault As String = "apilist.php"
Const cMSDNPage As String = "http://social.msdn.microsoft.com/Search/en-us?query="
'Const cGooglePage As String = "http://www.google.com/#hl=en&q="
Const cGooglePage As String = "http://www.google.com/custom?q="
Const cVBAccSuffix As String = "&sa=Search+Google&domains=vbaccelerator.com&sitesearch=vbaccelerator.com"
Const cPSCPage As String = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria="

'http://social.msdn.microsoft.com/Search/en-us?query=[API Name]
'http://allapi.mentalis.org/apilist/[API Name].shtml
'http://www.google.co.uk/#hl=en&q=[API Name]
'http://www.google.com/custom?q=[API Name]&sa=Search+Google&domains=vbaccelerator.com&sitesearch=vbaccelerator.com
'http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=[API Name]&lngWId=1
    
    If cboSource.ListIndex > 0 Or cboLib.ListIndex > 0 Then
        ' ... tells us to look at the filter list view.
        Set xItem = lvFilter.SelectedItem
    ElseIf cboSource.ListIndex = 0 Or cboLib.ListIndex = 0 Then
        ' ... tells us to look at the main list view.
        Set xItem = lv.SelectedItem
    Else
        Exit Sub
    End If

    If Not xItem Is Nothing Then
        
        sName = xItem.SubItems(2)
        
        If Len(sName) Then
            
            bAddTag = chkIncVB6Tag.Value And vbChecked
            
            If optSearch(0).Value Then
                sPage = cAPIPage & cAPIDefault
                sPage = cAPIPage & sName & ".shtml"
            ElseIf optSearch(1).Value Then
                sPage = cMSDNPage & sName & IIf(bAddTag, "+vb6", "")
            ElseIf optSearch(2).Value Then
                sPage = cGooglePage & sName & IIf(bAddTag, "+vb6", "")
            ElseIf optSearch(3).Value Then
                sPage = cGooglePage & sName & cVBAccSuffix
            ElseIf optSearch(4).Value Then
                sPage = cPSCPage & sName & "&lngWId=1"
            End If
            
        End If
        
        If Len(sPage) Then
           modGeneral.OpenWebPage sPage
        End If
        
    End If
        
    sName = vbNullString
    
    If Not xItem Is Nothing Then
    
        Set xItem = Nothing
    
    End If

End Sub ' ... lblAPIDec_Click:

Private Sub lblAPIDec_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute lblAPIDec_MouseDown.VB_Description = "changes forecolor of web link label to indicate that it has been engaged."

    If mLMouseDown Then Exit Sub
    mLMouseDown = True
    lblAPIDec.ForeColor = &H800000 ' &HFF0000
    
End Sub ' ... lblAPIDec_MouseDown:

Private Sub lblAPIDec_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute lblAPIDec_MouseUp.VB_Description = "returns the forecolor of the web link label to indicate it has been dis-engaged."
    
    mLMouseDown = False
    lblAPIDec.ForeColor = &H404040   ' &H800000

End Sub ' ... lblAPIDec_MouseUp:

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Attribute lv_ColumnClick.VB_Description = "main list view, sort by column header."
    
    pLVColumnClick lv, ColumnHeader
    
End Sub ' ... lv_ColumnClick:

Private Sub lv_DblClick()
Attribute lv_DblClick.VB_Description = "copies the selected item's declaration to the clipboard."

    If Len(lblAPIDec.Caption) > 0 Then
        Clipboard.Clear
        Clipboard.SetText lblAPIDec.Caption
    End If
    
End Sub ' ... lv_DblClick:

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Attribute lvFilter_ColumnClick.VB_Description = "filter list view, sort by column header."
    
    pLVColumnClick lvFilter, ColumnHeader
    
End Sub ' ... lvFilter_ColumnClick:

Private Sub pLVColumnClick(pLV As ListView, pColumnHeader As ComctlLib.ColumnHeader)
Attribute pLVColumnClick.VB_Description = "sorts the list view referenced by the column header clicked."

Dim lngSortIndex As Long

    If pLV Is Nothing Then Exit Sub
    If pColumnHeader Is Nothing Then Exit Sub
    
    ' ... reverse current sort order.
    pLV.SortOrder = IIf(pLV.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    lngSortIndex = pColumnHeader.Index
        
    pLV.SortKey = lngSortIndex - 1
    pLV.Sorted = True
    

End Sub ' ... pLVColumnClick:

Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)
Attribute lv_ItemClick.VB_Description = "refreshes the name of the selected API."

    pLVItemClick Item

End Sub ' ... lv_ItemClick:

Private Sub lvFilter_DblClick()
Attribute lvFilter_DblClick.VB_Description = "copies the selected API's Declaration to the clipboard."

    If Len(lblAPIDec.Caption) > 0 Then
        Clipboard.Clear
        Clipboard.SetText lblAPIDec.Caption
    End If

End Sub ' ... lvFilter_DblClick:

Private Sub lvFilter_ItemClick(ByVal Item As ComctlLib.ListItem)
Attribute lvFilter_ItemClick.VB_Description = "refreshes the name of the selected API."

    pLVItemClick Item

End Sub ' ... lvFilter_ItemClick:

Private Sub pLVItemClick(pItem As ListItem)
Attribute pLVItemClick.VB_Description = "refreshes the API Declaration Label's Caption to the selected API."

Dim lngFound As Long
Dim sKey As String

    lblAPIDec.Caption = ""
    
    If pItem Is Nothing Then Exit Sub
    
    sKey = pItem.Key
    
    lngFound = InStr(1, sKey, "|")
    
    If lngFound > 0 Then
        sKey = Left$(sKey, lngFound - 1)
    End If
    
    If Len(sKey) Then
        lblAPIDec.Caption = sKey
    End If
    
    lblAPIDec.Refresh
    
    sKey = vbNullString
    lngFound = 0&


End Sub ' ... pLVItemClick:

Private Sub pLoadHandCursor()
Attribute pLoadHandCursor.VB_Description = "attempts to load the hand cursor icon."

' ... try and load the hand cursor.

    myHand_handle = modHandCursor.LoadHandCursor
    
    If myHand_handle <> 0 Then
        
        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
        
        If myHand_handle = 0 Then Exit Sub
        
        lblAPIDec.MouseIcon = myHandCursor
        lblAPIDec.MousePointer = vbCustom
        
    End If

End Sub ' ... pLoadHandCursor:

Private Sub cboLib_Click()
Attribute cboLib_Click.VB_Description = "filter list data by Library."

    pFilterList

End Sub ' ... cboLib_Click:

Private Sub cboSource_Click()
Attribute cboSource_Click.VB_Description = "filter list data by source file."

    pFilterList

End Sub ' ... cboSource_Click:

Private Sub pFilterList()
Attribute pFilterList.VB_Description = "filter the list data based upon current selected filters."

' ... present available list items dependent upon filter
' ... from cbosource and cbolib text.

' Note:
'   using two list views, the first has all the available data
'   the second shows a filtered subset of the data in the first.
'   so filtering is just running through the items in the first list view
'   and adding them to the second if the filter criteria matches
'   and then forcing the second list view on top with zorder.
'   an empty filter zorders the first list view back on top.

Dim xItem As ListItem
Dim zItem As ListItem
Dim sSource As String
Dim sLib As String
Dim sSourceFind As String
Dim sLibFind As String
Dim bAddSource As Boolean
Dim bAddLib As Boolean
Dim bAdd As Boolean

    On Error GoTo ErrHan:
    
    If cboSource.ListIndex > 0 Then
        sSourceFind = cboSource.Text
    End If
    
    If cboLib.ListIndex > 0 Then
        sLibFind = UCase$(cboLib.Text)
    End If
    
    If Len(sSourceFind) = 0 And Len(sLibFind) = 0 Then
        
        ' ... no filter.
        If Not lv.SelectedItem Is Nothing Then
            lv_ItemClick lv.SelectedItem
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
            bAddSource = False: bAddLib = False: bAdd = False
            
            If Len(sSourceFind) Then
                sSource = xItem.Text ' ... source filter.
                bAddSource = CBool(sSourceFind = sSource)
            End If
            
            If Len(sLibFind) Then
                sLib = UCase$(xItem.SubItems(3)) ' ... library filter.
                If Right$(sLib, 4) <> ".DLL" Then
                    sLib = sLib & ".DLL"
                End If
                bAddLib = CBool(sLib = sLibFind)
            End If
            
            If Len(sSourceFind) > 0 And Len(sLibFind) > 0 Then
                bAdd = bAddSource And bAddLib
            ElseIf Len(sSourceFind) > 0 And Len(sLibFind) = 0 Then
                bAdd = bAddSource
            ElseIf Len(sSourceFind) = 0 And Len(sLibFind) > 0 Then
                bAdd = bAddLib
            End If
            
            If bAdd Then
                With xItem
                    Set zItem = lvFilter.ListItems.Add(, .Key, .Text)
                    zItem.SubItems(1) = .SubItems(1)
                    zItem.SubItems(2) = .SubItems(2)
                    zItem.SubItems(3) = .SubItems(3)
                    zItem.SubItems(4) = .SubItems(4)
                    zItem.SubItems(5) = .SubItems(5)
                    zItem.SubItems(6) = .SubItems(6)
                End With
            End If
            
        Next xItem
        
        If lvFilter.ListItems.Count > 0 Then
        
            lvFilter.ListItems(1).Selected = True
            lvFilter_ItemClick lvFilter.SelectedItem
        
        End If
        
        lblItemCount.Caption = "Item Count: showing " & Format$(lvFilter.ListItems.Count, cNumFormat) & " of " & Format$(lv.ListItems.Count, cNumFormat)
        
        lvFilter.ZOrder
        
    End If
    
ResumeError:
    
    On Error Resume Next
    
    sSource = vbNullString
    sSourceFind = vbNullString
    sLib = vbNullString
    sLibFind = vbNullString
    
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

