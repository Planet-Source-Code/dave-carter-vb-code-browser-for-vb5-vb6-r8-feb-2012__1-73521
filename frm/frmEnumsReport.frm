VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmEnumsReport 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Project Enumerators"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
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
   Icon            =   "frmEnumsReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   6390
   Begin VB.PictureBox picSplitMain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2640
      ScaleHeight     =   1455
      ScaleWidth      =   75
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   75
   End
   Begin RichTextLib.RichTextBox txt 
      Height          =   3015
      Left            =   3120
      TabIndex        =   8
      Top             =   780
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5318
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmEnumsReport.frx":014A
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
   Begin ComctlLib.TreeView tv 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   780
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   5318
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   26
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox picMain 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6390
      TabIndex        =   2
      Top             =   0
      Width           =   6390
      Begin VB.Label lblAPIDec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enumerator Name"
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
         TabIndex        =   5
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   1485
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
         Left            =   4800
         TabIndex        =   4
         Top             =   120
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
         Left            =   3540
         TabIndex        =   3
         Top             =   120
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
      ScaleWidth      =   6390
      TabIndex        =   0
      Top             =   4020
      Width           =   6390
      Begin VB.Label lblItemCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Coiunt:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   90
         Width           =   900
      End
   End
   Begin VB.Image imgSplitMain 
      Height          =   1425
      Left            =   2460
      MousePointer    =   9  'Size W E
      ToolTipText     =   "Resize Me"
      Top             =   840
      Width           =   105
   End
End
Attribute VB_Name = "frmEnumsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface for viewing unique Enumerators declared in a project."

' what?
'  a visual user interface to reading all the Enumerators
'  declared in a project.
' why?
'  to make this information more readily available to the developer.
' when?
'  there is a need to check up on Enumerators declared.
' how?
'  call this form from the code browser (frmViewer) project explorer
'  pop-up menu.
'  select an enumerator from the tree view and read generated code
'  in rich text box.
' who?
'  d.c.

Option Explicit

Private moVBPInfo As VBPInfo
Private mAttributeDelimiter As String
Private mInitialised As Boolean

Private mMouseDown As Boolean
Private Const cHBorder As Long = 60
Private Const cSplitLimit As Long = 660

' -------------------------------------------------------------------
' v6, report form interface.
Private mItemCount As Long

Implements IReportForm

' ... IReportForm Interface methods.

Private Sub IReportForm_Init(pVBPInfo As VBPInfo, Optional pOK As Boolean = False, Optional pErrMsg As String = vbNullString)
    Init pVBPInfo, pOK, pErrMsg
End Sub

Private Property Get IReportForm_ItemCount() As Long
    IReportForm_ItemCount = ItemCount
End Property

Private Sub IReportForm_ZOrder(Optional pOrder As Long = 0&)
    ZOrder pOrder
End Sub

' -------------------------------------------------------------------

Public Property Get ItemCount() As Long
    ItemCount = mItemCount
End Property


Public Sub Init(ByRef pVBPInfo As VBPInfo, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
    
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
    
    tv.Visible = False
    
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
        pGenerateEnumsReport
    End If
    
    Me.Caption = IIf(mItemCount, mItemCount & " unique ", "") & "Enumerators found in " & moVBPInfo.Title
    
    mInitialised = bOK
    
    tv.Visible = True
    DoEvents
    Screen.MousePointer = vbDefault
Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "frmEnumsReport.Init.Error: " & Err.Number & "; " & Err.Description
    Resume ErrResume:

End Sub

Private Sub pRelease()
    
    On Error Resume Next
    
    If tv.Nodes.Count > 0 Then
        tv.Nodes(1).Selected = True
        tv.Nodes.Clear
    End If
    
    If Not moVBPInfo Is Nothing Then
        Set moVBPInfo = Nothing
    End If
    
    mItemCount = 0&
    
    lblPInfo(0).Caption = ""
    lblPInfo(1).Caption = ""
    lblAPIDec.Caption = ""
    
    mInitialised = False

End Sub

Private Sub pInit()
    
    pRelease
    modGeneral.LVFullRowSelect tv.hwnd

End Sub

Private Sub Form_Load()
        
    On Error GoTo ErrHan:
       
    Set tv.ImageList = mdiMain.liMember

    LoadFormPosition Me, mdiMain.Height, mdiMain.Width
    
    modGeneral.WordWrapRTFBox txt.hwnd
    
    ClearMemory
    
ResumeError:

Exit Sub

ErrHan:

    Debug.Print "frmEnumsReport.Form_Load.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:

End Sub

Private Sub Form_Resize()
    
    pResize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    pRelease
    SaveFormPosition Me
    ClearMemory
    
End Sub


Private Sub pGenerateEnumsReport()

Dim oTmpA As StringArray
Dim oCodeInfo As CodeInfo
Dim sTmp As String
Dim lngLoop As Long
Dim lngCount As Long
Dim lngTCount As Long
Dim lngTypesLoop As Long
Dim tDataInfo As DataInfo
Dim lngFileLoop As Long
Dim lngLIIndex As Long
Dim oTmpEnums As StringArray
Dim oTmpEnumsX As StringArray
Dim lngTypeLoop As Long
Dim xNode As Node
Dim sKey As String
Dim sMainKey As String
Dim oTypes As StringArray

    On Error GoTo ErrHan:
    
    Set oTypes = New StringArray
    oTypes.DuplicatesAllowed = False
    oTypes.Sortable = True
    
    sMainKey = "Enumerators Declared in " & moVBPInfo.Title
    Set xNode = tv.Nodes.Add(, , sMainKey, "Enumerators", "info")
    
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
                
                Set oTmpEnums = oCodeInfo.EnumsStringArray
                lngTCount = oTmpEnums.Count
                
                If lngTCount > 0 Then
                    
                    For lngTypesLoop = 1 To lngTCount
                        
                        oTypes.AddItemString oTmpEnums(lngTypesLoop)
                        
                    Next lngTypesLoop
                    
                End If
                
                Set oCodeInfo = Nothing
                
            Next lngLoop
        
        End If
        
    Next lngFileLoop

    If Not oTypes Is Nothing Then
        
        lngTCount = oTypes.Count
        
        If lngTCount > 0 Then
            
            oTypes.Sort
            
            For lngTypesLoop = 1 To lngTCount
                
                Set oTmpEnums = oTypes.ItemAsStringArray(lngTypesLoop, ":")
                sKey = oTypes(lngTypesLoop)
                
                Set xNode = tv.Nodes.Add(sMainKey, tvwChild, sKey, oTmpEnums(1), "enum")
                xNode.Tag = cEnusNodeKey
                
                Set oTmpEnumsX = oTmpEnums.ItemAsStringArray(2, ";")
                
                For lngTypeLoop = 1 To oTmpEnumsX.Count
                    lngLIIndex = lngLIIndex + 1
                    Set xNode = tv.Nodes.Add(sKey, tvwChild, , oTmpEnumsX(lngTypeLoop), "enum")
                Next lngTypeLoop
                
            Next lngTypesLoop
            
                    
        End If
        
    End If

ResumeError:

    If Not oTmpA Is Nothing Then
        Set oTmpA = Nothing
    End If
    
    If Not oTmpEnums Is Nothing Then
        Set oTmpEnums = Nothing
    End If
    
    If Not oTmpEnumsX Is Nothing Then
        Set oTmpEnumsX = Nothing
    End If
    
    If Not oCodeInfo Is Nothing Then
        Set oCodeInfo = Nothing
    End If
        
    If tv.Nodes.Count Then
        tv.Nodes(1).Expanded = True
        tv_NodeClick tv.Nodes(1)
    End If
    
    lblItemCount.Caption = "Item Count: " & Format$(lngTCount, cNumFormat)
    mItemCount = lngTCount ' lngLIIndex
    
'    If lngLIIndex > lngTCount Then
'        lblItemCount.Caption = lblItemCount.Caption & " unique, in total: " & Format$(lngLIIndex, cNumFormat)
'        picSB.ToolTipText = "There are more declarations but some are duplicated"
'    Else
'        picSB.ToolTipText = ""
'    End If
    
    lngLIIndex = 0&
    sTmp = vbNullString
    sMainKey = vbNullString
    
Exit Sub

ErrHan:

    Debug.Print "frmEnumsReport.pGenerateEnumsReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub


Private Sub pUpdateText()
    
Dim xNode As Node
Dim sKey As String
Dim stext As String
Dim sString As String

    If tv.Nodes.Count = 0 Then Exit Sub
    
    Set xNode = tv.SelectedItem
    
    If Not xNode Is Nothing Then
            
        If xNode.Tag = cEnusNodeKey Then
                        
            sKey = xNode.Key
            
            stext = modGeneral.EnumDeclaration(sKey)
            sString = vbNewLine & stext & vbNewLine & vbNewLine
                                
            stext = modGeneral.SelectCaseFromEnum(sKey)
            sString = sString & stext & vbNewLine & vbNewLine
            
            stext = modGeneral.TranslateEnumFunction(sKey)
            sString = sString & stext & vbNewLine & vbNewLine
            
            sString = modEncode.BuildRTFString(sString)
            
            txt.TextRTF = sString
            
        End If
    
        Set xNode = Nothing
        
        sKey = vbNullString
        stext = vbNullString
        sString = vbNullString
        
    End If
    
End Sub

Private Sub tv_NodeClick(ByVal Node As ComctlLib.Node)

Dim sTmp As String
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    If Not Node Is Nothing Then
        
        If Not Node.Parent Is Nothing Then
            If Node.Parent.Index <> 1 Then
                sTmp = Node.Parent.Text & "."
            End If
        End If
        
        If Len(Node.Key) Then
            If Len(sTmp) Then
                sTmp = sTmp & Node.Key
            Else
                sTmp = Node.Key
            End If
            lblAPIDec.Caption = sTmp
        Else
            If Len(sTmp) Then sTmp = sTmp & Node.Text
            lblAPIDec.Caption = sTmp
        End If
        
        pUpdateText
        
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub imgSplitMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mMouseDown = True
    picSplitMain.Visible = True
End Sub

Private Sub imgSplitMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim xcd As CoDim, xmm As MinMax
    If mMouseDown Then
        xcd.Left = imgSplitMain.Left + x
        xmm.Min = (2 * cSplitLimit)
        xmm.max = ScaleWidth - (2 * cSplitLimit)
        If xcd.Left > xmm.Min And xcd.Left < xmm.max Then
            picSplitMain.Move xcd.Left
        End If
    End If
End Sub

Private Sub imgSplitMain_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mMouseDown = False
    picSplitMain.Visible = False
    pResize
End Sub

Private Sub pResize()
    
Dim lngHeight As Long

    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    imgSplitMain.Left = picSplitMain.Left
    lngHeight = ScaleHeight - picMain.Height - picSB.Height
    
    With tv
        .Left = 120
        .Width = picSplitMain.Left - .Left
        .Height = lngHeight
        .Top = picMain.Height
        imgSplitMain.Height = .Height
        imgSplitMain.Top = .Top
        picSplitMain.Height = .Height
        picSplitMain.Top = .Top
    End With
    
    With txt
        .Left = picSplitMain.Left + picSplitMain.Width
        .Top = picMain.Height
        .Width = ScaleWidth - .Left - (2 * cHBorder)
        .Height = lngHeight
    End With
            
    lngHeight = 0&
            
End Sub


