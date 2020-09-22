VERSION 5.00
Begin VB.Form frmRefViewer 
   Caption         =   "Mini Reference Viewer"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   37
   Icon            =   "frmRefViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8940
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   8940
      TabIndex        =   13
      Top             =   5640
      Width           =   8940
      Begin VB.Label lblMember 
         Height          =   795
         Left            =   60
         TabIndex        =   15
         Top             =   300
         Width           =   8655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCoClass 
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
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   8655
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTB 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   8940
      TabIndex        =   4
      Top             =   0
      Width           =   8940
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   495
         Left            =   6660
         TabIndex        =   11
         Tag             =   "4096"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   180
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         ToolTipText     =   "Drag DLL or OCX file from Explorer to try loading into mini ref viewer"
         Top             =   180
         Width           =   7695
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "All Types"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Tag             =   "239"
         Top             =   660
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Constants"
         Height          =   375
         Index           =   1
         Left            =   1470
         TabIndex        =   9
         Tag             =   "4"
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Events"
         Height          =   375
         Index           =   2
         Left            =   4035
         TabIndex        =   8
         Tag             =   "2"
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Types"
         Height          =   375
         Index           =   3
         Left            =   5310
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "32"
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Unions"
         Height          =   375
         Index           =   4
         Left            =   6600
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "128"
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "CoClasses"
         Height          =   375
         Index           =   5
         Left            =   2745
         TabIndex        =   5
         Tag             =   "1"
         Top             =   660
         Width           =   1215
      End
   End
   Begin VB.PictureBox picSplitMain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2640
      ScaleHeight     =   1455
      ScaleWidth      =   75
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.ListBox List2 
      Height          =   4155
      Left            =   3420
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Label lblLib 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   6060
      Visible         =   0   'False
      Width           =   7695
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgSplitMain 
      Height          =   1425
      Left            =   2460
      MousePointer    =   9  'Size W E
      ToolTipText     =   "Resize Me"
      Top             =   1320
      Width           =   105
   End
End
Attribute VB_Name = "frmRefViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface for viewing methods and members of references added to a project.  The code comes from TlbInf32.chm."
Option Explicit

' The code in this form comes from TlbInf32 help file (last) available at:
' http://support.microsoft.com/kb/224331
' FILE: Tlbinf32.exe : Help Files for Tlbinf32.dll
' This article was previously published under Q224331

' All I have done is copy code from the help file into this form.
' Many original methods are not used but included all the same at the
' end of this source file, commented.

' The simple trick in use is the TlBInf32's ability to output type and member
' info directly into a list or combo box with the Get[...]InfoDirect methods :)


Enum BaseType
  btNone
  btIUnknown
  btIDispatch
End Enum

Private m_TLInf As TypeLibInfo

Private mMouseDown As Boolean
Private Const cHBorder As Long = 60
Private Const cSplitLimit As Long = 660

' -------------------------------------------------------------------

Public Sub LoadRef(pRefName As String)
    
    Text1 = pRefName
    Command1_Click
    
End Sub

Private Sub Form_Load()

  On Error Resume Next ' ... in case of delete setting being run and section not existing.

  Set m_TLInf = New TypeLibInfo
  m_TLInf.AppObjString = "<Unqualified>"
  
  ' -------------------------------------------------------------------
  ' ... run the following statement to clean registry
  ' ... remove the get and save settings to avoid using registry
'  DeleteSetting App.Title, Name
  
  LoadFormPosition Me, mdiMain.Height, mdiMain.Width
  
  picSplitMain.Left = CLng(Val(GetSetting(App.Title, Name, "SizeBarLeft", "2640")))
  
  pResize
  
  ClearMemory
  
End Sub

Private Sub Command1_Click()

Dim lngOptions As Long
Dim i As Long

  On Error Resume Next
  
  ' ... if minimised then normalise.
  If WindowState = 1 Then WindowState = 0
  
  lngOptions = 4096
    For i = 0 To optOptions.Count - 1
        If optOptions(i).Value = True Then
            lngOptions = CLng(Val(optOptions(i).Tag))
            Exit For
        End If
    Next i
  lblLib = ""
  lblCoClass = ""
  lblMember = ""
  List1.Clear
  List2.Clear
  m_TLInf.ContainingFile = Text1
  If Err Then Beep: MsgBox "Unable to read load Type Library Info.", vbExclamation, Caption: Exit Sub
  lblLib = m_TLInf
  With List1
    m_TLInf.GetTypesDirect .hwnd, , lngOptions ' tliStConstants  '+ tliStClasses '+ tliStDeclarations + tliStEvents
    If .ListCount Then .ListIndex = 0
  End With
  
End Sub

Private Sub Form_Resize()
    pResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If Me.WindowState <> 0 Then Exit Sub
    
    SaveFormPosition Me
    
    SaveSetting App.Title, Name, "SizeBarLeft", CLng(picSplitMain.Left)
    
    ClearMemory
    
End Sub

Private Sub List1_Click()
  lblMember = ""
  With List1
    List2.Clear
    'Retrieve the SearchData from the ItemData property
    m_TLInf.GetMembersDirect .ItemData(.ListIndex), List2.hwnd
    lblCoClass = lblLib & "." & List1.Text
    If List2.ListCount Then List2.ListIndex = 0
  End With
End Sub
Private Sub List2_Click()
Dim InvKinds As TLI.InvokeKinds
On Error Resume Next
lblMember = ""
    With List2
        InvKinds = .ItemData(.ListIndex)
        lblMember = PrototypeMember(m_TLInf, _
                                 List1.ItemData(List1.ListIndex), _
                                 InvKinds, , .[_Default])
    End With
End Sub

Private Function PrototypeMember( _
  TLInf As TypeLibInfo, _
  ByVal SearchData As Long, _
  ByVal InvokeKinds As InvokeKinds, _
  Optional ByVal MemberId As Long = -1, _
  Optional ByVal MemberName As String) As String
Dim pi As ParameterInfo
Dim fFirstParameter As Boolean
Dim fIsConstant As Boolean
Dim fByVal As Boolean
Dim retVal As String
Dim ConstVal As Variant
Dim strTypeName As String
Dim VarTypeCur As Integer
Dim fDefault As Boolean, fOptional As Boolean, fParamArray As Boolean
Dim TIType As TypeInfo
Dim TIResolved As TypeInfo
Dim TKind As TypeKinds
On Error Resume Next
  With TLInf
    fIsConstant = GetSearchType(SearchData) And tliStConstants
    With .GetMemberInfo(SearchData, InvokeKinds, MemberId, MemberName)
      If fIsConstant Then
        retVal = "Const "
      ElseIf InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then
        Select Case .ReturnType.VarType
          Case VT_VOID, VT_HRESULT
            retVal = "Sub "
          Case Else
            retVal = "Function "
        End Select
      Else
        retVal = "Property "
      End If
      retVal = retVal & .Name
      With .Parameters
        If .Count Then
          retVal = retVal & "("
          fFirstParameter = True
          fParamArray = .OptionalCount = -1
          For Each pi In .Me
            If Not fFirstParameter Then
              retVal = retVal & ", "
            End If
            fFirstParameter = False
            fDefault = pi.Default
            fOptional = fDefault Or pi.Optional
            If fOptional Then
              If fParamArray Then
                'This will be the only optional parameter
                retVal = retVal & "[ParamArray "
              Else
                retVal = retVal & "["
              End If
            End If
            With pi.VarTypeInfo
              Set TIType = Nothing
              Set TIResolved = Nothing
              TKind = TKIND_MAX
              VarTypeCur = .VarType
              If (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
              'If Not .TypeInfoNumber Then 'This may error, don't use here
                On Error Resume Next
                Set TIType = .TypeInfo
                If Not TIType Is Nothing Then
                  Set TIResolved = TIType
                  TKind = TIResolved.TypeKind
                  Do While TKind = TKIND_ALIAS
                    TKind = TKIND_MAX
                    Set TIResolved = TIResolved.ResolvedType
                    If Err Then
                      Err.Clear
                    Else
                      TKind = TIResolved.TypeKind
                    End If
                  Loop
                End If
                Select Case TKind
                  Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                    fByVal = .PointerLevel = 1
                  Case TKIND_RECORD
                    'Records not passed ByVal in VB
                    fByVal = False
                  Case Else
                    fByVal = .PointerLevel = 0
                End Select
                If fByVal Then retVal = retVal & "ByVal "
                retVal = retVal & pi.Name
                If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then retVal = retVal & "()"
                If TIType Is Nothing Then 'Error
                  retVal = retVal & " As ?"
                Else
                  If .IsExternalType Then
                    retVal = retVal & " As " & _
                             .TypeLibInfoExternal.Name & "." & TIType.Name
                  Else
                    retVal = retVal & " As " & TIType.Name
                  End If
                End If
                On Error GoTo 0
              Else
                If .PointerLevel = 0 Then retVal = retVal & "ByVal "
                retVal = retVal & pi.Name
                If VarTypeCur <> vbVariant Then
                  strTypeName = TypeName(.TypedVariant)
                  If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                    retVal = retVal & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                  Else
                    retVal = retVal & " As " & strTypeName
                  End If
                End If
              End If
              If fOptional Then
                If fDefault Then
                  retVal = retVal & ProduceDefaultValue(pi.DefaultValue, TIResolved)
                End If
                retVal = retVal & "]"
              End If
            End With
          Next
          retVal = retVal & ")"
        End If
      End With
      If fIsConstant Then
        ConstVal = .Value
        retVal = retVal & " = " & ConstVal
        Select Case VarType(ConstVal)
          Case vbInteger, vbLong
            If ConstVal < 0 Or ConstVal > 15 Then
              retVal = retVal & " (&H" & Hex$(ConstVal) & ")"
            End If
        End Select
      Else
        With .ReturnType
          VarTypeCur = .VarType
          If VarTypeCur = 0 Or (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
          'If Not .TypeInfoNumber Then 'This may error, don't use here
            On Error Resume Next
            If Not .TypeInfo Is Nothing Then
              If Err Then 'Information not available
                retVal = retVal & " As ?"
              Else
                If .IsExternalType Then
                  retVal = retVal & " As " & _
                           .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                Else
                  retVal = retVal & " As " & .TypeInfo.Name
                End If
              End If
            End If
            If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then retVal = retVal & "()"
            On Error GoTo 0
          Else
            Select Case VarTypeCur
              Case VT_VARIANT, VT_VOID, VT_HRESULT
              Case Else
                strTypeName = TypeName(.TypedVariant)
                If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                  retVal = retVal & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                Else
                  retVal = retVal & " As " & strTypeName
                End If
            End Select
          End If
        End With
      End If
      PrototypeMember = retVal & vbCrLf & "  " & _
                        "Member of " & TLInf.Name & "." & _
                        TLInf.GetTypeInfo(SearchData And &HFFFF&).Name & _
                        vbCrLf & "  " & .HelpString
    End With
  End With
End Function

'VB SearchData routines
Private Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes
  If SearchData And &H80000000 Then
    GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F&) Or &H80
  Else
    GetSearchType = SearchData \ &H1000000 And &HFF&
  End If
End Function

Private Function ProduceDefaultValue(DefVal As Variant, ByVal TI As TypeInfo) As String
Dim lTrackVal As Long
Dim MI ' As MemberInfo
Dim TKind As TypeKinds
On Error Resume Next
    If TI Is Nothing Then
        Select Case VarType(DefVal)
            Case vbString
                If Len(DefVal) Then
                    ProduceDefaultValue = """" & DefVal & """"
                End If
            Case vbBoolean 'Always show for Boolean
                ProduceDefaultValue = DefVal
            Case vbDate
                If DefVal Then
                    ProduceDefaultValue = "#" & DefVal & "#"
                End If
            Case Else 'Numeric Values
                If DefVal <> 0 Then
                    ProduceDefaultValue = DefVal
                End If
        End Select
    Else
        'See if we have an enum and track the matching member
        'If the type is an object, then there will never be a
        'default value other than Nothing
        TKind = TI.TypeKind
        Do While TKind = TKIND_ALIAS
            TKind = TKIND_MAX
            On Error Resume Next
            Set TI = TI.ResolvedType
            If Err = 0 Then TKind = TI.TypeKind
            On Error GoTo 0
        Loop
        If TI.TypeKind = TKIND_ENUM Then
            lTrackVal = DefVal
            For Each MI In TI.Members
                If MI.Value = lTrackVal Then
                    ProduceDefaultValue = MI.Name
                    Exit For
                End If
            Next
        End If
    End If
End Function

Private Sub optOptions_Click(Index As Integer)
    Command1_Click
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Data.GetFormat(1) Then
        Text1 = Data.GetData(vbCFText)
        Command1_Click
        ZOrder
    ElseIf Data.GetFormat(vbCFFiles) Then
        Text1 = Data.Files(1)
        Command1_Click
        ZOrder
    End If
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
    lngHeight = ScaleHeight - picTB.Height - picSB.Height
    With List1
        .Left = 120
        .Width = picSplitMain.Left - .Left
        .Height = lngHeight
        imgSplitMain.Height = .Height
        imgSplitMain.Top = .Top
        picSplitMain.Height = .Height
        picSplitMain.Top = .Top
    End With
    
    With List2
        .Left = picSplitMain.Left + picSplitMain.Width
        .Width = ScaleWidth - .Left - (2 * cHBorder)
        .Height = lngHeight
    End With
    
    lblMember.Width = ScaleWidth - 120
    Text1.Width = picTB.ScaleWidth - 360
    
End Sub

' -------------------------------------------------------------------
' ... methods found in tlb help but not used here as yet.

' -------------------------------------------------------------------

' ... (these constants are required in the declarations if their use is restored
' ... in any of the following methods).

'''Private Const strIID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
'''Private Const strIID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"

' -------------------------------------------------------------------

'''' ... NOTE:
'''' ... the following function is not valid in VB5 because it returns an array
'''' ... so I've made it a variant to help.
'''' ... originally it returned an array of TLI.MemberInfo()
'''
'''Private Function OrderedVTableFunctions( _
'''  ByVal TLInf As TypeLibInfo, _
'''  IFaceName As String, _
'''  BaseType As BaseType) As Variant ' TLI.MemberInfo()
'''Dim TI As InterfaceInfo
'''Dim TIStart As InterfaceInfo
'''Dim Bases As TypeInfos
'''Dim MI As MemberInfo
'''Dim iStandardEntries As Integer
'''Dim iProcessed As Integer
'''Dim fCheckDispatch As Boolean
'''Dim OVF() As MemberInfo
'''Dim MaxOffset As Integer
'''Dim CurOffset As Integer
'''  'Make sure we get a VTable interface before proceeding
'''  On Error Resume Next
'''  Set TI = TLInf.TypeInfos.NamedItem(IFaceName)
'''  Set TI = TI.VTableInterface
'''  If Err Or TI Is Nothing Then Exit Function
'''  On Error GoTo 0
'''  'To avoid retrieving the standard IDispatch and
'''  'IUnknown vtable entries on every call, first
'''  'walk the base interfaces to figure out if we
'''  'derived from IDispatch/IUnknown
'''  If TI.AttributeMask And TYPEFLAG_FDUAL Then
'''    BaseType = btIDispatch
'''    iStandardEntries = 7
'''    fCheckDispatch = True
'''  Else
'''    Set TIStart = TI
'''    BaseType = btNone
'''    Do Until TI Is Nothing
'''      iProcessed = iProcessed + 1
'''      Set Bases = TI.ImpliedInterfaces
'''      Set TI = Nothing
'''      If Bases.Count Then
'''        Set TI = Bases(1)
'''        If TI.Guid = strIID_IDispatch Then
'''          BaseType = btIDispatch
'''          Set TI = TIStart
'''          iStandardEntries = 7
'''          Exit Do
'''        ElseIf TI.Guid = strIID_IUnknown Then
'''          BaseType = btIUnknown
'''          iStandardEntries = 3
'''          Exit Do
'''        End If
'''      End If
'''    Loop
'''    Set TI = TIStart
'''  End If
'''  With TI.Members
'''    'Largest VTableOffset is generally on the last
'''    'member in this collection
'''    MaxOffset = .Item(.Count).VTableOffset
'''    ReDim OVF((MaxOffset \ 4) - iStandardEntries)
'''  End With
'''  Do Until TI Is Nothing
'''    'Walk each member
'''    For Each MI In TI.Members
'''      CurOffset = MI.VTableOffset
'''      If CurOffset > MaxOffset Then
'''        'This is extremely rare
'''        MaxOffset = CurOffset
'''        ReDim Preserve OVF((MaxOffset \ 4) - iStandardEntries)
'''      End If
'''      Set OVF((CurOffset \ 4) - iStandardEntries) = MI
'''    Next
'''    'Get the next base
'''    If fCheckDispatch Then
'''      Set Bases = TI.ImpliedInterfaces
'''      Set TI = Nothing
'''      If Bases.Count Then
'''        Set TI = Bases(1)
'''        If TI.Guid = strIID_IDispatch Then
'''          Exit Do
'''        End If
'''      End If
'''    Else
'''      iProcessed = iProcessed - 1
'''      If iProcessed = 0 Then Exit Do
'''      Set TI = TI.ImpliedInterfaces(1)
'''    End If
'''  Loop
'''  OrderedVTableFunctions = OVF
'''End Function


'''Private Function GetTypeInfoNumber(ByVal SearchData As Long) As Integer
'''  GetTypeInfoNumber = SearchData And &HFFF&
'''End Function
'''Private Function GetLibNum(ByVal SearchData As Long) As Integer
'''  SearchData = SearchData And &H7FFFFFFF
'''  GetLibNum = ((SearchData \ &H2000& And &H7) * &H100&) Or _
'''               (SearchData \ &H10000 And &HFF&)
'''End Function

'''Private Function GetHidden(ByVal SearchData As Long) As Boolean
'''    If SearchData And &H1000& Then GetHidden = True
'''End Function

'''Private Function BuildSearchData( _
'''   ByVal TypeInfoNumber As Integer, _
'''   ByVal SearchTypes As TliSearchTypes, _
'''   Optional ByVal LibNum As Integer, _
'''   Optional ByVal Hidden As Boolean = False) As Long
'''  If SearchTypes And &H80 Then
'''    BuildSearchData = _
'''      (TypeInfoNumber And &H1FFF&) Or _
'''      ((SearchTypes And &H7F) * &H1000000) Or &H80000000
'''  Else
'''    BuildSearchData = _
'''      (TypeInfoNumber And &H1FFF&) Or _
'''      (SearchTypes * &H1000000)
'''  End If
'''
'''  If LibNum Then
'''    BuildSearchData = BuildSearchData Or _
'''      ((LibNum And &HFF) * &H10000) Or _
'''      ((LibNum And &H700) * &H20&)
'''  End If
'''  If Hidden Then
'''    BuildSearchData = BuildSearchData Or &H1000&
'''  End If
'''End Function


