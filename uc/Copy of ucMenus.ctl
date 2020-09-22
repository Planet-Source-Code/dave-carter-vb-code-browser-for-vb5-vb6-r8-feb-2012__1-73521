VERSION 5.00
Begin VB.UserControl ucMenus 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1125
   ScaleWidth      =   825
   Begin VB.Menu mnuProject 
      Caption         =   "Project"
      Begin VB.Menu mnuPLoadProject 
         Caption         =   "Load Project"
      End
      Begin VB.Menu mnuPSep1a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPOpen 
         Caption         =   "Open"
         Begin VB.Menu mnuPNewWindow 
            Caption         =   "in New Window"
         End
         Begin VB.Menu mnuPOpenFolder 
            Caption         =   "Containing Folder"
         End
         Begin VB.Menu mnuPNotePad 
            Caption         =   "in Text Editor"
         End
         Begin VB.Menu mnuPIDE2 
            Caption         =   "VB5"
         End
         Begin VB.Menu mnuPIDE 
            Caption         =   "VB6"
         End
      End
      Begin VB.Menu mnuPSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPCopyProject 
         Caption         =   "Copy Project"
      End
      Begin VB.Menu mnuPCopyFile 
         Caption         =   "Copy File"
      End
      Begin VB.Menu mnuPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPCreateManifest 
         Caption         =   "Create Compiled Manifest Resource"
      End
      Begin VB.Menu mnuPSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPCompile 
         Caption         =   "Compile"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPReports 
         Caption         =   "Text Reports"
         Begin VB.Menu mnuPQuickReport 
            Caption         =   "Quick Report"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPFullReport 
            Caption         =   "Project Report"
         End
         Begin VB.Menu mnuPAPIReport 
            Caption         =   "API Report (Actual)"
         End
         Begin VB.Menu mnuPAPIReportDistinct 
            Caption         =   "API Report (Distinct)"
         End
      End
      Begin VB.Menu mnuTypesEtc 
         Caption         =   "Members"
         Begin VB.Menu mnuPDevHelp 
            Caption         =   "Methods"
         End
         Begin VB.Menu mnuPSepM1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPAPI 
            Caption         =   "API"
         End
         Begin VB.Menu mnuPConstant 
            Caption         =   "Constant"
         End
         Begin VB.Menu mnuPType 
            Caption         =   "Type"
         End
         Begin VB.Menu mnuPEnum 
            Caption         =   "Enum"
         End
         Begin VB.Menu mnuPSepM2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPMemberPages 
            Caption         =   "HTML Member Pages"
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPFindDeps 
         Caption         =   "Find Dependencies"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPSep1d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPSearch 
         Caption         =   "Search Project"
      End
      Begin VB.Menu mnuPSep1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuPSep1c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "Class"
      Begin VB.Menu mnuCCopySig 
         Caption         =   "Copy Signature"
      End
      Begin VB.Menu mnuCCopyMethod 
         Caption         =   "Copy Method"
      End
      Begin VB.Menu mnuCReports 
         Caption         =   "Reports"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCFullReport 
         Caption         =   "Full Report"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCRefresh 
         Caption         =   "Refresh"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuViewer 
      Caption         =   "Viewer"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuVShow 
         Caption         =   "Show / Hide"
         Begin VB.Menu mnuVSep1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuVProjExp 
            Caption         =   "  Project Explorer"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuVClassExp 
            Caption         =   "  Class Explorer"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuVToolBar 
            Caption         =   "  ToolBar"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuVStatusBar 
            Caption         =   "  Status Bar"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuSep1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyntaxColuring 
         Caption         =   "Syntax Colouring"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVInterface 
         Caption         =   "Interface"
         Begin VB.Menu mnuVSep3 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuVIntPubOnly 
            Caption         =   "  Public Members Only"
         End
         Begin VB.Menu mnuVIntAllMembs 
            Caption         =   "  All Members"
         End
      End
      Begin VB.Menu mnuRSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCQuickReport 
         Caption         =   "Quick Reporter"
      End
      Begin VB.Menu mnuRSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRHist 
         Caption         =   "History"
         Enabled         =   0   'False
         Begin VB.Menu mnuRHItem 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "ucMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "a user control to provide a simple pop-up menu for the browser form."

' Purpose:  Provide PopUp Menus to viewer form.
'           (preference to keep mdi menu in tact
'           and this seems to be the cheapest route to that.)
'           A very basic implementation, none of the nice to have
'           or fancy bits, RAD Style!.
' who?
'  d.c.

Option Explicit
' P   ShowTextEditor: Boolean.  ' ... Sets whether the Open with Text Editor option on the project menu is enabled.
' P   ShowVB5IDE: Boolean.  ' ... Sets whether the Open with VB5 option on the project menu is enabled.
' P   ShowVB6IDE: Boolean.  ' ... Sets whether the Open with VB6 option on the project menu is enabled.

' ... event passes menu caption, group menu and menu item within group.
Public Event MenuItemClick(Caption As String, Menu As Long, Item As Long)
Attribute MenuItemClick.VB_Description = "Event Raised when an item on a menu is clicked."

' ... menu groups.
Private Const mPMenu As Long = 100  ' ... project menu.
Private Const mCMenu As Long = 200  ' ... class menu.
Private Const mVMenu As Long = 300  ' ... viewer menu.

Private moBackMenu As StringArray   ' ... a string array for the viewer > history menu items.

Private m_ShowTextEditor As Boolean ' ... private field for property ShowTextEditor.
Private m_ShowVB5IDE As Boolean ' ... private field for property ShowVB5IDE.
Private m_ShowVB6IDE As Boolean ' ... private field for property ShowVB6IDE.

Public Property Let ShowVB6IDE(ByVal pNewValue As Boolean)
Attribute ShowVB6IDE.VB_Description = "Sets whether the Open with VB6 option on the project menu is enabled."


    m_ShowVB6IDE = pNewValue
    mnuPIDE.Visible = pNewValue

End Property ' ... ShowVB6IDE: Boolean.


Public Property Let ShowVB5IDE(ByVal pNewValue As Boolean)
Attribute ShowVB5IDE.VB_Description = "Sets whether the Open with VB5 option on the project menu is enabled."

    m_ShowVB5IDE = pNewValue
    mnuPIDE2.Visible = pNewValue

End Property ' ... ShowVB5IDE: Boolean.


Public Property Let ShowTextEditor(ByVal pNewValue As Boolean)
Attribute ShowTextEditor.VB_Description = "Sets whether the Open with Text Editor option on the project menu is enabled."

    m_ShowTextEditor = pNewValue
    mnuPNotePad.Visible = pNewValue

End Property ' ... ShowTextEditor: Boolean.

Private Sub pGenerateHistoryMenu()
' ... set up the history items for the viewer pop-up menu.
Dim lngCount As Long
Dim lngLoop As Long
Dim sTmpA As StringArray
Dim lngMCount As Long

    On Error GoTo ErrHan:
    
    mnuRHist.Enabled = False
    mnuRHItem(0).Caption = ""
        
    ' ... release current history menu items.
    lngMCount = mnuRHItem.Count
    If lngMCount > 0 Then
        For lngLoop = lngMCount - 1 To 1 Step -1
            Unload mnuRHItem(lngLoop)
        Next lngLoop
    End If
    
    If Not moBackMenu Is Nothing Then
        
        lngCount = moBackMenu.Count
        mnuRHist.Enabled = lngCount > 0
        ' ... expand available history menu items.
        If lngCount > 1 Then
            For lngLoop = 1 To lngCount - 1
                Load mnuRHItem(lngLoop)
            Next lngLoop
        End If
        ' ... set captions for history menu items.
        For lngLoop = 0 To lngCount - 1
            Set sTmpA = moBackMenu.ItemAsStringArray(lngLoop + 1, Chr$(0))
            mnuRHItem(lngLoop).Caption = sTmpA(2)
            mnuRHItem(lngLoop).Tag = sTmpA(1)
        Next lngLoop
    
    End If

Exit Sub
ErrHan:
    Debug.Print "ucMenus.pGenerateHistoryMenu.Error: " & Err.Description
    Err.Clear
    Resume Next
End Sub

Private Sub mnuCCopyMethod_Click()
    RaiseEvent MenuItemClick(mnuCCopyMethod.Caption, mCMenu, 2)
End Sub

Private Sub mnuCCopySig_Click()
    RaiseEvent MenuItemClick(mnuCCopySig.Caption, mCMenu, 1)
End Sub

Private Sub mnuCFullReport_Click()
    RaiseEvent MenuItemClick(mnuCFullReport.Caption, mCMenu, 22)
End Sub

Private Sub mnuCQuickReport_Click()
    RaiseEvent MenuItemClick(mnuCQuickReport.Caption, mCMenu, 21)
End Sub

Private Sub mnuCRefresh_Click()
    RaiseEvent MenuItemClick(mnuCRefresh.Caption, mCMenu, 3)
End Sub

Private Sub mnuPAPI_Click()
    RaiseEvent MenuItemClick(mnuPAPI.Caption, mPMenu, 36)
End Sub

Private Sub mnuPAPIReport_Click()
    RaiseEvent MenuItemClick(mnuPAPIReport.Caption, mPMenu, 23)
End Sub

Private Sub mnuPAPIReportDistinct_Click()
    RaiseEvent MenuItemClick(mnuPAPIReportDistinct.Caption, mPMenu, 24)
End Sub

Private Sub mnuPClose_Click()
    RaiseEvent MenuItemClick(mnuPClose.Caption, mPMenu, 102)
End Sub

Private Sub mnuPCompile_Click()
    RaiseEvent MenuItemClick(mnuPCompile.Caption, mPMenu, 3)
End Sub

Private Sub mnuPConstant_Click()
    RaiseEvent MenuItemClick(mnuPConstant.Caption, mPMenu, 37)
End Sub

Private Sub mnuPCopyFile_Click()
    RaiseEvent MenuItemClick(mnuPCopyFile.Caption, mPMenu, 31)
End Sub

Private Sub mnuPCopyProject_Click()
    RaiseEvent MenuItemClick(mnuPCopyProject.Caption, mPMenu, 30)
End Sub

Private Sub mnuPCreateManifest_Click()
    RaiseEvent MenuItemClick(mnuPCreateManifest.Caption, mPMenu, 80)
End Sub

Private Sub mnuPDevHelp_Click()
    RaiseEvent MenuItemClick(mnuPDevHelp.Caption, mPMenu, 60)
End Sub

Private Sub mnuPEnum_Click()
    RaiseEvent MenuItemClick(mnuPEnum.Caption, mPMenu, 38)
End Sub

Private Sub mnuPFindDeps_Click()
    RaiseEvent MenuItemClick(mnuPFindDeps.Caption, mPMenu, 4)
End Sub

Private Sub mnuPFullReport_Click()
    RaiseEvent MenuItemClick(mnuPFullReport.Caption, mPMenu, 22)
End Sub

Private Sub mnuPIDE_Click()
    RaiseEvent MenuItemClick(mnuPIDE.Caption, mPMenu, 12)
End Sub

Private Sub mnuPIDE2_Click()
    RaiseEvent MenuItemClick(mnuPIDE2.Caption, mPMenu, 14)
End Sub

Private Sub mnuPLoadProject_Click()
    RaiseEvent MenuItemClick(mnuPLoadProject.Caption, mPMenu, 101)
End Sub

Private Sub mnuPMemberPages_Click()
    RaiseEvent MenuItemClick(mnuPMemberPages.Caption, mPMenu, 70)
End Sub

Private Sub mnuPNewWindow_Click()
    RaiseEvent MenuItemClick(mnuPNewWindow.Caption, mPMenu, 10)
End Sub

Private Sub mnuPNotePad_Click()
    RaiseEvent MenuItemClick(mnuPNotePad.Caption, mPMenu, 11)
End Sub

Private Sub mnuPOpenFolder_Click()
    RaiseEvent MenuItemClick(mnuPOpenFolder.Caption, mPMenu, 13)
End Sub

Private Sub mnuPQuickReport_Click()
    RaiseEvent MenuItemClick(mnuPQuickReport.Caption, mPMenu, 21)
End Sub

Private Sub mnuPRefresh_Click()
    RaiseEvent MenuItemClick(mnuPRefresh.Caption, mPMenu, 6)
End Sub

Private Sub mnuPSearch_Click()
    RaiseEvent MenuItemClick(mnuPSearch.Caption, mPMenu, 5)
End Sub

Private Sub mnuPType_Click()
    RaiseEvent MenuItemClick(mnuPType.Caption, mPMenu, 39)
End Sub

Private Sub mnuRHItem_Click(Index As Integer)
    RaiseEvent MenuItemClick(mnuRHItem(Index).Caption, mVMenu, 100 + CLng(Str$(mnuRHItem(Index).Tag)))
End Sub

Private Sub mnuSyntaxColuring_Click()
    RaiseEvent MenuItemClick(mnuSyntaxColuring.Caption, mVMenu, 33)
End Sub

Private Sub mnuVClassExp_Click()
    RaiseEvent MenuItemClick(mnuVClassExp.Caption, mVMenu, 2)
End Sub

Private Sub mnuVIntAllMembs_Click()
    RaiseEvent MenuItemClick(mnuVIntAllMembs.Caption, mVMenu, 6)
End Sub

Private Sub mnuVIntPubOnly_Click()
    RaiseEvent MenuItemClick(mnuVIntPubOnly.Caption, mVMenu, 5)
End Sub

Private Sub mnuVProjExp_Click()
    RaiseEvent MenuItemClick(mnuVProjExp.Caption, mVMenu, 1)
End Sub

Private Sub mnuVStatusBar_Click()
    RaiseEvent MenuItemClick(mnuVStatusBar.Caption, mVMenu, 4)
End Sub

Private Sub mnuVToolBar_Click()
    RaiseEvent MenuItemClick(mnuVToolBar.Caption, mVMenu, 3)
End Sub

Public Sub ShowClassMenu()
Attribute ShowClassMenu.VB_Description = "Shows the Class Explorer Pop-Up Menu."
' ... show class menu.
    UserControl.PopupMenu mnuClass

End Sub

Public Sub ShowProjectMenu()
Attribute ShowProjectMenu.VB_Description = "Shows the Project Explorer Pop-Up Menu."
' ... show project menu.
    UserControl.PopupMenu mnuProject

End Sub

Public Sub ShowViewerMenu(Optional pHistoryArray As StringArray = Nothing, _
                          Optional pWithSyntaxColours As Boolean = True, _
                          Optional pShowProjectExp As Boolean = True, _
                          Optional pShowClassExp As Boolean = True, _
                          Optional pShowToolBar As Boolean = True, _
                          Optional pShowStatusBar As Boolean = True)
' ... show viewer menu.
    Set moBackMenu = pHistoryArray
    pGenerateHistoryMenu
    ' -------------------------------------------------------------------
    ' v6
    ' ... added check marks and new with syntax colours...
    
    mnuSyntaxColuring.Checked = pWithSyntaxColours
    mnuVProjExp.Checked = pShowProjectExp
    mnuVClassExp.Checked = pShowClassExp
    mnuVToolBar.Checked = pShowToolBar
    mnuVStatusBar.Checked = pShowStatusBar
    ' -------------------------------------------------------------------
    
    UserControl.PopupMenu mnuViewer
    If Not moBackMenu Is Nothing Then Set moBackMenu = Nothing
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' ... load resource file text for menu captions.
    On Error Resume Next    ' ... in case res file is lost
                            ' ... unfortunately result is no captions
    mnuCQuickReport.Caption = LoadResString(132)
    mnuPQuickReport.Caption = LoadResString(132)
    
End Sub

