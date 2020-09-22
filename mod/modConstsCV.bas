Attribute VB_Name = "modConsts"
Attribute VB_Description = "mod. for program const declarations"
Option Explicit

Public Const cReleaseName As String = "VBCB"
Public Const cReleaseVersion As String = "R8"

' -------------------------------------------------------------------
' ... constants for tree views.
' -------------------------------------------------------------------
' v 5
Public Const cMissingFileSig As String = "cvMissingFile"
' -------------------------------------------------------------------

Public Const cFileSig As String = "cvFile"
Public Const cGUIDSig As String = "cvGUID"
Public Const cMembSig As String = "Member"
Public Const cParaSig As String = "Parameter"

Public Const cProjNodeKey As String = "Project"
Public Const cMainNodeKey As String = "Main"
Public Const cDecsNodeKey As String = "Declarations"
Public Const cSubsNodeKey As String = "Subs"
Public Const cFuncNodeKey As String = "Functions"
Public Const cPropNodeKey As String = "Properties"
Public Const cHeadNodeKey As String = "Header"

Public Const cEvesNodeKey As String = "Events"
Public Const cConsNodeKey As String = "Constants"
Public Const cTypsNodeKey As String = "Types"
Public Const cEnusNodeKey As String = "Enums"
Public Const cAPIsNodeKey As String = "APIs"
Public Const cAPINodeKey As String = "API"
Public Const cImpsNodeKey As String = "Implements"

Public Const cFormNodeKey As String = "Forms "
Public Const cClasNodeKey As String = "Classes "
Public Const cModsNodeKey As String = "Modules "
Public Const cUCtlNodeKey As String = "User Controls "
Public Const cDesgNodeKey As String = "Designers "
Public Const cRefsNodeKey As String = "References "
Public Const cCompNodeKey As String = "Components "
Public Const cPrpPNodeKey As String = "Property Pages "
Public Const cInfoNodeKey As String = "Info"
Public Const cExeNodeKey As String = "Exe"
Public Const cRDocNodeKey As String = "RelatedDocs"
' -------------------------------------------------------------------
Public Const cNumFormat As String = "#,##0"
' -------------------------------------------------------------------
' ... accessor related fields.
Public Const c_len_Public As Long = 6
Public Const c_len_Private As Long = 7
Public Const c_len_Friend As Long = 6

Public Const c_word_Public As String = "Public"
Public Const c_word_Private As String = "Private"
Public Const c_word_Friend As String = "Friend"
' -------------------------------------------------------------------
Public Const c_word_End As String = "End"
' -------------------------------------------------------------------
' ... method type related fields.
Public Const c_len_Sub As Long = 3
Public Const c_len_Function As Long = 8
Public Const c_len_Property As Long = 8

Public Const c_word_Sub As String = "Sub"
Public Const c_word_Function As String = "Function"
Public Const c_word_Property As String = "Property"
' -------------------------------------------------------------------
Public Const c_Word_Const As String = "Const "
Public Const c_len_Const As Long = 6
' -------------------------------------------------------------------
Public Const cFormEvents As String = " Form_Activate Form_Click Form_DblClick Form_Deactivate " & _
"Form_DragDrop Form_DragOver Form_GotFocus Form_Initialize Form_KeyDown Form_KeyPress " & _
"Form_KeyUp Form_LinkClose Form_LinkError Form_LinkOpen Form_Load Form_LostFocus " & _
"Form_MouseDown Form_MouseMove Form_MouseUp Form_OLECompleteDrag Form_OLEDragDrop " & _
"Form_OLEDragOver Form_OLEGiveFeedback Form_OLESetData Form_OLEStartDrag Form_Paint " & _
"Form_QueryUnload Form_Resize Form_Terminate Form_Unload "

Public Const cMDIFormEvents As String = " MDIForm_Activate MDIForm_Click MDIForm_DblClick MDIForm_Deactivate " & _
"MDIForm_DragDrop MDIForm_DragOver MDIForm_Initialize " & _
"MDIForm_LinkClose MDIForm_LinkError MDIForm_LinkExecute MDIForm_LinkOpen MDIForm_Load " & _
"MDIForm_MouseDown MDIForm_MouseMove MDIForm_MouseUp MDIForm_OLECompleteDrag MDIForm_OLEDragDrop " & _
"MDIForm_OLEDragOver MDIForm_OLEGiveFeedback MDIForm_OLESetData MDIForm_OLEStartDrag " & _
"MDIForm_QueryUnload MDIForm_Resize MDIForm_Terminate MDIForm_Unload "

Public Const cUCEvents As String = " UserControl_AccessKeyPress UserControl_AmbientChanged " & _
"UserControl_AsyncReadComplete UserControl_AsyncReadProgress UserControl_Click " & _
"UserControl_DblClick UserControl_DragDrop UserControl_DragOver UserControl_EnterFocus " & _
"UserControl_ExitFocus UserControl_GetDataMember UserControl_GotFocus UserControl_Hide " & _
"UserControl_HitTest UserControl_Initialize UserControl_InitProperties UserControl_KeyDown UserControl_KeyPress " & _
"UserControl_KeyUp UserControl_LostFocus UserControl_MouseDown UserControl_MouseMove " & _
"UserControl_MouseUp UserControl_OLECompleteDrag UserControl_OLEDragDrop " & _
"UserControl_OLEDragOver UserControl_OLEGiveFeedback UserControl_OLESetData " & _
"UserControl_OLEStartDrag UserControl_Paint UserControl_ReadProperties " & _
"UserControl_Resize UserControl_Show UserControl_Terminate UserControl_WriteProperties "

Public Const cClassEvents As String = " Class_GetDataMember Class_Initialize Class_Terminate "

Public Const cDlgCancelErr As Long = 32755

' -------------------------------------------------------------------
Public Const c_def_UnZipFolder As String = "UnZipFiles" ' ... default value for property UnZipFolder.

Public Function AppTitle() As String
Attribute AppTitle.VB_Description = "Returns the Name and Version of the program."

' ... v8, not really a constant strictly speaking
' ...     but acts like one.

    AppTitle = cReleaseName & " " & cReleaseVersion

End Function

Public Function AppVersion() As String
Attribute AppVersion.VB_Description = "Returns the Release Version for the program."

    AppVersion = cReleaseVersion
    
End Function
