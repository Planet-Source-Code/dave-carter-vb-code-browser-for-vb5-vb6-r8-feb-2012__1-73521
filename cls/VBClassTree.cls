VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBClassTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A simple Class to manage populating a Tree View control with information found in a Class/Code FIle that was read by a CodeInfo class."
' what?
'  a class dedicated to loading code file / class methods
'  into a tree view by digesting a loaded CodeInfo Class instance.
' why?
'  main reason is function re-use via small client code footprint
'  more than this, it provides an easy root to extensibility so that
'  other things can be included.
' when?
'  when a form / user control requires the Class Explorer Tree View.
' how?
'  'dimension and create a new instance:
'   Dim oCETree As VBClassTree
'       Set oCETree = New VBClassTree
'
'  'call the Init method.
'       oCETree.Init oCodeInfo, TreeView1, ImageList1, boolIncludeHeadCount
'
'  'release / dispose
'       Set OCETree = Nothing
'
'  'the boolIncludeHeadCount parameter instructs the class to include / omit
'  'the count of items in parent nodes.

' who?
'  d.c.

' dependencies?
'   ... comctl32.ocx / mscomctl.ocx: either will do, preference for first for xp themes.
'   ... modStringArrays.bas
'   ... modStrings.bas
'   ... modGeneral.bas
'   ... StringArray.cls
'   ... CodeInfo.cls
'
' requirements?
'  requires a Tree View and Image List pair.
'  the Image List must contain all the images required (currently 25 of them).
'  an Initialised CodeInfo instance.
'
' Note:
'   ... I have avoided using treeview events e.g. private withevents motv as treeview
'   ... in order to keep compatibility with vb5 where the tree view events are not fired within the class,
'   ... else this class could be a fully encapsulated function.
'   ... This means that the client has to take some responsibilty in response to
'   ... user activity regarding node selection.

Option Explicit

Private moTV As TreeView
Private moCodeInfo As CodeInfo
Private moImageList As ImageList    ' ... provides access to the ImageList passed.

Private mTVIcons As Boolean
Private mHaveImageList As Boolean
Private mInitialised As Boolean

Private mIncludeHeadCount As Boolean ' ... private field for property IncludeHeadCount.

'Private Const cFormEvents As String = " Form_Activate Form_Click Form_DblClick Form_Deactivate " & _
'"Form_DragDrop Form_DragOver Form_GotFocus Form_Initialize Form_KeyDown Form_KeyPress " & _
'"Form_KeyUp Form_LinkClose Form_LinkError Form_LinkOpen Form_Load Form_LostFocus " & _
'"Form_MouseDown Form_MouseMove Form_MouseUp Form_OLECompleteDrag Form_OLEDragDrop " & _
'"Form_OLEDragOver Form_OLEGiveFeedback Form_OLESetData Form_OLEStartDrag Form_Paint " & _
'"Form_QueryUnload Form_Resize Form_Terminate Form_Unload "
'
'Private Const cMDIFormEvents As String = " MDIForm_Activate MDIForm_Click MDIForm_DblClick MDIForm_Deactivate " & _
'"MDIForm_DragDrop MDIForm_DragOver MDIForm_Initialize " & _
'"MDIForm_LinkClose MDIForm_LinkError MDIForm_LinkExecute MDIForm_LinkOpen MDIForm_Load " & _
'"MDIForm_MouseDown MDIForm_MouseMove MDIForm_MouseUp MDIForm_OLECompleteDrag MDIForm_OLEDragDrop " & _
'"MDIForm_OLEDragOver MDIForm_OLEGiveFeedback MDIForm_OLESetData MDIForm_OLEStartDrag " & _
'"MDIForm_QueryUnload MDIForm_Resize MDIForm_Terminate MDIForm_Unload "
'
'Private Const cUCEvents As String = " UserControl_AccessKeyPress UserControl_AmbientChanged " & _
'"UserControl_AsyncReadComplete UserControl_AsyncReadProgress UserControl_Click " & _
'"UserControl_DblClick UserControl_DragDrop UserControl_DragOver UserControl_EnterFocus " & _
'"UserControl_ExitFocus UserControl_GetDataMember UserControl_GotFocus UserControl_Hide " & _
'"UserControl_HitTest UserControl_Initialize UserControl_InitProperties UserControl_KeyDown UserControl_KeyPress " & _
'"UserControl_KeyUp UserControl_LostFocus UserControl_MouseDown UserControl_MouseMove " & _
'"UserControl_MouseUp UserControl_OLECompleteDrag UserControl_OLEDragDrop " & _
'"UserControl_OLEDragOver UserControl_OLEGiveFeedback UserControl_OLESetData " & _
'"UserControl_OLEStartDrag UserControl_Paint UserControl_ReadProperties " & _
'"UserControl_Resize UserControl_Show UserControl_Terminate UserControl_WriteProperties "
'
'Private Const cClassEvents As String = " Class_GetDataMember Class_Initialize Class_Terminate "

Private Sub Class_Terminate()
    
    Clear
    
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the current data and tree view nodes ahead of re-using or terminating."

' Sub:             Clear
' Description:     Clears the current data and tree view nodes ahead of re-using or terminating.

    On Error GoTo ErrHan:
    
    mInitialised = False
    mHaveImageList = False
    mTVIcons = False
    
    If Not moTV Is Nothing Then Set moTV = Nothing
    If Not moImageList Is Nothing Then Set moImageList = Nothing
    If Not moCodeInfo Is Nothing Then Set moCodeInfo = Nothing
    
Exit Sub
ErrHan:

    Debug.Print "VBClassTree.Clear", Err.Number, Err.Description
    Resume Next

End Sub ' ... Clear.

Public Property Get IncludeHeadCount() As Boolean
Attribute IncludeHeadCount.VB_Description = "Describes whether no. of sub items is incuded in parent node text, True = yes, False = no."

    Let IncludeHeadCount = mIncludeHeadCount

End Property ' ... IncludeHeadCount: Boolean.

Public Property Let IncludeHeadCount(ByVal pNewValue As Boolean)

    Let mIncludeHeadCount = pNewValue

End Property ' ... IncludeHeadCount: Boolean.

Public Sub Init(ByRef pCodeInfo As CodeInfo, _
                ByRef pTreeview As TreeView, _
       Optional ByRef pImageList As ImageList, _
       Optional ByVal pIncHeadCount As Boolean = True, _
       Optional ByRef pOK As Boolean = False, _
       Optional ByRef pErrMsg As String = vbNullString)
Attribute Init.VB_Description = "Main entry point, instruction to load data in codeinfo object into a tree view."

' Sub:             Init
' Description:     Initialise the Class Explorer with CodeInfo instance and Tree View.

    On Error GoTo ErrHan:

    pOK = Not pCodeInfo Is Nothing
    If pOK = False Then
        Err.Raise vbObjectError + 1000, , "Code Info object not instanced."
    Else
        pOK = pCodeInfo.Initialised
        If pOK = False Then
            Err.Raise vbObjectError + 1000, , "Code Info object not initialised."
        End If
    End If
    
    pOK = Not pTreeview Is Nothing
    If pOK = False Then
        Err.Raise vbObjectError + 1000, , "Class Tree View not instanced."
    End If
    ' -------------------------------------------------------------------
    ' ... if this is reused then reset module fields and objects.
    If mInitialised Then Clear
    ' -------------------------------------------------------------------
    ' ... set up the module code info and tree view objects.
    Set moCodeInfo = pCodeInfo
    Set moTV = pTreeview
    ' -------------------------------------------------------------------
    ' ... empty treeview.
    If moTV.Nodes.Count > 0 Then
        moTV.Nodes(1).Selected = True
        moTV.Nodes.Clear
    End If
    ' -------------------------------------------------------------------
    ' ... set up tree view image list.
    If Not pImageList Is Nothing Then
        Set moImageList = pImageList
        Set moTV.ImageList = moImageList
        mHaveImageList = True
    End If
    mTVIcons = Not moTV.ImageList Is Nothing
    ' -------------------------------------------------------------------
    ' ... include count of subitems in parent node text?
    mIncludeHeadCount = pIncHeadCount
    ' -------------------------------------------------------------------
    ' ... load the code info stuff into the tree view.
    mInitialised = True
    
    pLoadNodes
    ' -------------------------------------------------------------------
'    If pOK = False Then
'        Err.Raise vbobjectError + 1000, , "Unable to load Class Explorer Nodes:" & vbNewLine & pErrMsg
'    End If
    
ErrResume:

'    mInitialised = pOK

Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "VBClassTree.Init", Err.Number, Err.Description
    Resume ErrResume:


End Sub ' ... Init.

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a boolean describing whether the tree view has been set up etc."
    
    Initialised = mInitialised
    
End Property

Private Function pAddNode(Optional ByVal pParent As String = vbNullString, _
                          Optional ByVal pRelation As TreeRelationshipConstants = tvwChild, _
                          Optional ByVal pKey As String = vbNullString, _
                          Optional ByVal pText As String = vbNullString, _
                          Optional ByVal pIcon As Long = 0, _
                          Optional ByVal pTag As String = vbNullString, _
                          Optional ByRef pOK As Boolean = False, _
                          Optional ByRef pErrMsg As String = vbNullString) As Node
Attribute pAddNode.VB_Description = "single method for adding a node to the tree view, includes image validation."
        
    On Error GoTo ErrHan:
    
    If mInitialised = True Then
        
        If Len(pParent) Then
            Set pAddNode = moTV.Nodes.Add(pParent, pRelation, pKey, pText)
        Else
            Set pAddNode = moTV.Nodes.Add(, pRelation, pKey, pText)
        End If
        pAddNode.Tag = pTag
        
        If pIcon > 0 Then
            If mTVIcons = True Then
                pAddNode.Image = pIcon
            End If
        End If
    
    Else
        
        Err.Raise vbObjectError + 1000, , "Unable to load Nodes if Class is not Initialised."
    
    End If
    
ResumeError:

Exit Function

ErrHan:
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "VBClassTree.pAddNode.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Private Function plImageIndexByKey(Optional ByVal pKey As String = vbNullString) As Long
Attribute plImageIndexByKey.VB_Description = "retrieves the index of an image in an image list via its key."

Dim lngIcon As Long

' ... Note:
'   ... ImageList First Image Index is 1.
'   ... pAddNode checks the value of the icon index, if < 1 then ignores
    
    On Error GoTo ErrHan:

    If mHaveImageList Then
        If moImageList.ListImages.Count > 0 Then
            lngIcon = moImageList.ListImages(pKey).Index
        End If
    End If

ResumeError:

    plImageIndexByKey = lngIcon

Exit Function

ErrHan:
    lngIcon = -1
    Debug.Print "VBClassTree.plImageIndexByKey.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function

Private Sub pLoadDecNodes(ByVal Node As ComctlLib.Node)
Attribute pLoadDecNodes.VB_Description = "method to load the declarations nodes e.g. constants, enums, types & apis."
' ... process class explorer node click.
Dim lngKey As Long
Dim stext As String
Dim sParams() As String
Dim sParamString As String
Dim i As Long
Dim sKey As String
Dim j As Long
Dim sMembers() As String
Dim sMember As String
Dim lngFound As Long
Dim sAccessor As String
Dim lngAccessor As Long
Dim saTmp As StringArray
Dim xSA As SArray
Dim lngIcon As Long
Dim lngPIcon As Long
Dim tAPIInfo As APIInfo

Dim tVarInfo As VariableInfo ' ... v7/8

Dim sAPIText As String
Dim sTmp As String
Dim oNode As Node
Dim lF2 As Long
Dim lF1 As Long
Dim lngCount As Long

' -------------------------------------------------------------------
' ... Used to load the declarations node and its children
    
    On Error GoTo ErrHan:
    
    If moCodeInfo Is Nothing Then Err.Raise vbObjectError + 1000, , "Class Explorer Code Info Class not instanced."
    If moCodeInfo.Initialised = False Then Err.Raise vbObjectError + 1000, , "Class Explorer Info Class not Initialised."
    ' -------------------------------------------------------------------
    
    
    lngPIcon = 2
        
        
    ' ... this proceeds to add to the declarations node, stuff like
    ' ... constants, events, apis etc...
        
'    moCodeInfo.Declarations   ' ... required to initialise reading of declarations section.
        
    ' ... process Events.
    sParamString = moCodeInfo.EventsString
    If Len(sParamString) Then
        
        ' ... events definitions are delimited by vbCrLf.
        modStringArrays.SplitString sParamString, sParams, vbCrLf, lngCount
        If lngCount > 1 Then ' ... no sorting if one item or no items.
            modStringArrays.QuickSortOnStringArray sParams, 0, lngCount - 1
        End If
        lngIcon = 3
        
        stext = cEvesNodeKey
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        
        Set oNode = pAddNode(cDecsNodeKey, tvwChild, cEvesNodeKey, stext, lngIcon, cEvesNodeKey)
        
        For i = 0 To lngCount - 1
            
            sKey = sParams(i)
            stext = sParams(i)
            
            ' ... bit of a hack to extract the event name for the node text.
            lF2 = InStr(1, stext, "(")
            If lF2 > 1 Then
                lF1 = modStrings.InStrRevChar(stext, " ", lF2)
                If lF1 > 0 Then
                    stext = Mid$(stext, lF1 + 1, lF2 - lF1 - 1)
                End If
            End If
            
            Set oNode = pAddNode(cEvesNodeKey, tvwChild, sKey, stext, lngIcon, cEvesNodeKey)
            
            pLoadParamNodes sKey, sKey
        
        Next i
    
    End If
    ' -------------------------------------------------------------------
    ' ... process Variables.
    sParamString = moCodeInfo.VarsString ' 17 & 18
    If Len(sParamString) Then
        ' -------------------------------------------------------------------
        ' ... get params array and then sort it.
        modStringArrays.SplitString sParamString, sParams, vbCrLf, lngCount
        If lngCount > 1 Then ' ... no sorting if one item or no items.
            modStringArrays.QuickSortOnStringArray sParams, 0, lngCount - 1
        End If
        ' -------------------------------------------------------------------
        stext = "Variables"
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        
        If lngCount Then
            Set oNode = pAddNode(cDecsNodeKey, tvwChild, "Variables", stext, 1, "Variables")
        End If
        For i = 0 To lngCount - 1
            stext = sParams(i)
            ParseVariableInfoItem stext, tVarInfo
            stext = tVarInfo.Name & ": " & tVarInfo.Type
            lngIcon = 17
            If tVarInfo.Accessor = 2 Then lngIcon = 20
'            If tVarInfo.ObjectWithEvents Then lngIcon = 22
            pAddNode "Variables", tvwChild, stext, stext & IIf(tVarInfo.ObjectWithEvents, " [WithEvents]", ""), lngIcon, "Variable#" & CStr(tVarInfo.EditorLineStart)
        Next i
    End If
    
    ' -------------------------------------------------------------------
    ' ... process Constants.
    sParamString = moCodeInfo.ConstsString
    If Len(sParamString) Then
        sParamString = moCodeInfo.ConstsStringv9
        modStringArrays.SplitString sParamString, sParams, vbCrLf, lngCount
        If lngCount > 1 Then ' ... no sorting if one item or no items.
            modStringArrays.QuickSortOnStringArray sParams, 0, lngCount - 1
        End If
        lngIcon = 4
        
        stext = cConsNodeKey
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        
        Set oNode = pAddNode(cDecsNodeKey, tvwChild, cConsNodeKey, stext, lngIcon, cConsNodeKey)
        lngIcon = lngIcon - 1
        
        For i = 0 To lngCount - 1
            
            sKey = sParams(i)
            stext = sParams(i)
            
            sAccessor = vbNullString: lngAccessor = 1
            If Left$(stext, c_len_Private) = c_word_Private Then
                lngAccessor = 2
            ElseIf Left$(stext, c_len_Public) = c_word_Public Then
                lngAccessor = 1
            End If
            lngFound = InStr(1, stext, c_Word_Const)
            If lngFound > 0 Then
                stext = Mid$(stext, lngFound + c_len_Const)
            End If
            
            lngFound = InStr(1, stext, "|")
            If lngFound > 0 Then
                stext = Left$(stext, lngFound - 1)
            End If
            modStrings.ReplaceChars stext, " As", ":"
            
            Set oNode = pAddNode(cConsNodeKey, tvwChild, sKey, stext, lngIcon + lngAccessor, cConsNodeKey)
            
        Next i
    
    End If
        
    ' -------------------------------------------------------------------
    ' ... process Types.
    sParamString = moCodeInfo.TypesString
    If Len(sParamString) Then
        
        ' ... Types are defined as single lines with Name : [members} 1 ; 2 ; 3 ... vbCrLf.
        modStringArrays.SplitString sParamString, sParams, vbCrLf, lngCount
        If lngCount > 1 Then ' ... no sorting if one item or no items.
            modStringArrays.QuickSortOnStringArray sParams, 0, lngCount - 1
        End If
        lngIcon = 6
        stext = cTypsNodeKey
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        
        Set oNode = pAddNode(cDecsNodeKey, tvwChild, cTypsNodeKey, stext, lngIcon, cTypsNodeKey)
        
        For i = 0 To lngCount - 1 ' UBound(sParams)
            
            stext = sParams(i)
            lngKey = InStr(1, stext, ":")
            
            If lngKey > 0 Then
                
                sKey = Left$(stext, lngKey - 1)
                
                sKey = modStrings.LeftOfComment(sKey)
                                    
                Set oNode = pAddNode(cTypsNodeKey, tvwChild, sKey, sKey, lngIcon, stext)
                
                stext = Mid$(stext, lngKey + 1)
                
                If Len(stext) Then
                    
                    modStringArrays.SplitString stext, sMembers, ";"
                    
                    For j = 0 To UBound(sMembers)
                        
                        sMember = sMembers(j)
                        sTmp = modStrings.Replace(sMember, " As", ":")
                        
                        sTmp = modStrings.LeftOfComment(sTmp)
                        
                        Set oNode = pAddNode(sKey, tvwChild, sKey & sMember, sTmp, lngIcon)
                            
                    Next j
                
                End If
                
            End If
            
        Next i
    End If
        
    ' -------------------------------------------------------------------
    ' ... process Enumerators.
    sParamString = moCodeInfo.EnumsString
    
    If Len(sParamString) Then
    
        ' ... enums are formatted as types e.g. Name : [members} 1 ; 2 ; 3 ... vbCrLf
        modStringArrays.SplitString sParamString, sParams, vbCrLf, lngCount
        If lngCount > 1 Then ' ... no sorting if one item or no items.
            modStringArrays.QuickSortOnStringArray sParams, 0, lngCount - 1
        End If
        lngIcon = 7
        stext = cEnusNodeKey
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        
        Set oNode = moTV.Nodes.Add(cDecsNodeKey, tvwChild, cEnusNodeKey, stext)
        If mTVIcons Then oNode.Image = lngIcon
        
        For i = 0 To UBound(sParams)
            
            stext = sParams(i)
            lngKey = InStr(1, stext, ":")
            
            If lngKey > 0 Then
                
                sKey = stext
                
                Set oNode = pAddNode(cEnusNodeKey, tvwChild, sKey, Left$(stext, lngKey - 1), lngIcon)
                stext = Mid$(stext, lngKey + 1)
                
                If Len(stext) Then
                
                    modStringArrays.SplitString stext, sMembers, ";"
                    
                    For j = 0 To UBound(sMembers)
                        
                        sMember = sMembers(j)
                        sTmp = modStrings.Replace(sMember, " =", ":")
                        
                        sTmp = modStrings.LeftOfComment(sTmp)
                        
                        Set oNode = pAddNode(sKey, tvwChild, sKey & sMember, sTmp, lngIcon)
                        
                    Next j
                
                End If
                
            End If
            
        Next i
    End If
        
    ' -------------------------------------------------------------------
    ' ... process API Declarations.
    Set saTmp = moCodeInfo.APIStringArray
    
    lngCount = saTmp.Count
    
    If lngCount > 0 Then
        
        lngIcon = 8
        stext = cAPIsNodeKey
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        
        ' ... add apis node.
        Set oNode = pAddNode(cDecsNodeKey, tvwChild, cAPIsNodeKey, stext, lngIcon)
        oNode.Tag = cAPIsNodeKey
        
        For i = 1 To lngCount
            
            ' ... convert API Declarative String into APIInfo Type.
            modGeneral.ParseAPIInfoItem saTmp, i, tAPIInfo

            sMember = saTmp(i)
            sAPIText = tAPIInfo.Name
            If Len(tAPIInfo.ReturnValue) > 0 Then
                sAPIText = sAPIText & ": " & tAPIInfo.ReturnValue
            End If
            
            ' ... note secondary use of tag.
            Set oNode = pAddNode(cAPIsNodeKey, tvwChild, sMember, sAPIText, lngIcon + tAPIInfo.Scope - 1, tAPIInfo.Declaration)
            oNode.Tag = cAPINodeKey
            ' ... add library name.
            Set oNode = pAddNode(sMember, tvwChild, , "Library: " & tAPIInfo.Lib, 21)
            
            If Len(tAPIInfo.Alias) Then
                ' ... add alias name.
                Set oNode = pAddNode(sMember, tvwChild, , "Alias: " & tAPIInfo.Alias, 21)
            End If
            
            ' ... load parameters.
            pLoadParamNodes sMember, sMember
        
        Next i
    
    End If
        
    ' -------------------------------------------------------------------
    ' ... process Implements Statements.
    Set saTmp = moCodeInfo.ImplementsStringArray
    
    lngCount = saTmp.Count
    
    If lngCount > 0 Then
    
        lngIcon = 10
        stext = cImpsNodeKey
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(lngCount, cNumFormat)
        End If
        Set oNode = pAddNode(cDecsNodeKey, tvwChild, cImpsNodeKey, stext, lngIcon)
        
        For i = 1 To lngCount
            sMember = saTmp(i)
            Set oNode = pAddNode(cImpsNodeKey, tvwChild, sMember, sMember, lngIcon, cImpsNodeKey)
        Next i
    
    End If
    
    If Not xSA Is Nothing Then
        Set xSA = Nothing
    End If
        
        ' -------------------------------------------------------------------
        
ResumeError:
    
    ' -------------------------------------------------------------------
Exit Sub
ErrHan:
    Debug.Print "VBClassTree.pLoadDecNodes.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
'    Resume
End Sub

Private Sub pLoadNodes()
Attribute pLoadNodes.VB_Description = "internal method to load tree view."

Dim lngLoop As Long
Dim lngMemberCount As Long
Dim lngIcon As Long
Dim lngPIcon As Long
Dim oNode As Node
Dim sKey As String
Dim sParentKey As String
Dim sTag As String
Dim stext As String
Dim q As QuickMemberInfo
Dim lngNodes As Long
Dim oMArray As StringArray
Dim otmpArray As StringArray
Dim lngIndex As Long
Dim lngType As Long
Dim lngAccessor As Long
Dim sVType As String
Dim sHeader As String
Dim bIsForm As Boolean ' v6, add form events to form events node under subs.
Dim bIsUC As Boolean
Dim bIsClass As Boolean
Dim tMemberInfo As MemberInfo ' ... v8
Dim bHasMenu As Boolean
Dim sMenuMethods As String

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    If moCodeInfo Is Nothing Then Err.Raise vbObjectError + 1000, , "Code Reader not available."
    If moTV Is Nothing Then Err.Raise vbObjectError + 1000, , "No Class Explorer TreeView available."
    If moCodeInfo.Initialised = False Then Err.Raise vbObjectError + 1000, , "Code Reader not Initialised."
    ' -------------------------------------------------------------------
    lngPIcon = 2    ' ... the declarations / parameters icon.
    ' -------------------------------------------------------------------
    lngMemberCount = moCodeInfo.MemberCount
    sHeader = Trim$(moCodeInfo.GetHeader)
    ' -------------------------------------------------------------------
    ' ... top node; need to get the image from image list if we got one.
    stext = moCodeInfo.Name
    If mIncludeHeadCount Then
        stext = stext & " " & Format$(lngMemberCount, cNumFormat)
    End If
    sKey = cMainNodeKey: sTag = cMainNodeKey
    ' ... get the icon from the image list if available.
    lngIcon = plImageIndexByKey(moCodeInfo.MemberType)
    Set oNode = pAddNode("", tvwFirst, sKey, stext, lngIcon, sTag)
    ' -------------------------------------------------------------------
    ' ... set the Parent Node's Key to add on to.
    sParentKey = cMainNodeKey
        ' v8/9, add info node to capture report -----------------------------
'    lngPIcon = 21
'    Set oNode = pAddNode(sParentKey, tvwChild, cInfoNodeKey, "Report", 21, cInfoNodeKey)

    ' -------------------------------------------------------------------
        bIsForm = moCodeInfo.MemberType = "Form"
        If bIsForm Then
            ' ... add a node for form events even if we don't get any.
            pAddNode cMainNodeKey, tvwChild, "FormEvents", "Form", 26
        Else
            bIsUC = moCodeInfo.MemberType = "User Control"
            If bIsUC Then
                ' ... add a node for form events even if we don't get any.
                pAddNode cMainNodeKey, tvwChild, "UCEvents", "User Control", 26
            
            Else
                bIsClass = moCodeInfo.MemberType = "Class"
                If bIsClass Then
                    ' ... add a node for form events even if we don't get any.
                    pAddNode cMainNodeKey, tvwChild, "ClassEvents", "Class", 26
                End If
            End If
        End If
        
    If moCodeInfo.HasMenu Then
        bHasMenu = True
        sMenuMethods = moCodeInfo.MenuMethods
        Set oNode = pAddNode(sParentKey, tvwChild, "MenuItems", "Menu", 26, "MenuItems")
    End If
    ' ... header.
    Set oNode = pAddNode(sParentKey, tvwChild, cHeadNodeKey, cHeadNodeKey, lngPIcon, cHeadNodeKey)
    ' -------------------------------------------------------------------
    ' ... declarations.
    lngIcon = lngPIcon
    stext = cDecsNodeKey
    sKey = cDecsNodeKey: sTag = cDecsNodeKey
    Set oNode = pAddNode(sParentKey, tvwChild, sKey, stext, lngIcon, sTag)
    pLoadDecNodes oNode
    If oNode.Children Then oNode.Expanded = True
    ' -------------------------------------------------------------------
    ' ... subs node.
    lngMemberCount = moCodeInfo.CountOfSubs
    If lngMemberCount > 0 Then
        pLoadParentMethodNode lngMemberCount, cSubsNodeKey, 11
        lngNodes = lngNodes + 2
    End If
    ' -------------------------------------------------------------------
    ' ... functions node.
    lngMemberCount = moCodeInfo.CountOfFunctions
    If lngMemberCount > 0 Then
        pLoadParentMethodNode lngMemberCount, cFuncNodeKey, 14
        lngNodes = lngNodes + 4
    End If
    ' -------------------------------------------------------------------
    ' ... properties node.
    lngMemberCount = moCodeInfo.CountOfProperties
    If lngMemberCount > 0 Then
        pLoadParentMethodNode lngMemberCount, cPropNodeKey, 17
        lngNodes = lngNodes + 8
    End If
    ' -------------------------------------------------------------------
    ' ... member nodes (available as a stringarray from codeinfo object).
    Set oMArray = moCodeInfo.MembersStringArray
    lngMemberCount = oMArray.Count
    
    
    For lngLoop = 1 To lngMemberCount
    
        ' ... each item in stringarray is a delimited string describing Member Name, Index, Type and Accessor.
        ' ... convert this into a stringarray to access individual members.
        ' ... see CodeInfo.pParseClass for building this string.
        ' -------------------------------------------------------------------
        ParseMemberInfoItem oMArray, lngLoop, tMemberInfo
        ' -------------------------------------------------------------------
        Set otmpArray = oMArray.ItemAsStringArray(lngLoop, "|")
        ' -------------------------------------------------------------------
        stext = otmpArray(1)
        lngIndex = otmpArray.ItemAsNumberValue(2)
        lngType = otmpArray.ItemAsNumberValue(3)
        lngAccessor = otmpArray.ItemAsNumberValue(4)
        ' -------------------------------------------------------------------
        sParentKey = vbNullString
        Select Case lngType
            Case 1: sParentKey = cSubsNodeKey
                ' ... v6, added form events to Subs \ Form.
                If bIsForm Then
                    If Left$(stext, 5) = "Form_" Then
                        If InStr(1, cFormEvents, " " & stext & " ") Then
                            sParentKey = "FormEvents"
                            stext = Mid$(stext, 6) ' ... lose the Form_ prefix.
                        End If
                    ElseIf Left$(stext, 8) = "MDIForm_" Then
                        If InStr(1, cMDIFormEvents, " " & stext & " ") Then
                            sParentKey = "FormEvents"
                            stext = Mid$(stext, 9)
                        End If
                        If moTV.Nodes("FormEvents").Text <> "MDIForm" Then
                            moTV.Nodes("FormEvents").Text = "MDIForm"
                        End If
                    End If
                Else
                    If bIsUC Then
                        If Left$(stext, 12) = "UserControl_" Then
                            If InStr(1, cUCEvents, " " & stext & " ") Then
                                sParentKey = "UCEvents"
                                stext = Mid$(stext, 13) ' ... lose the UserControl_ prefix.
                            End If
                        End If
                    Else
                        If bIsClass Then
                            If Left$(stext, 6) = "Class_" Then
                                If InStr(1, cClassEvents, " " & stext & " ") Then
                                    sParentKey = "ClassEvents"
                                    stext = Mid$(stext, 7) ' ... lose the Class_ prefix.
                                End If
                            End If
                        End If
                    End If
                End If
                If (bIsForm Or bIsUC) And bHasMenu Then
                    If InStr(1, sMenuMethods, " " & stext & " ") Then
                        sParentKey = "MenuItems"
                        stext = Left$(stext, Len(stext) - 6)
                    End If
                End If
                
            Case 2: sParentKey = cFuncNodeKey
            Case 3: sParentKey = cPropNodeKey
        End Select
        ' -------------------------------------------------------------------
        If Len(sParentKey) Then
            lngIcon = 11
            ' ... lngType Values: 1 = Sub, 2 = Function, 3 = Property.
            If lngType > 1 Then
                sVType = otmpArray(5)
                If Len(sVType) Then
                    stext = stext & ": " & sVType ' ... add return / value type on to end of name.
                End If
                lngIcon = lngIcon + 3
                If lngType > 2 Then
                    lngIcon = lngIcon + 3
                End If
            End If
            
            lngIcon = lngIcon + lngAccessor - 1
            sKey = CStr(lngIndex) & "x"
            sTag = cMembSig
            
            Set oNode = pAddNode(sParentKey, tvwChild, sKey, stext, lngIcon, sTag)
            
            ' -------------------------------------------------------------------
            ' ... method / member parameters.
            ' ... access the quick member item directly.
            q = moCodeInfo.QuickMember(lngIndex)
            
            pLoadParamNodes sKey, q.Declaration
                        
        End If
        
    Next lngLoop

    If moTV.Nodes.Count > 0 Then
        ' -------------------------------------------------------------------
        If 2 And lngNodes Then moTV.Nodes(cSubsNodeKey).Expanded = True
        If 4 And lngNodes Then moTV.Nodes(cFuncNodeKey).Expanded = True
        If 8 And lngNodes Then moTV.Nodes(cPropNodeKey).Expanded = True
        ' -------------------------------------------------------------------
        moTV.Nodes(cMainNodeKey).Expanded = True
        moTV.Nodes(cMainNodeKey).Selected = True
        moTV.Nodes(cMainNodeKey).EnsureVisible
        ' -------------------------------------------------------------------
    End If

ResumeError:
    
    sHeader = vbNullString
    
Exit Sub
ErrHan:
    Debug.Print "VBClassTree.pLoadNodes.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    Resume
End Sub

Private Sub pLoadParamNodes(pParentKey As String, _
                            pTheDeclaration As String)
Attribute pLoadParamNodes.VB_Description = "method to extract parameters from a declaraion string and add them as nodes to the owner method/member."
                            
Dim lngPCount As Long
Dim sParams() As String
Dim sParamString As String
Dim lngLoop As Long
Dim sMember As String
Dim lngIcon As Long
Dim oNode As Node

' -------------------------------------------------------------------
' ... Helper:   Load parameter nodes off a parent node.
' ... Note:     Presumes that moCodeInfo is instanced, a valid node key is passed and an imagelist is associated with the Members TreeView.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    sParamString = moCodeInfo.ExtractParamsString(pTheDeclaration)
    moCodeInfo.GetDecParams sParamString, sParams, lngPCount
    
    If lngPCount > 0 Then
        
        For lngLoop = LBound(sParams) To UBound(sParams)
            
            sMember = sParams(lngLoop)
            lngIcon = 2
            
            If InStr(sMember, "ByVal ") > 0 Then
                lngIcon = 20
            End If
            
            Set oNode = pAddNode(pParentKey, tvwChild, , sMember, lngIcon, cParaSig)
            
        Next lngLoop
    
    End If

Exit Sub
ErrHan:
    Debug.Print "VBClassTree.pLoadParamNodes.Error: " & Err.Number & "; " & Err.Description
End Sub

Private Sub pLoadParentMethodNode(ByVal pCount As Long, _
                                  ByVal pParent As String, _
                                  ByVal pIcon As Long)
Attribute pLoadParentMethodNode.VB_Description = "loads main Subs, Functions & Properties nodes to the tree view."
                                  
' ... adds either the Subs, Functions or Properties Parent Node.
' ... only purpose is to reduce footprint within main node loading sub (currently pLoadNodes)
' ... and eliminate duplicate code.

Dim stext As String
    
    If pCount > 0 Then
        stext = pParent
        If mIncludeHeadCount Then
            stext = stext & " " & Format$(pCount, cNumFormat)
        End If
        pAddNode cMainNodeKey, tvwChild, pParent, stext, pIcon, pParent
    End If

End Sub

