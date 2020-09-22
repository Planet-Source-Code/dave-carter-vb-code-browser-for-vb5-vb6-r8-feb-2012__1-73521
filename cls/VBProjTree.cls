VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBProjTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A simple Class to manage populating a Tree View control with information found in a VBP that was read by a VBPInfo class."

' what?
'  a class to handle the loading of a vbp into a tree view
'  and responding to user interaction.
' why?
'  to enable re-use of the function where-ever it may be required.
' when?
'  one wants to display info in a vbp in a tree view.
' how?
'  set up a new instance
'  e.g.
'       Dim oVBPTree As VBProjTree
'
'           Set oVBPTree = New VBProjTree
'
'   then call its Init method with a VBPInfo and Tree View.
'
'           ' ... assuming a tree view (tvProject) with an appropriate image list
'           ' ... exists along with an Initialised VBPInfo instance (oVBPInfo).
'
'           oVBPTree.Init oVBPInfo, TreeView1
'
' who?
'  d.c.

Option Explicit

' note: big problem in vb5, control events in a class don't appear to fire
'       as in vb6...
'       e.g.
'       Private WithEvents moTV As TreeView should mean that moTV's events
'       are fired here-in but it just doesn't happen.
'       I done found this out after writing this class which was meant to
'       handle such things, should be ok in vb6 though.


Private moTV As TreeView     ' Private WithEvents moTV As TreeView
Private moVBPInfo As VBPInfo
Private mWhichNodes As ProjectExplorerNodesEnum

Private moSelectedNode As Node

Private mHaveIMGList As Boolean
Private mInitialised As Boolean

Private mAttributeDelimiter As String

Private Const ctvChild As Long = 4
' v6
Private mIncHeadCount As Boolean

'' v 5
'Private Const cMissingForms As Long = 2
'Private Const cMissingMods As Long = 4
'Private Const cMissingClasses As Long = 8
'Private Const cMissingUCs As Long = 16
'Private mMissing As Long

'' v 5
'Private Sub pProcessMissing()
'
'Dim sTmpA As StringArray
'Dim lngCount As Long
'Dim lngLoop As Long
'Dim tFInfo As FileNameInfo
'Dim sTmp As String
'
'    mMissing = 0
'
'    If Not moVBPInfo Is Nothing Then
'        lngCount = moVBPInfo.MissingCount
'        If lngCount > 0 Then
'            Set sTmpA = moVBPInfo.MissingFiles
'            lngCount = sTmpA.Count  ' ... sanity ...
'            For lngLoop = 1 To lngCount
'                sTmp = sTmpA(lngLoop)
'                modFileName.ParseFileNameEx sTmp, tFInfo
'                Select Case LCase$(tFInfo.Extension)
'                    Case "frm":
'                        If CBool(mMissing And cMissingForms) = False Then mMissing = mMissing + cMissingForms
'                    Case "bas"
'                        If CBool(mMissing And cMissingMods) = False Then mMissing = mMissing + cMissingMods
'                    Case "cls"
'                        If CBool(mMissing And cMissingClasses) = False Then mMissing = mMissing + cMissingClasses
'                    Case "ctl"
'                        If CBool(mMissing And cMissingUCs) = False Then mMissing = mMissing + cMissingUCs
'                End Select
'            Next lngLoop
'        End If
'
'    End If
'
'End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the current data and tree view nodes ahead of re-using or terminating."

    On Error GoTo ErrHan:

    mInitialised = False
    mHaveIMGList = False
    mIncHeadCount = False
    mWhichNodes = 0&
    
    Set moSelectedNode = Nothing
    Set moTV = Nothing
    Set moVBPInfo = Nothing
    
Exit Sub
ErrHan:

    Debug.Print "VBProjTree.Clear", Err.Number, Err.Description
    Err.Clear
    Resume Next
    
End Sub ' ... Clear.

Public Sub Init(ByRef pVBPInfo As VBPInfo, _
                ByRef pTreeview As TreeView, _
       Optional ByVal pIncHeadCount As Boolean = False, _
       Optional ByVal pWhichNodes As ProjectExplorerNodesEnum = eAllNodes, _
       Optional ByRef pOK As Boolean = False, _
       Optional ByRef pErrMsg As String = vbNullString)

'... Parameters.
'    R__ pVBPInfo: VBPInfo           ' ... A VBPInfo instance loaded with data.
'    R__ pTreeview: TreeView ' ... A ComctlLib Tree View to display project info from VBPInfo instance.

Dim bOK As Boolean
Dim sErrMsg As String

    On Error GoTo ErrHan:
    
    mIncHeadCount = pIncHeadCount ' v6
    
    bOK = Not pVBPInfo Is Nothing
    If bOK = False Then
        Err.Raise vbObjectError + 1000, , "VBP Info object not instanced."
    Else
        bOK = pVBPInfo.Initialised
        If bOK = False Then
            Err.Raise vbObjectError + 1000, , "VBP Info object not initialised."
        End If
    End If
    
    bOK = Not pTreeview Is Nothing
    If bOK = False Then
        Err.Raise vbObjectError + 1000, , "Project Tree View not instanced."
    End If
    
    If mInitialised Then Clear
    
    mWhichNodes = pWhichNodes
    
    Set moVBPInfo = pVBPInfo
    Set moTV = pTreeview
    
    bOK = Not pTreeview.ImageList Is Nothing
    
    If moTV.Nodes.Count > 0 Then
        moTV.Nodes(1).Selected = True
        moTV.Nodes.Clear
    End If
    
    mHaveIMGList = bOK
    
    mAttributeDelimiter = moVBPInfo.AttributeDelimiter
    
'    ' -------------------------------------------------------------------
'    ' v 5
'    If moVBPInfo.MissingCount > 0 Then
'        pProcessMissing
'    End If
    
    pLoadVBPTree moVBPInfo.FileNameAndPath
    
    Let sErrMsg = vbNullString
    Let bOK = True

ErrResume:

    Let pErrMsg = sErrMsg
    Let pOK = bOK
    mInitialised = bOK
    
Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "VBProjTree.Init", Err.Number, Err.Description
    Resume ErrResume:


End Sub ' ... Init.

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a boolean describing whether the tree view has been set up etc."
    
    Initialised = mInitialised
    
End Property

Private Sub pLoadDesignerNodes(pSSArray As StringArray, pMainKey As String, pThisKey As String, pIcon As Integer)
Attribute pLoadDesignerNodes.VB_Description = "Add file item nodes to the tree view."

Dim lngLoop As Long
Dim tDataInfo As DataInfo
Dim bType As Byte
Dim iIconIndex As Integer
Dim oNode As Node
Dim stext As String
Dim sKey As String
Dim lngCount As Long
Dim sTag As String

'  within the items of the listed data (forms,classes...) is a delimited string
'  that follows the following format with a | [shift+\] delimiter:
' 1      2      3        4
' Name | Type | Exists | Location
'  where the listed items are objects or references the 4th item is its GUID.

'  regarding Type mentioned above...
' 0 Reference
' 1 Object
' 2 Form
' 3 Module
' 4 Class
' 5 User Control
' 6 Resource File
' 7 Data Environment
' 8 Data Report
' 9 Data Connection (inside dataenvironment).
' 10 Property Page ' v6.

' note: on forms
'  3 types of form described each prefixed with the form identifier 2;
' 2[1] mdi form         e.g. 21
' 2[2] normal form      e.g. 22
' 2[3] mdi child form   e.g. 23

    On Error GoTo ErrHan:
    
    If Not pSSArray Is Nothing Then
        lngCount = pSSArray.Count
        If lngCount > 0 Then
            stext = pThisKey ' & Str$(lngCount) ' ... v6 commented.
            If mIncHeadCount Then
                stext = stext & Str$(lngCount)  ' v6
            End If
            sKey = pThisKey
            ' ... group node e.g. Forms, Modules, Classes ...
            Set oNode = moTV.Nodes.Add(pMainKey, ctvChild, sKey, stext)
            If mHaveIMGList Then oNode.Image = pIcon
            oNode.Tag = sKey
            For lngLoop = 1 To lngCount
'                bType = -1
'                Debug.Print pSSArray(lngLoop)
                pParseVBPInfoItem pSSArray, lngLoop, tDataInfo
'                Debug.Print tDataInfo.ExtraInfo
                iIconIndex = 1              ' ... default, not found icon.
                sKey = tDataInfo.ExtraInfo  ' ... full filename and path.
                stext = tDataInfo.Name      ' ... name.
                bType = tDataInfo.Type
                If tDataInfo.Exists Then
                    ' ... set the icon to use.
                    iIconIndex = pIcon
                    If bType > 6 Then
                        Select Case bType
                            ' ... designers.
                            Case 7: iIconIndex = 13     ' ... data environment.
                            Case 8: iIconIndex = 12     ' ... data report.
                            ' ... forms
                            Case 21: iIconIndex = 5     ' ... mdi.
                            Case 22: iIconIndex = 7     ' ... normal.
                            Case 23: iIconIndex = 6     ' ... mdi child.
                        End Select
                    End If
                Else
                    ' -------------------------------------------------------------------
                    ' v5, expand group node to show missing files.
                    moTV.Nodes(pThisKey).Expanded = True
                End If
                
                Set oNode = moTV.Nodes.Add(pThisKey, ctvChild, sKey, stext)
                If mHaveIMGList Then oNode.Image = iIconIndex
                    ' -------------------------------------------------------------------
                    ' v5, indicate missing file.
                If tDataInfo.Exists = True Then
                    sTag = cFileSig
                Else
                    sTag = cMissingFileSig
                End If
                If bType < 2 Then
                    sTag = cGUIDSig
                End If
                oNode.Tag = sTag
                
            Next lngLoop
        End If
    End If

ResErr:
    
    On Error GoTo 0
    If Not oNode Is Nothing Then
        Set oNode = Nothing
    End If
    stext = vbNullString:   sKey = vbNullString:    sTag = vbNullString
    lngCount = 0:               lngLoop = 0
    iIconIndex = 0
    bType = 0
    
Exit Sub
ErrHan:
    Resume ResErr:
    
End Sub

Private Sub pLoadExeNode(pMainKey As String, pNodeTitle As String, pNodeText As String)
    
Dim oNode As Node

    If Len(pNodeText) > 0 Then
        Set oNode = moTV.Nodes.Add(cInfoNodeKey, tvwChild, cExeNodeKey, pNodeTitle & ": " & pNodeText)
        If mHaveIMGList Then oNode.Image = 15
        oNode.Tag = cExeNodeKey
    End If

End Sub

Private Sub pLoadInfoNode(pMainKey As String, pNodeTitle As String, pNodeText As String)
Attribute pLoadInfoNode.VB_Description = "Add a node to the Info nodes."
    
Dim oNode As Node

    If Len(pNodeText) > 0 Then
        Set oNode = moTV.Nodes.Add(cInfoNodeKey, tvwChild, , pNodeTitle & ": " & pNodeText)
        If mHaveIMGList Then oNode.Image = 15
        oNode.Tag = cInfoNodeKey
    End If

End Sub

Private Sub pLoadVBPTree(Optional pFileName As String = vbNullString)
Attribute pLoadVBPTree.VB_Description = "Method to manage the loading of the nodes into the tree view."

Dim i As Long
'Dim lngLoop As Long
Dim sMainKey As String
Dim sTmpData As StringArray
Dim tDataInfo As DataInfo
'Dim iIconIndex As Integer
'Dim bType As Byte
Dim sTmp As String
Dim oNode As Node
Dim xFileInfo As FileNameInfo
Dim sRelDocName As String

    On Error GoTo ErrHan:
    
    If moTV.Nodes.Count > 0 Then
        moTV.Nodes(1).Selected = True
        moTV.Nodes.Clear
    End If
    
    ' -------------------------------------------------------------------
    sMainKey = pFileName ' cFileSig & pFileName
    ' -------------------------------------------------------------------
    
    Set oNode = moTV.Nodes.Add(, , sMainKey, moVBPInfo.ProjectName)
    If mHaveIMGList Then oNode.Image = 2
    oNode.Tag = cProjNodeKey
    
    If mWhichNodes And eInfoNodes Then
        
        ' ... Info Nodes.
        
        Set oNode = moTV.Nodes.Add(sMainKey, tvwChild, cInfoNodeKey, moVBPInfo.ProjectType)
        If mHaveIMGList Then oNode.Image = 14
        
        sTmp = moVBPInfo.Title
        pLoadInfoNode sMainKey, "Title", sTmp
        
        sTmp = moVBPInfo.Description
        pLoadInfoNode sMainKey, "Description", sTmp
        
        sTmp = moVBPInfo.Version
        pLoadInfoNode sMainKey, "Version", sTmp
        
        ' ... v8
        If moVBPInfo.IsExe Then
            sTmp = moVBPInfo.ExeName32
            pLoadExeNode sMainKey, "Exe", sTmp
        End If
        
        sTmp = moVBPInfo.FileName
        pLoadInfoNode sMainKey, "File Name", sTmp
        
        sTmp = moVBPInfo.FilePath
        pLoadInfoNode sMainKey, "File Path", sTmp
        
        sTmp = moVBPInfo.HelpFile
        pLoadInfoNode sMainKey, "Help File", sTmp
        
        ' -------------------------------------------------------------------
        ' ... v6, added other vbp item info, note, if no value in sTmp then no node added!
        
        sTmp = moVBPInfo.Command32
        pLoadInfoNode sMainKey, "Command Line", sTmp
        
        sTmp = moVBPInfo.DllBaseAddress
        pLoadInfoNode sMainKey, "DLL Base Address", sTmp
    
        sTmp = moVBPInfo.CompatibleEXE32
        pLoadInfoNode sMainKey, "Binary Compatible file", sTmp
    
        sTmp = moVBPInfo.Path32
        pLoadInfoNode sMainKey, "Path 32", sTmp
    
        sTmp = moVBPInfo.VersionCompatible32
        pLoadInfoNode sMainKey, "Version Compatibilty", sTmp
    
        sTmp = moVBPInfo.ServerSupportFiles
        pLoadInfoNode sMainKey, "Server Support Files", sTmp
        
        ' -------------------------------------------------------------------
    
    End If
    
    If mWhichNodes And eObjectsNode Then
        pLoadDesignerNodes moVBPInfo.ReferencesData, sMainKey, cRefsNodeKey, 3
        pLoadDesignerNodes moVBPInfo.ComponentsData, sMainKey, cCompNodeKey, 4
    End If
    
    If mWhichNodes And eSourceFiles Then
        pLoadDesignerNodes moVBPInfo.FormsData, sMainKey, cFormNodeKey, 7
        pLoadDesignerNodes moVBPInfo.ClassesData, sMainKey, cClasNodeKey, 8
        pLoadDesignerNodes moVBPInfo.ModulesData, sMainKey, cModsNodeKey, 9
        pLoadDesignerNodes moVBPInfo.UserControlsData, sMainKey, cUCtlNodeKey, 10
        pLoadDesignerNodes moVBPInfo.PropertyPagesData, sMainKey, cPrpPNodeKey, 10
        pLoadDesignerNodes moVBPInfo.DesignersData, sMainKey, cDesgNodeKey, 12
    End If
    
    
    If moVBPInfo.DEConnectionsData.Count > 0 Then               ' ... designer Connections.
    ' ... not sure about this at moment having changed
    ' ... the designer nodes above.
'        Set sTmpData = moVBPInfo.DEConnectionsData
'        For i = 1 To sTmpData.Count
'            pParseVBPInfoItem sTmpData, i, tDataInfo
'            moTV.Nodes.Add cFileSig & tDataInfo.ExtraInfo, tvwChild, tDataInfo.ExtraInfo & "." & tDataInfo.Name, tDataInfo.Name, 13
'        Next i
    End If
    
    ' -------------------------------------------------------------------
    ' ... Related Docs
    If mWhichNodes And eRelatedDocuments Then
    
        Set sTmpData = moVBPInfo.RelatedDocs
        If Len(moVBPInfo.ResFileName) > 0 Or sTmpData.Count > 0 Then
            'Set oNode = moTV.Nodes.Add(sMainKey, tvwChild, cRDocNodeKey, "Related Docs" & IIf(mIncHeadCount, " " & sTmpData.Count, ""))
            Set oNode = moTV.Nodes.Add(sMainKey, tvwChild, cRDocNodeKey, "Related Docs" & IIf(mIncHeadCount, " " & sTmpData.Count + IIf(Len(moVBPInfo.ResFileName), 1, 0), ""))
    
            If mHaveIMGList Then oNode.Image = "folder"
        End If
        
        If Len(moVBPInfo.ResFileName) > 0 Then                      ' ... Resource File.
            
            sTmp = moVBPInfo.ResFileName
            modVB.ReadVBFilePath moVBPInfo.FilePath, sTmp   ' ... v5, adding more params to dummy parsevbpinfo... string below.
                                                            ' ... to make concurrent with its proper use.
                                                            ' ... e.g. adding "|0|1| & res file name to dummy string for data in tree view.
            Set sTmpData = New StringArray
            sTmpData.FromString moVBPInfo.ResFileName & "|0|1|" & sTmp, vbCr    ' ... one line expected so split on dummy char to get the one record.
            
            pParseVBPInfoItem sTmpData, 1, tDataInfo
            
            Set oNode = moTV.Nodes.Add(cRDocNodeKey, tvwChild, tDataInfo.ExtraInfo, tDataInfo.Name)
            If mHaveIMGList Then oNode.Image = AddIconToImageList(tDataInfo.ExtraInfo, moTV.ImageList, "res") ' oNode.Image = 11
            oNode.Tag = cFileSig
        
        End If
        
        Set sTmpData = moVBPInfo.RelatedDocs
        If sTmpData.Count > 0 Then
    '        Set oNode = moTV.Nodes.Add(sMainKey, tvwChild, cRDocNodeKey, "Related Docs" & IIf(mIncHeadCount, " " & sTmpData.Count, ""))
    '        If mHaveIMGList Then oNode.Image = "folder"
            For i = 1 To sTmpData.Count
                sTmp = sTmpData(i)
                modFileName.ParseFileNameEx sTmp, xFileInfo
                sRelDocName = xFileInfo.File
                Set oNode = moTV.Nodes.Add(cRDocNodeKey, tvwChild, sTmp, sRelDocName) 'sTmp)
                If mHaveIMGList Then
                    oNode.Image = AddIconToImageList(sTmpData(i), moTV.ImageList, "txt")
                    oNode.Tag = cFileSig
                End If
            Next i
        End If
    
    End If
ResumeError:
    
    If moTV.Nodes.Count > 0 Then
        moTV.Nodes(1).Expanded = True
        moTV.Nodes(1).Selected = True ' ... set as if top node selected.
    End If
    
    On Error GoTo 0
    
    If Not sTmpData Is Nothing Then Set sTmpData = Nothing
    If Not oNode Is Nothing Then Set oNode = Nothing
    
    sTmp = vbNullString
    sRelDocName = vbNullString
    sMainKey = vbNullString
    
Exit Sub
ErrHan:
    
    Debug.Print "frmMain.cmdOpenVBP_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

    Resume

End Sub ' ... pLoadVBPTree:

Private Sub pParseVBPInfoItem(pSArray As StringArray, pIndex As Long, pDataInfo As DataInfo)
Attribute pParseVBPInfoItem.VB_Description = "Shreds a VBPInfo Data item into a DataInfo structure."

    modGeneral.ParseDataInfoItem pSArray, pIndex, pDataInfo, mAttributeDelimiter

End Sub ' ... pParseVBPInfoItem:

Private Sub Class_Terminate()
    Clear
End Sub
