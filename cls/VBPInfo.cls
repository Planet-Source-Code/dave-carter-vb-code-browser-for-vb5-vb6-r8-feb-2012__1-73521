VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBPInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class to help read a Visual Basic (5/6) Project file (.vbp)."

' what?
'  reads a vbp providing access to its various attributes and objects.
' why?
'  to help with presenting the stuff inside a vb project.
' when?
'  one wants to drill into a vbp for whatever purpose.
' how?
'  Set up a New Instance
' e.g.  Dim oVBP As VBPInfo: Set oVBP = New VBPInfo.

'  Pass the full file name of the vbp to the ReadVBP method...
' e.g.  oVBP.ReadVBP "SomeProject.vbp"

'  access an attribute
' e.g.  MsgBox oVBP.ProjectName & ": " & oVBP.Description
' or
'       Dim tVBPInfo As QuickVBPInfo
'           tVBPInfo = oVBP.QuickVBPInfo
'           MsgBox tVBPInfo.Name & ": " & tVBPInfo.Description & vbNewLine & tVBPInfo.ProjectPath

'  access list info (forms, modules, classes... etc)
' e.g.  ' ... Print forms info to debug window.
'       If oVBP.CountForms > 0 Then
'           For lngLoop = 1 To oVBP.CountForms
'               Debug.Print oVBP.FormsData(lngLoop)
'           Next lngLoop
'       End If
' or
'       Dim oSSA As StringArray
'           Set oSSA = New StringArray
'           lngCount = oVBP.CountForms
'           For lngLoop = 1 To lngCount
'               oSSA.FromString oVBP.FormsData(lngLoop, "|")
'               Debug.Print oSSA(1)     ' ... name of form (object).
'               Debug.Print oSSA(4)     ' ... full file name and path of form (object).
'           Next lngLoop
'           Set oSSA = Nothing
'
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

' who?
' d.c.

' -------------------------------------------------------------------
' ... v6 updates required.
' ... need to include the folowing:
' ExeName32
' Path32
' CompatibleMode
' CompatibleEXE32
' VersionCompatible32
' ServerSupportFiles
' DllBaseAddress

' ... need to sort out the References Paths / Names etc.

' members.

'==== Subs: 5 ====
'   Class_Terminate:
'   pInit:                        "private helper to set up class for new vbp."
'   pRelease:                     "private helper to release resources and reset fields for re-read or terminate."
'   pGetVBPFileInfo:
'+  ReadVBP:                      "Client Interface to parsing a VBP file for reading its contents (this should be called before using any other members)."
'
'==== Functions: 2 ====
'   pTranslateVBPType: String     "private helper to translate vbpTypeEnum member into s string."
'+  ReadVBG: StringArray              "Returns a StringArray containg vbp information written to a VBG file."
'
'==== Properties: 27 ====
'+  DEConnectionsData(Get): S..   "Returns a StringArray dedicated to Data Connections in a Data Environment items found in the DE's DSR read from the VBP."
'+  ClassesData(Get): StringArray     "Returns a StringArray dedicated to Class items found in VBP."
'+  ComponentsData(Get): SSAr..   "Returns a StringArray dedicated to Component items found in VBP."
'+  CountClasses(Get): Long       "Returns the number of class references found."
'+  CountComponents(Get): Long    "Returns the number of object references found."
'+  CountDesigners(Get): Long     "Returns the number of designer references found."
'+  CountForms(Get): Long         "Returns the number of form references found."
'+  CountModules(Get): Long       "Returns the number of module references found."
'+  CountReferences(Get): Long    "Returns the number of reference references found."
'+  CountUserControls(Get): L..   "Returns the number of user control references found."
'+  Description(Get): String      "Returns a string describing the description given to the VB project."
'+  DesignersData(Get): SSArr..   "Returns a StringArray dedicated to Data Environment & Data Report items found in VBP."
'+  FileName(Get): String         "Returns a string describing the file name of the vb project."
'+  FilePath(Get): String         "Returns a string describing the path to the vb project."
'+  FormsData(Get): StringArray       "Returns a StringArray dedicated to Form items found in VBP."
'+  HelpFile(Get): String         "Returns a string describing the vb project's help file."
'+  Major(Get): String            "Returns a string describing the major version of the VB project."
'+  Minor(Get): String            "Returns a string describing the minor version of the vb project."
'+  ModulesData(Get): StringArray     "Returns a StringArray dedicated to Module items found in VBP."
'~  QuickInfo(Get): QuickVBPI..
'+! ProjectName(Get): String      "Returns a string describing the Name of VB project."|VB_UserMemId = 0
'+  ProjectType(Get): String      "Returns a string describing the type of VB project."
'+  ReferencesData(Get): SSAr..   "Returns a StringArray dedicated to Reference items found in VBP."
'+  ResFileName(Get): String      "Returns the name of the Resource File if present."
'+  Revision(Get): String         "Returns a string describing the minor version of the vb project."
'+  UserControlsData(Get): SS..   "Returns a StringArray dedicated to User Control items found in VBP."
'+  Version(Get): String          "Returns a String concatenation of the various version no. attributes."

Option Explicit

' Dependencies: StringArray.
'               modFileName.
'               modReader
'               modStrings
'               modVB

' v8/9
Private msSourceFileNames As String
Private mVBPTextAsString As String
' v8
Private mIsExe As Boolean
Private mHasRes As Boolean
' v6, more extra properties of the vbp.
Private moRelatedDocs As StringArray
Private moPropertyPages As StringArray

' -------------------------------------------------------------------
' v5, extra properties of the vbp.
Private mExeName32 As String
Private mCommand32 As String
Private mMissingCount As String
Private moMissingFiles As StringArray
' -------------------------------------------------------------------

' ... bunch of SSArrays to help with reading info.
Private moVBPLines As StringArray           ' ... the text in the vbp in lines.
Private moRefs As StringArray               ' ... the references in the vbp.
Private moObjects As StringArray            ' ... the objects in the vbp.
Private moForms As StringArray              ' ... the forms in avbp.
Private moModules As StringArray            ' ... the modules in a vbp.
Private moClasses As StringArray            ' ... the classes in a vbp.
Private moUCs As StringArray                ' ... the user controls in a vbp.
Private moDesigners As StringArray          ' ... the Data Environments / Reports in a vbp.
Private moConnections As StringArray        ' ... the DE Connections in DataEnvironment.
Private moCommands As StringArray           ' ... the DE.Connection.Commands in a DataEnvironment (not in use).

Private moFiles As StringArray

Private mInitialised As Boolean

Private Enum vbpTypeEnum

    evbtNotSet = -1
    
    evbtStandardExe
    evbtActiveXDLL
    evbtActiveXServer
    evbtUserControl
    
End Enum

Private mCompatibleEXE32 As String ' ... private field for property CompatibleEXE32.
Private mCompatibleMode As String ' ... private field for property CompatibleMode.
Private mDescription As String          ' ... private field for property Description.
Private mDllBaseAddress As String ' ... private field for property DllBaseAddress.
Private mVBPFileName As String          ' ... private field for property FileName.
Private mVBPPath As String              ' ... private field for property FilePath.
Private mHelpFile As String             ' ... private field for property HelpFile.
Private mMajor As String                ' ... private field for property Major.
Private mMinor As String                ' ... private field for property Minor.
Private mPath32 As String               ' ... private field for property Path32.
Private mProjectName As String          ' ... private field for property ProjectName.
Private mProjectType As vbpTypeEnum     ' ... private field for property ProjectType.
Private mRevision As String             ' ... private field for property Revision.
Private mResFileName As String
Private mFileInfo As FileNameInfo
Private mServerSupportFiles As String ' ... private field for property ServerSupportFiles.
Private mTitle As String
Private mVersionCompatible32 As String ' ... private field for property VersionCompatible32.

Private mFileNameAndPAth As String

Private Const cFSlash As String = "\"
Private Const cSColon As String = ";"
Private Const cEquals As String = "="
Private Const cHash As String = "#"
Private Const cBang As String = "|"
' -------------------------------------------------------------------
' v6
Private Const c_def_VersionCompatible32 As String = "VersionCompatible32" ' ... default value for property VersionCompatible32.
Private Const c_def_Path32 As String = "Path32" ' ... default value for property Path32.
Private Const c_def_DllBaseAddress As String = "DllBaseAddress" ' ... default value for property DllBaseAddress.
Private Const c_def_ServerSupportFiles As String = "ServerSupportFiles" ' ... default value for property ServerSupportFiles.
Private Const c_def_CompatibleEXE32 As String = "CompatibleEXE32" ' ... default value for property CompatibleEXE32.
Private Const c_def_CompatibleMode As String = "CompatibleMode" ' ... default value for property CompatibleMode.

' v8
Private msRepText As String

Function ProjectReport() As String

Dim cPRep As ProjectReport
Dim stext As String
Dim sColourTable As String

    On Error GoTo ErrHan:
    
    If mInitialised = False Then Exit Function
    
    If Len(msRepText) = 0 Then
            
        Set cPRep = New ProjectReport
        cPRep.Init Me
        
        cPRep.GenerateVBPReport stext
        
        If Len(stext) > 0 Then
        
            sColourTable = modEncode.AllBlackFontColours
            stext = modEncode.BuildRTFString(stext, , , , sColourTable)
        
            msRepText = stext
            stext = vbNullString
            
        End If
        
    End If

ResumeError:
    
    Set cPRep = Nothing
    ProjectReport = msRepText
    
Exit Function

ErrHan:

    Debug.Print "frmViewer.pCreateProjectReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function


Property Get VBPTextAsString() As String
    VBPTextAsString = mVBPTextAsString
End Property

Public Property Get IsExe() As Boolean
Attribute IsExe.VB_Description = "Describes if the project is of the type Standard Exe."
    IsExe = mIsExe
End Property

Public Property Get HasResource() As Boolean
Attribute HasResource.VB_Description = "Describes if the project has a Resource File."
    HasResource = mHasRes
End Property

' -------------------------------------------------------------------

Public Property Get CompatibleMode() As String
Attribute CompatibleMode.VB_Description = "Reads the CompatibleMode entry in a Visual Basic VBP."

    Let CompatibleMode = mCompatibleMode

End Property ' ... CompatibleMode: String.

Public Property Get CompatibleEXE32() As String
Attribute CompatibleEXE32.VB_Description = "Reads the CompatibleEXE32 entry in a Visual Basic VBP."

    Let CompatibleEXE32 = mCompatibleEXE32

End Property ' ... CompatibleEXE32: String.

Public Property Get ServerSupportFiles() As String
Attribute ServerSupportFiles.VB_Description = "Reads the ServerSupportFiles entry in a Visual Basic VBP."

    Let ServerSupportFiles = mServerSupportFiles

End Property ' ... ServerSupportFiles: String.

Public Property Get DllBaseAddress() As String
Attribute DllBaseAddress.VB_Description = "Reads the DllBaseAddress entry in a Visual Basic VBP."

    Let DllBaseAddress = mDllBaseAddress

End Property ' ... DllBaseAddress: String.

Public Property Get Path32() As String
Attribute Path32.VB_Description = "Reads the Path32 entry in a Visual Basic VBP."

    Let Path32 = mPath32

End Property ' ... Path32: String.

Public Property Get VersionCompatible32() As String
Attribute VersionCompatible32.VB_Description = "Reads the VersionCompatible32 entry in a Visual Basic VBP."

    Let VersionCompatible32 = mVersionCompatible32

End Property ' ... VersionCompatible32: String.

Public Property Get PropertyPagesData() As StringArray
Attribute PropertyPagesData.VB_Description = "Returns a StringArray of Property Pages data found in the VBP."
    If Not moPropertyPages Is Nothing Then
        Set PropertyPagesData = moPropertyPages
    Else
        Set PropertyPagesData = New StringArray
    End If
End Property


Public Property Get RelatedDocs() As StringArray
Attribute RelatedDocs.VB_Description = "Returns a StringArray of the related documents found in the VBP."
    If Not moRelatedDocs Is Nothing Then
        Set RelatedDocs = moRelatedDocs
    Else
        Set RelatedDocs = New StringArray
    End If
End Property

Public Property Get ResFileNameAndPath() As String
Attribute ResFileNameAndPath.VB_Description = "Returns the Full File Name and Path of the Resource file referenced in the VBP."

Dim sTmp As String
Dim sPath As String
Dim bExists As Boolean

    sTmp = mResFileName
    
    If Len(sTmp) Then                                   ' ... resource file name?
        
        sPath = mVBPPath                                ' ... project folder?
        
        If Len(sPath) Then
            
            modVB.ReadVBFilePath sPath, sTmp, bExists   ' ... path to project resource file, exists?
            
            If bExists Then
            
                ResFileNameAndPath = sTmp
            
            End If
        
        End If
    
    End If
    
End Property

Public Property Get VBPTextLines() As StringArray
Attribute VBPTextLines.VB_Description = "Returns the text lines of the vbp file as a StringArray.  Caution, this property provides live access to all the StringArray methods available."

    Set VBPTextLines = New StringArray
    
    If Not moVBPLines Is Nothing Then
        Set VBPTextLines = moVBPLines
    End If
    
End Property
' -------------------------------------------------------------------
' v5, extra properties of the vbp.
Public Property Get ExeName32() As String
Attribute ExeName32.VB_Description = "Returns the ExeName32 value from a VBP (provides the name of the last compiled file)."
    ExeName32 = mExeName32
End Property

Public Property Get Command32() As String
Attribute Command32.VB_Description = "The Command Line from a VB Project."
    Command32 = mCommand32
End Property

Public Property Get MissingCount() As Long
Attribute MissingCount.VB_Description = "Returns the number of physical file references found in a VBP that could not be located."
    MissingCount = mMissingCount
End Property

Public Property Get MissingFiles() As StringArray
Attribute MissingFiles.VB_Description = "Returns a StringArray of the missing VBP files."
    If Not moMissingFiles Is Nothing Then
        Set MissingFiles = moMissingFiles
    Else
        Set MissingFiles = New StringArray
    End If
End Property
' -------------------------------------------------------------------'...

Public Property Get AttributeDelimiter() As String
Attribute AttributeDelimiter.VB_Description = "Exposes the Delimiter used to concatenate the various attributes of an object, reference or file based item in a vbp."
    AttributeDelimiter = cBang
End Property

Private Sub Class_Terminate()
    pRelease
End Sub

Public Property Get DEConnectionsData() As StringArray
Attribute DEConnectionsData.VB_Description = "Returns a SSArray dedicated to Data Connections in a Data Environment items found in the DE's DSR read from the VBP."
    If Not moConnections Is Nothing Then
        Set DEConnectionsData = moConnections
    Else
        Set DEConnectionsData = New StringArray
    End If
End Property

Public Property Get ClassesData() As StringArray
Attribute ClassesData.VB_Description = "Returns a SSArray dedicated to Class items found in VBP."
    If Not moClasses Is Nothing Then
        Set ClassesData = moClasses
    Else
        Set ClassesData = New StringArray
    End If
End Property

Public Property Get ComponentsData() As StringArray
Attribute ComponentsData.VB_Description = "Returns a SSArray dedicated to Component items found in VBP."
    If Not moObjects Is Nothing Then
        Set ComponentsData = moObjects
    Else
        Set ComponentsData = New StringArray
    End If
End Property

Public Property Get CountClasses() As Long
Attribute CountClasses.VB_Description = "Returns the number of class references found."
    If Not moClasses Is Nothing Then CountClasses = moClasses.Count
End Property

Public Property Get CountComponents() As Long
Attribute CountComponents.VB_Description = "Returns the number of object references found."
    If Not moObjects Is Nothing Then CountComponents = moObjects.Count
End Property

Public Property Get CountDesigners() As Long
Attribute CountDesigners.VB_Description = "Returns the number of designer references found."
    If Not moDesigners Is Nothing Then CountDesigners = moDesigners.Count
End Property

Public Property Get CountForms() As Long
Attribute CountForms.VB_Description = "Returns the number of form references found."
    If Not moForms Is Nothing Then CountForms = moForms.Count
End Property

Public Property Get CountModules() As Long
Attribute CountModules.VB_Description = "Returns the number of module references found."
    If Not moModules Is Nothing Then CountModules = moModules.Count
End Property

Public Property Get CountReferences() As Long
Attribute CountReferences.VB_Description = "Returns the number of reference references found."
    If Not moRefs Is Nothing Then CountReferences = moRefs.Count
End Property

Public Property Get CountUserControls() As Long
Attribute CountUserControls.VB_Description = "Returns the number of user control references found."
    If Not moUCs Is Nothing Then CountUserControls = moUCs.Count
End Property
' -------------------------------------------------------------------
' v6, added, number of property pages.
Public Property Get CountPropertyPages() As Long
Attribute CountPropertyPages.VB_Description = "Returns a count of the number of Property Pages found in a VBP."
    If Not moPropertyPages Is Nothing Then CountPropertyPages = moPropertyPages.Count
End Property
' -------------------------------------------------------------------

Public Property Get CountVBPFiles() As Long
Attribute CountVBPFiles.VB_Description = "Returns the number of files referenced in the vbp."
Dim lngResult As Long

    lngResult = lngResult + Me.CountClasses
    lngResult = lngResult + Me.CountDesigners
    lngResult = lngResult + Me.CountForms
    lngResult = lngResult + Me.CountModules
    lngResult = lngResult + Me.CountUserControls
    lngResult = lngResult + Me.CountPropertyPages
    
    CountVBPFiles = lngResult
    
    lngResult = 0
    
End Property

Public Property Get Description() As String
Attribute Description.VB_Description = "Returns a string describing the description given to the VB project."
    Let Description = mDescription
End Property

Public Property Get DesignersData() As StringArray
Attribute DesignersData.VB_Description = "Returns a SSArray dedicated to Data Environment & Data Report items found in VBP."
    If Not moUCs Is Nothing Then
        Set DesignersData = moDesigners
    Else
        Set DesignersData = New StringArray ' ... not much use as empty and read only but means client caller doesn't need to handle a Nothing return.
    End If
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns a string describing the file name of the vb project."
    Let FileName = mVBPFileName
End Property

Public Property Get FileNameAndPath() As String
Attribute FileNameAndPath.VB_Description = "Returns the Full File Name and Path of a VB Project."
    FileNameAndPath = mFileNameAndPAth
End Property

Public Property Get FilePath() As String
Attribute FilePath.VB_Description = "Returns a string describing the path to the vb project."
    Let FilePath = mVBPPath
End Property

Public Property Get FilesData() As StringArray
Attribute FilesData.VB_Description = "Returns a StringArray of physical files found referenced in a VBP."
    If Not moFiles Is Nothing Then
        Set FilesData = moFiles
    Else
        Set FilesData = New StringArray
    End If
End Property

Public Property Get FormsData() As StringArray
Attribute FormsData.VB_Description = "Returns a SSArray dedicated to Form items found in VBP."
    If Not moForms Is Nothing Then
        Set FormsData = moForms
    Else
        Set FormsData = New StringArray
    End If
End Property

Public Property Get HelpFile() As String
Attribute HelpFile.VB_Description = "Returns a string describing the vb project's help file."
    Let HelpFile = mHelpFile
End Property

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a Boolean describing whether the current instance of the class has been initialised and therefore useable as is, True, or not, False."
    Initialised = mInitialised
End Property

Public Property Get Major() As String
Attribute Major.VB_Description = "Returns a string describing the major version of the VB project."
    Let Major = mMajor
End Property

Public Property Get Minor() As String
Attribute Minor.VB_Description = "Returns a string describing the minor version of the vb project."
    Let Minor = mMinor
End Property

Public Property Get ModulesData() As StringArray
Attribute ModulesData.VB_Description = "Returns a SSArray dedicated to Module items found in VBP."
    If Not moModules Is Nothing Then
        Set ModulesData = moModules
    Else
        Set ModulesData = New StringArray
    End If
End Property


Private Sub pInit()
Attribute pInit.VB_Description = "private helper to set up class for new vbp."
' -------------------------------------------------------------------
' ... helper to set up class for new vbp.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    pRelease     ' ... see note within.
    ' -------------------------------------------------------------------
    ' ... multiple statements per line.
    Set moVBPLines = New StringArray:       moVBPLines.Sortable = False
    Set moRefs = New StringArray:           moRefs.Sortable = True
    Set moObjects = New StringArray:        moObjects.Sortable = True
    Set moForms = New StringArray:          moForms.Sortable = True
    Set moModules = New StringArray:        moModules.Sortable = True
    Set moClasses = New StringArray:        moClasses.Sortable = True
    Set moUCs = New StringArray:            moUCs.Sortable = True
    Set moDesigners = New StringArray:      moDesigners.Sortable = True
    Set moConnections = New StringArray:    moConnections.Sortable = True
    Set moCommands = New StringArray:       moCommands.Sortable = True
    Set moFiles = New StringArray:          moFiles.Sortable = True
    ' -------------------------------------------------------------------
    ' v 5
    Set moMissingFiles = New StringArray:   moMissingFiles.Sortable = True
    ' -------------------------------------------------------------------
    ' v6.
    Set moRelatedDocs = New StringArray:    moRelatedDocs.Sortable = True
    Set moPropertyPages = New StringArray:  moPropertyPages.Sortable = True
    ' -------------------------------------------------------------------
    
Exit Sub
ErrHan:
    Debug.Print "VBPInfo.pInit.Error: " & Err.Number & "; " & Err.Description
    Err.Clear: Resume Next
End Sub ' ... pInit:


Private Sub pRelease()
Attribute pRelease.VB_Description = "private helper to release resources and reset fields for re-read or terminate."
' -------------------------------------------------------------------
' ... helper to release resources and reset fields for re-read or terminate.
' ... Note: clears current fields so beware calling it after setting setting up.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    mInitialised = False    ' ... set this to true upon opening and reading the vbp.
    ' -------------------------------------------------------------------
    ' ... objects.
    If Not moVBPLines Is Nothing Then Set moVBPLines = Nothing
    If Not moRefs Is Nothing Then Set moRefs = Nothing
    If Not moObjects Is Nothing Then Set moObjects = Nothing
    If Not moForms Is Nothing Then Set moForms = Nothing
    If Not moModules Is Nothing Then Set moModules = Nothing
    If Not moClasses Is Nothing Then Set moClasses = Nothing
    If Not moUCs Is Nothing Then Set moUCs = Nothing
    If Not moDesigners Is Nothing Then Set moDesigners = Nothing
    If Not moConnections Is Nothing Then Set moConnections = Nothing
    If Not moCommands Is Nothing Then Set moCommands = Nothing
    If Not moFiles Is Nothing Then Set moFiles = Nothing
    If Not moRelatedDocs Is Nothing Then Set moRelatedDocs = Nothing ' v6.
    If Not moPropertyPages Is Nothing Then Set moPropertyPages = Nothing ' v6.
    ' -------------------------------------------------------------------
    ' ... fields.
    mDescription = vbNullString
    mVBPFileName = vbNullString
    mVBPPath = vbNullString
    mHelpFile = vbNullString
    mMajor = vbNullString
    mMinor = vbNullString
    mProjectName = vbNullString
    mRevision = vbNullString
    mResFileName = vbNullString
    mFileNameAndPAth = vbNullString
    mTitle = vbNullString
    ' -------------------------------------------------------------------
    mProjectType = vbpTypeEnum.evbtNotSet
    ' -------------------------------------------------------------------
    ' v5
    mExeName32 = vbNullString
    mCommand32 = vbNullString
    mDllBaseAddress = vbNullString
    mPath32 = vbNullString
    mCompatibleEXE32 = vbNullString
    mVersionCompatible32 = vbNullString
    mServerSupportFiles = vbNullString
    mCompatibleMode = vbNullString
    
    mMissingCount = 0
    If Not moMissingFiles Is Nothing Then Set moMissingFiles = Nothing
    ' v8 ----------------------------------------------------------------
    mIsExe = False
    mHasRes = False
    ' v8/9 --------------------------------------------------------------
    msSourceFileNames = vbNullString
    mVBPTextAsString = vbNullString
Exit Sub
ErrHan:
    Debug.Print "VBPInfo.pRelease.Error: " & Err.Number & "; " & Err.Description
    Err.Clear: Resume Next
End Sub ' ... pRelease:

Friend Property Get QuickInfo() As QuickVBPInfo
Dim xInfo As QuickVBPInfo
' -------------------------------------------------------------------
' ... Return a Read-Only QuickVPInfo describing the current VBP in use.
' -------------------------------------------------------------------
    
    xInfo.Description = mDescription
    xInfo.HelpFile = mHelpFile
    xInfo.MajorVersion = mMajor
    xInfo.MinorVersion = mMinor
    xInfo.Name = mProjectName
    xInfo.ProjectFileName = mVBPFileName
    xInfo.ProjectPath = mVBPPath
    xInfo.RevisionVersion = mRevision
    xInfo.Type = Me.ProjectType
    xInfo.ResourceFile = mResFileName
    xInfo.Version = Me.Version
    
    QuickInfo = xInfo
    
End Property

Public Property Get ProjectName() As String
Attribute ProjectName.VB_Description = "Returns a string describing the Name of VB project."
Attribute ProjectName.VB_UserMemId = 0
    Let ProjectName = mProjectName
End Property

Public Property Get ProjectType() As String
Attribute ProjectType.VB_Description = "Returns a string describing the type of VB project."
    Let ProjectType = pTranslateVBPType(mProjectType)
End Property

Private Function pTranslateVBPType(pMember As vbpTypeEnum) As String
Attribute pTranslateVBPType.VB_Description = "private helper to translate vbpTypeEnum member into s string."
Dim sReturn As String
' -------------------------------------------------------------------
' ... helper to translate vbpTypeEnum member into s string.
' -------------------------------------------------------------------
    
    sReturn = "Unknown"
    Select Case pMember
        Case evbtNotSet
        Case evbtStandardExe:       sReturn = "Standard Exe"
        Case evbtActiveXDLL:        sReturn = "DLL"
        Case evbtActiveXServer:     sReturn = "ActiveX Server"
        Case evbtUserControl:       sReturn = "User Control"
    End Select
    pTranslateVBPType = sReturn
    
End Function

Public Function ReadVBG(ByVal FileName As String) As StringArray
Attribute ReadVBG.VB_Description = "Returns a SSArray containg vbp information written to a VBG file."
Dim sTmpFile As String
Dim sTmpFileText As String
Dim oNewObject As StringArray
Dim oRetObject As StringArray
Dim lngCount As Long
Dim lngLoop As Long
Dim sTmpString As String
Dim bOK As Boolean
Dim sLeft As String
Dim sRight As String
'Dim sFile As String
Dim bExists As Boolean
Dim lngFound As Long

    ' -------------------------------------------------------------------
    Set oRetObject = New StringArray
    ' -------------------------------------------------------------------
    sTmpFile = LCase$(FileName)
    ' -------------------------------------------------------------------
    If Len(sTmpFile) Then
        If Dir$(sTmpFile, vbNormal) <> "" Then
            lngFound = modStrings.InStrRevChar(FileName, cFSlash)
            If lngFound > 0 Then
                mVBPPath = Left$(FileName, lngFound - 1)
            End If
            mVBPFileName = FileName
            If Right$(sTmpFile, 4) = ".vbg" Then
                sTmpFileText = modReader.ReadFile(sTmpFile, bOK)
                Set oNewObject = New StringArray
                oNewObject.FromString sTmpFileText, vbCrLf
                lngCount = oNewObject.Count
                For lngLoop = 1 To lngCount
                    sTmpString = oNewObject(lngLoop)
                    If lngLoop = 1 Then
                        
                        If sTmpString = "VBGROUP 5.0" Then
                            bOK = True
                        Else
                            ' ... no group signature on line 1.
                            Exit For
                        End If
                    Else
                        bExists = False
                        modStrings.SplitStringPair sTmpString, cEquals, sLeft, sRight
                        Select Case LCase$(sLeft)
                            Case "project", "startupproject"
                                modVB.ReadVBFilePath mVBPPath, sRight, bExists
                        End Select
                        If bExists Then
                            oRetObject.AddItemString sRight
                        End If
                    End If
                Next lngLoop
            End If
        End If
    End If
    ' -------------------------------------------------------------------
    Set ReadVBG = oRetObject
    ' -------------------------------------------------------------------
    If Not oRetObject Is Nothing Then
        Set oRetObject = Nothing
    End If
    If Not oNewObject Is Nothing Then
        Set oNewObject = Nothing
    End If
End Function

Private Sub pGetVBPFileInfo(ByVal pFileName As String)
    ' split the VBP's file name into its parts.
    ' gotta have this even if just reading text passed.
    modFileName.ParseFileNameEx pFileName, mFileInfo
    mVBPFileName = mFileInfo.File
    mVBPPath = mFileInfo.Path
    mFileNameAndPAth = mFileInfo.PathAndName
End Sub

Public Sub ReadVBP(ByVal FileName As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute ReadVBP.VB_Description = "Client Interface to parsing a VBP file for reading its contents (this should be called before using any other members)."

Dim bOK As Boolean
Dim sErrMsg As String
Dim sFileText As String
Dim lngLineCount As Long
Dim lngLoop As Long
Dim sTmpLine As String
Dim sLeft As String
Dim sRight As String
Dim sFileName As String
Dim lngFound As Long
Dim bIsRef As Boolean
Dim bIsClass As Boolean
Dim bIsForm As Boolean
Dim bIsDataEnv As Boolean
Dim sTmpAdd As String
Dim intFormType As Integer
Dim bFileExists As Boolean
Dim lngMissingCount As Long
Dim sTmp As String
' ... Data Environment.
Dim cnLoop As Long  ' ... connections loop
'Dim cmLoop As Long  ' ... commands loop
Dim numConnections As Long
Dim currConnection As Long
Dim cnCount As Long
'Dim cmCount As Long
Dim saDE As StringArray
Dim sTmpDELine As String
Dim sLeftDE As String
Dim sRightDE As String

Dim sRefPath As String ' v4/5   ' ... path to a reference / object.
Dim sWinDrive As String ' v4/5    ' ... windows drive.
' -------------------------------------------------------------------
Dim xRefInfo As ReferenceInfo ' ... v6, attempting to refine and hone References which can too easily come out wrong.


' -------------------------------------------------------------------
' ... Main Interface to this class; read and parse vbp for refs, objs, forms, classes, mods & usercontrols.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    pInit
    sFileName = FileName
    sFileText = modReader.ReadFile(sFileName, bOK, sErrMsg)
    
    If bOK = True Then
    
        mVBPTextAsString = sFileText
        moVBPLines.FromString sFileText, vbCrLf
        
        pGetVBPFileInfo sFileName
        
        lngLineCount = moVBPLines.Count
        
        sWinDrive = Left$(GetWinSysDir, 3)
        
        For lngLoop = 1 To lngLineCount
            
            sRefPath = vbNullString ' v4/5
            bFileExists = False
            
            lngFound = 0
            sTmpLine = moVBPLines(lngLoop)
            
            If InStr(1, sTmpLine, Chr$(34)) > 0 Then
                ' ... if the line has quotes then remove them for easier life.
                modStrings.RemoveQuotes sTmpLine
            End If
            
            modStrings.SplitStringPair sTmpLine, cEquals, sLeft, sRight, True, True
            
            Select Case sLeft
                
                Case "Type"
                    Select Case sRight
                        Case "Exe":     mProjectType = evbtStandardExe
                        Case "OleDll":  mProjectType = evbtActiveXDLL
                        Case "OleExe":  mProjectType = evbtActiveXServer
                        Case "Control": mProjectType = evbtUserControl
                    End Select
                    ' -------------------------------------------------------------------
                    mIsExe = mProjectType = evbtStandardExe ' v8, is an exe program?
                ' -------------------------------------------------------------------
                ' ... v6, added.
                Case c_def_DllBaseAddress
                    mDllBaseAddress = sRight
                    
                Case c_def_Path32
                    mPath32 = sRight
                    
                Case c_def_CompatibleEXE32
                    mCompatibleEXE32 = sRight
                
                Case c_def_VersionCompatible32
                    mVersionCompatible32 = sRight
                
                Case c_def_ServerSupportFiles
                    mServerSupportFiles = sRight
                
                Case c_def_CompatibleMode
                    mCompatibleMode = sRight
                ' -------------------------------------------------------------------
                
                Case "Name":            mProjectName = sRight
                Case "Description":
                    ' -------------------------------------------------------------------
                    ' ... v6, update.
                    ' ... MZ-Tools stores Tasks in the VBP and these
                    ' ... are prefixed with the word ' Description '
                    ' ... so, if Description is already set, the plan will be
                    ' ... to ignore any further updating.
                    If Len(mDescription) = 0 Then
                        mDescription = sRight
                    End If
                    ' -------------------------------------------------------------------
                Case "HelpFile":        mHelpFile = sRight
                Case "MajorVer":        mMajor = sRight
                Case "MinorVer":        mMinor = sRight
                Case "RevisionVer":     mRevision = sRight
                Case "Title":           mTitle = sRight
                Case "ResFile32":       mResFileName = sRight
                Case "ExeName32":       mExeName32 = sRight
                
                Case "PropertyPage" ' ... v6.
                    
                    modVB.ReadVBFilePath mVBPPath, sRight, bFileExists
                    ' ------------------------------------------------------------------- ' v 5
                    If bFileExists = False Then
                        ' -------------------------------------------------------------------
                        ' v 5
                        lngMissingCount = lngMissingCount + 1
                        sLeft = "Missing File: " & CStr(lngMissingCount)
                        If Len(Trim$(sRight)) = 0 Then
                            sRight = sLeft
                        End If
                        moMissingFiles.AddItemString sRight
                        ' -------------------------------------------------------------------
                    Else
                        modVB.ReadVBName sRight, sLeft ', bIsForm, intFormType
                    End If
                    
                    ' -------------------------------------------------------------------
                    ' ... Property Page added as Name | Type | File Exists | File Name with Full Path
                    sTmpAdd = sLeft & cBang & "10" & cBang & IIf(bFileExists, "1", "0") & cBang & sRight
                    moPropertyPages.AddItemString sTmpAdd
                    moFiles.AddItemString sTmpAdd
                    
                Case "RelatedDoc":      ' ... array of related documents for copy. v6
                    If Len(sRight) Then
                        modVB.ReadVBFilePath mVBPPath, sRight, bFileExists
                        If bFileExists Then
                            moRelatedDocs.AddItemString sRight
                        End If
                    End If
                Case "Command32":       mCommand32 = sRight
                    
                    ' ... this is the command line entry for conditional compilation.
                    
                Case "Object", "Reference"
                    ' -------------------------------------------------------------------
                    ' ... TLB Info 32 can accept a file name and path or a registered Name,
                    ' ... haven't tried it with a GUID though.
                    ' -------------------------------------------------------------------
                    modVB.ParseReferenceInfo sTmpLine, xRefInfo, bOK, sErrMsg
                    ' -------------------------------------------------------------------
                    sRight = xRefInfo.FileDescription
                    If Len(sRight) = 0 Then
                        sRight = xRefInfo.FileName
                    End If
                    ' -------------------------------------------------------------------
                    bIsRef = xRefInfo.Reference
                    ' -------------------------------------------------------------------
                    sLeft = xRefInfo.Guid
                    ' -------------------------------------------------------------------
                    sRefPath = xRefInfo.FileNameAndPath
                    If Len(sRefPath) > 0 Then
                        ' -------------------------------------------------------------------
                        modVB.ReadVBFilePath mVBPPath, sRefPath, bFileExists
                        ' -------------------------------------------------------------------
                        If bFileExists Then
                            sLeft = sLeft & " " & sRefPath
                        Else
                            sLeft = sLeft & " " & xRefInfo.FileName
                        End If
                    Else
                        sLeft = sLeft & " " & xRefInfo.FileName
                    End If
                    ' -------------------------------------------------------------------
                    sTmpAdd = sRight & cBang & IIf(bIsRef = True, "0", "1") & "|1|" & sLeft
                    ' -------------------------------------------------------------------
                    If bIsRef Then
                        moRefs.AddItemString sTmpAdd
                    Else
                        moObjects.AddItemString sTmpAdd
                    End If
                    ' -------------------------------------------------------------------
                    
                Case "Form", "UserControl"
                    bIsForm = Left$(sLeft, 1) = "F"
                    modVB.ReadVBFilePath mVBPPath, sRight, bFileExists
                    intFormType = 0
                    If bFileExists = True Then
                        modVB.ReadVBName sRight, sLeft, bIsForm, intFormType
                    Else
                        lngMissingCount = lngMissingCount + 1
                        sLeft = "Missing File: " & CStr(lngMissingCount)
                        If Len(Trim$(sRight)) = 0 Then
                            sRight = sLeft
                        End If
                        ' -------------------------------------------------------------------
                        ' v 5
                        moMissingFiles.AddItemString sRight
                        ' -------------------------------------------------------------------
                    End If
                    ' ... note: form, type = 2 + form type, 21 = MDI, 22 = Normal, 23 = MDI Child else invalid.
                    ' ... Form/UserControl added as Name | Type | File Exists | File Name with Full Path
                    sTmpAdd = sLeft & cBang & IIf(bIsForm = True, "2" & CStr(intFormType), "5") & cBang & IIf(bFileExists, "1", "0") & cBang & sRight
                    If bIsForm Then
                        moForms.AddItemString sTmpAdd
                    Else
                        moUCs.AddItemString sTmpAdd
                    End If
                    moFiles.AddItemString sTmpAdd
                    
                Case "Class", "Module"
                    
                    bIsClass = Left$(sLeft, 1) = "C"
                    modStrings.SplitStringPair sRight, cSColon, sLeft, sRight, True, True
                    modVB.ReadVBFilePath mVBPPath, sRight, bFileExists
                    ' ------------------------------------------------------------------- ' v 5
                    If bFileExists = False Then
                        ' -------------------------------------------------------------------
                        ' v 5
                        lngMissingCount = lngMissingCount + 1
                        sLeft = "Missing File: " & CStr(lngMissingCount)
                        If Len(Trim$(sRight)) = 0 Then
                            sRight = sLeft
                        End If
                        moMissingFiles.AddItemString sRight
                        ' -------------------------------------------------------------------
                    End If
                    
                    ' -------------------------------------------------------------------
                    ' ... Class/Module added as Name | Type | File Exists | File Name with Full Path
                    sTmpAdd = sLeft & cBang & IIf(bIsClass = True, "4", "3") & cBang & IIf(bFileExists, "1", "0") & cBang & sRight
                    If bIsClass Then
                        moClasses.AddItemString sTmpAdd
                    Else
                        moModules.AddItemString sTmpAdd
                    End If
                    moFiles.AddItemString sTmpAdd
                    
                Case "Designer"
                    ' -------------------------------------------------------------------
                    ' v6, remove line causing bug reading designers.
'                    modStrings.SplitStringPair sRight, cEquals, sLeft, sRight, True, True
                    modVB.ReadVBFilePath mVBPPath, sRight, bFileExists
                    
                    If bFileExists = True Then
                        modVB.ReadVBName sRight, sLeft
                        sTmp = modReader.ReadFile(sRight)
                        bIsDataEnv = InStr(1, sTmp, "Begin {C0E45035-5775-11D0-B388-00A0C9055D8E}") > 0 ' ... Date Environment GUID.
                        
                        If bIsDataEnv Then
                            ' ... could get the names of the connections and the commands?
                            Set saDE = New StringArray
                            saDE.FromString sTmp, vbCrLf
                            cnCount = saDE.Count
                            
                            For cnLoop = 1 To cnCount
                                sTmpDELine = Trim$(saDE(cnLoop))
                                If Left$(sTmpDELine, 14) = "NumConnections" Then
                                    modStrings.SplitStringPair sTmpDELine, cEquals, sLeftDE, sRightDE, True, True
                                    numConnections = Val(sRightDE)
                                    Exit For
                                End If
                            Next cnLoop
                            
                            If numConnections > 0 Then
                                currConnection = 1
                                For cnLoop = cnLoop + 1 To cnCount
                                    sTmpDELine = Trim$(saDE(cnLoop))
                                    If sTmpDELine = "BeginProperty Connection" & CStr(currConnection) Then
                                        If cnLoop + 1 <= cnCount Then
                                            currConnection = currConnection + 1
                                            cnLoop = cnLoop + 1
                                            sTmpDELine = Trim$(saDE(cnLoop))
                                            modStrings.SplitStringPair sTmpDELine, cEquals, sLeftDE, sRightDE, True, True
                                            If InStr(sRightDE, Chr$(34)) > 0 Then
                                                sRightDE = Mid$(sRightDE, 2, Len(sRightDE) - 2) ' ... trim quotes?
                                            End If
                                            sTmpAdd = sRightDE & "|9|1|" & sRight 'sLeft    ' connection name | type | exists (true) | DE Name
                                            moConnections.AddItemString sTmpAdd
                                        End If
                                    End If
                                    
                                Next cnLoop
                            End If
                            
                        End If
                    Else
                        lngMissingCount = lngMissingCount + 1
                        sLeft = "Missing File: " & CStr(lngMissingCount)
                        If Len(Trim$(sRight)) = 0 Then
                            sRight = sLeft
                        End If
                        ' -------------------------------------------------------------------
                        ' v 5
                        moMissingFiles.AddItemString sRight
                        ' -------------------------------------------------------------------

                    End If
                    ' -------------------------------------------------------------------
                    ' ... Class/Module added as Name | Type | File Exists | File Name with Full Path
                    sTmpAdd = sLeft & cBang & IIf(bIsDataEnv = True, "7", "8") & cBang & IIf(bFileExists, "1", "0") & cBang & sRight
                    moDesigners.AddItemString sTmpAdd
                    moFiles.AddItemString sTmpAdd
                    
            End Select
        Next lngLoop
        ' -------------------------------------------------------------------
        ' ... multiple statements per line: compact and sort array data.
        If moRefs.Count > 0 Then moRefs.Sort
        If moObjects.Count > 0 Then moObjects.Sort
        If moForms.Count > 0 Then moForms.Sort
        If moModules.Count > 0 Then moModules.Sort
        If moClasses.Count > 0 Then moClasses.Sort
        If moUCs.Count > 0 Then moUCs.Sort
        If moFiles.Count > 0 Then moFiles.Sort
        If moRelatedDocs.Count > 0 Then moRelatedDocs.Sort ' v6.
        ' -------------------------------------------------------------------
        ' v 5
        If moMissingFiles.Count > 0 Then moMissingFiles.Sort
        ' v8 ----------------------------------------------------------------
        If Len(ResFileNameAndPath) > 0 Then
            mHasRes = True
        End If
        ' -------------------------------------------------------------------
    Else
        ' ... raise an error to the handler.
        Err.Raise vbObject + 1000, , "Failed to Read VBP..." & vbNewLine & sErrMsg
    End If
        
ResumeError:
    On Error GoTo 0
    ' -------------------------------------------------------------------
    ' v 5
    mMissingCount = lngMissingCount
    ' -------------------------------------------------------------------
    mInitialised = bOK
    pOK = bOK
    pErrMsg = sErrMsg
    ' ... multiple statements on following lines...
    sFileText = vbNullString: sTmpLine = vbNullString: sLeft = vbNullString: sRight = vbNullString
    sTmpDELine = vbNullString: sLeftDE = vbNullString: sRightDE = vbNullString
    sRefPath = vbNullString: sWinDrive = vbNullString: sFileName = vbNullString
    
    lngFound = 0&: cnCount = 0&: cnLoop = 0&: numConnections = 0&
    
    If Not saDE Is Nothing Then
        Set saDE = Nothing
    End If
    
Exit Sub
ErrHan:
    bOK = False
    sErrMsg = Err.Description
    Debug.Print "VBPInfo.ReadVBP.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub ' ... ReadVBP:

Public Property Get ReferencesData() As StringArray
Attribute ReferencesData.VB_Description = "Returns a SSArray dedicated to Reference items found in VBP."
    If Not moRefs Is Nothing Then
        Set ReferencesData = moRefs
    Else
        Set ReferencesData = New StringArray
    End If
End Property

Public Property Get ResFileName() As String
Attribute ResFileName.VB_Description = "Returns the name of the Resource File if present."
    ResFileName = mResFileName
End Property

Public Property Get Revision() As String
Attribute Revision.VB_Description = "Returns a string describing the minor version of the vb project."
    Let Revision = mRevision
End Property

Public Property Get Title() As String
Attribute Title.VB_Description = "Returns the Title Attribute of a VB Project."
    Title = mTitle
End Property

Public Property Get UserControlsData() As StringArray
Attribute UserControlsData.VB_Description = "Returns a SSArray dedicated to User Control items found in VBP."
    If Not moUCs Is Nothing Then
        Set UserControlsData = moUCs
    Else
        Set UserControlsData = New StringArray
    End If
End Property

Public Property Get Version() As String
Attribute Version.VB_Description = "Returns a String concatenation of the various version no. attributes."
Dim sReturn As String
    On Error GoTo ErrHan:
    sReturn = CStr(Val(mMajor)) & "." & CStr(Val(mMinor)) & "." & Format$(Val(mRevision), cNumFormat)
    Version = sReturn
Exit Property
ErrHan:
    Debug.Print "VBPInfo.Version.Error: " & Err.Number & "; " & Err.Description
End Property ' ... Version: String

