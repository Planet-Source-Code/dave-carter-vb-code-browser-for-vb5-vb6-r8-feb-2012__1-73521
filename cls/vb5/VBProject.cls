VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class to load and access all vbp source code members into an array of codeinfo objects"
' what?
'   a class to read and process the entire source code references of a visual basic (5/6) project
' why?
'   provide access to a project and its members and member code
' when?
'   i wrote this in an effort to contain a project within a single object
'   to make it easier to write new forms and functions.
'   the first example of its use is frmMembers.
' how?
'
' who?
'   dc
' ------------------------------------------------------------

Option Explicit

' -------------------------------------------------------------------
' some events to inform the owner that things are on their way, as it were.

Event ProcessFile(ByVal ItemIndex As Long, ByVal MaxItems As Long, ByVal ItemName As String)
Attribute ProcessFile.VB_Description = "Event saying a source file is being processed, giving its index in the vbp, the no. of source files in the vbp and the name of the source code file."
Event ProcessMember(ByVal MemberIndex As Long, ByVal MaxMembers As Long, ByVal MemberName As String)
Attribute ProcessMember.VB_Description = "Event, declaring the member of a source code file is being processed, with its member position and name and the total no. of members in the source code file."
Event MembersAndRTFArrays(CodeMembers() As CodeInfo, CodeAsRTF() As StringWorker)
Attribute MembersAndRTFArrays.VB_Description = "Event passing a complete array of CodeInfo instances representing the project's source files and an array of RTF encoded StringWorker instances presenting the entire code of each source file in Rich Text."


Private mUBoundCodeArray As Long
Private mCodeFileCount As Long                  ' no. of code files in vbp

Private moCodeInfoArray() As CodeInfo           ' array of code info instances
Private mswCodeRTFArray() As StringWorker       ' array of code text to rich text, stringworker()

Private moVBPInfo As VBPInfo                    ' the vbp inof instance in use

Private mVBPFilesLoaded As Boolean
Private mInitialised As Boolean                 ' flag, returns true if main sub successfully executed
Private mLoadRTF As Boolean

Private mAttributeDelimiter As String

Private msaFileNamesData As StringArray         ' names of all the vbp's source files
Private msaMembersData As StringArray           ' list of project's members
Private msaMembersDescriptions As StringArray   ' list of project member's descriptions

Public Property Get MembersData() As StringArray
Attribute MembersData.VB_Description = "Returns a sorted string array of members and their various attributes."
' returns a list of all members found in project
    If Not msaMembersData Is Nothing Then
        Set MembersData = msaMembersData
    Else
        Set MembersData = New StringArray
    End If
End Property

Public Property Get CountOfSourceFiles() As Long
Attribute CountOfSourceFiles.VB_Description = "Returns the number of Source Code Files described in the Project."
' ... Returns the number of Source Code Files described in the Project.
    CountOfSourceFiles = mCodeFileCount
End Property

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a boolean describing whether the VBP has been set up and code files read."
' returns true if class instance successfully loaded, else false
    Initialised = mInitialised
End Property

Public Property Get FilesNameInfoData() As StringArray
Attribute FilesNameInfoData.VB_Description = "Returns a string array of vbp source file info. that may be parsed in to a DataInfo structure / type."
' ... Returns a string array of vbp source file info. that may be parsed in to a DataInfo structure / type.
    If msaFileNamesData Is Nothing Then
        Set msaFileNamesData = New StringArray
        msaFileNamesData.Sortable = True
    End If
    Set FilesNameInfoData = msaFileNamesData
End Property

Private Sub pRelease()
Attribute pRelease.VB_Description = "release local resources, e.g. re-run / terminate."
' release local resources, e.g. re-run / terminate
    On Error GoTo ErrHan:
    
    Erase moCodeInfoArray
    Erase mswCodeRTFArray
    
    Set moVBPInfo = Nothing
    Set msaFileNamesData = Nothing
    Set msaMembersData = Nothing
    Set msaMembersDescriptions = Nothing
    
Exit Sub
ErrHan:

    Debug.Print "VBProject.pRelease.Error: " & Err.Number & "; " & Err.Description
    Err.Clear
    Resume Next

End Sub

Private Sub pInit()
Attribute pInit.VB_Description = "reset local resources to defaults for new run."
' reset local resources to defaults for new run
    pRelease
    
    mLoadRTF = False
    mVBPFilesLoaded = False
    mInitialised = False
    mCodeFileCount = 0
    mUBoundCodeArray = -1
    
    Set msaFileNamesData = New StringArray
    msaFileNamesData.Sortable = True
    
    Set msaMembersData = New StringArray
    msaMembersData.Sortable = True
    
    Set msaMembersDescriptions = New StringArray
    msaMembersDescriptions.Sortable = True
    
End Sub

Public Sub Init(ByRef pVBPInfo As VBPInfo, _
       Optional ByVal pWriteRTF As Boolean = False, _
       Optional ByRef pOK As Boolean = False, _
       Optional ByRef pErrMsg As String = vbNullString)
Attribute Init.VB_Description = "main sub to kick off reading a project's source files."
' main sub to kick off reading a project's source files

'... Parameters.
'    R__ pVBPInfo: VBPInfo           ' ... A VBPInfo instance loaded with data.
'    V__ pWriteRTF: Boolean          ' ... flag, when true compiles RTF for each code module
'    R__ pOK: Boolean                ' ... returns true if sub executed successfully
'    R__ pErrMsg: String             ' ... returns an error message if sub failed successful execution

Dim bOK As Boolean
Dim sErrMsg As String

    On Error GoTo ErrHan:
    
    pInit
    ' check the vbpinfo object (needs to be initialised) ----------------
    bOK = Not pVBPInfo Is Nothing
    If bOK = False Then
        Err.Raise vbObjectError + 1000, , "VBP Info object not instanced."
    Else
        bOK = pVBPInfo.Initialised
        If bOK = False Then
            Err.Raise vbObjectError + 1000, , "VBP Info object not initialised."
        End If
    End If
    ' vbpinfo object ok -------------------------------------------------
    Set moVBPInfo = pVBPInfo
    ' grab vbpinfo attribute delimiter ----------------------------------
    mAttributeDelimiter = moVBPInfo.AttributeDelimiter
    ' set flag for writing code module rtf ------------------------------
    mLoadRTF = pWriteRTF
    ' call method to load project codeinfo array ------------------------
    pLoadCodeFiles bOK, sErrMsg
    ' -------------------------------------------------------------------
    If bOK = False Then Err.Raise vbObjectError + 1000, , "Unable to Load Code Files:" & vbNewLine & sErrMsg
        
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

Private Sub pLoadCodeFiles(Optional ByRef pOK As Boolean = False, _
                           Optional ByRef pErrMsg As String = vbNullString)
Attribute pLoadCodeFiles.VB_Description = "attempts to load all code info objects for all source files found in the VBP."
' ... attempts to load all code info objects for all source files found in the VBP.
' ... in doing so it also tries to parse each source file for members and encode
' ... the declarations and code sections of the source file to RTF and save it to disc.

Dim xMemberArray As StringArray
Dim saFilesArray As StringArray
Dim xCodeInfo As CodeInfo
Dim xString As StringWorker

Dim lngFilesCount As Long
Dim lngFileLoop As Long
Dim lngFileIndex As Long
Dim lngMemberLoop As Long
Dim lngMemberCount As Long

Dim tDataInfo As DataInfo
Dim tMemberInfo As MemberInfo
Dim tQuickMember As QuickMemberInfo
Dim iMemberID As Long

Dim sCurrentFile As String
Dim sTmpFolder As String
Dim sTmpFileName As String
Dim sTmpProject As String
Dim sTmp As String
Dim sTmpDesc As String
Dim sFileName As String
Dim sSourceName As String

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    mVBPFilesLoaded = False
    ' -------------------------------------------------------------------
    ' ... bit of validation.
    If moVBPInfo Is Nothing Then Err.Raise vbObjectError + 1000, , "VBP Info not available."
    If moVBPInfo.Initialised = False Then Err.Raise vbObjectError + 1000, , "VBP Info not Initialised."
        
    ' -------------------------------------------------------------------
    sTmpProject = moVBPInfo.ProjectName
    
    sTmpFolder = moVBPInfo.FilePath & "\" & sTmpProject & "RTF"
    
    If Dir$(sTmpFolder, vbDirectory) = "" Then
        MkDir sTmpFolder
    End If
    
    ' -------------------------------------------------------------------
    ' ... set the files data string array locally.
    Set saFilesArray = moVBPInfo.FilesData
    lngFilesCount = saFilesArray.Count
    
    If lngFilesCount = 0 Then Err.Raise vbObjectError + 1000, , "No Files to process in VBP."
    ' -------------------------------------------------------------------
    ' ... resize CodeInfo and CodeRTF arrays.
    ReDim moCodeInfoArray(lngFilesCount - 1)
    ReDim mswCodeRTFArray(lngFilesCount - 1)
    ' -------------------------------------------------------------------
    mCodeFileCount = lngFilesCount
    
    ' -------------------------------------------------------------------
    ' ... loop through the files data from the vbp
    ' ... and create a new code info instance for each source file.
    For lngFileLoop = 1 To lngFilesCount
        ' -------------------------------------------------------------------
        ' ... add file name and details to project file names / details data with index to array item.
        
        msaFileNamesData.AddItemString saFilesArray(lngFileLoop) & "#" & lngFileIndex
        ' -------------------------------------------------------------------
        ParseDataInfoItem saFilesArray, lngFileLoop, tDataInfo
        ' -------------------------------------------------------------------
        ' ... get the file name of the item to read.
        sCurrentFile = tDataInfo.ExtraInfo
        
        RaiseEvent ProcessFile(lngFileLoop, lngFilesCount, "Processing: " & sCurrentFile)
        
        Set xCodeInfo = New CodeInfo
        Set xString = New StringWorker
        
        ' -------------------------------------------------------------------
        ' ... read the file and parse it for members.
        xCodeInfo.ReadCodeFile sCurrentFile
        ' -------------------------------------------------------------------
        ' ... calling Declarations below, forces the CodeInfo instance
        ' ... to parse its declarations for contants, apis, types, enumerators ...
        ' ... these are not automatically read.
        xCodeInfo.Declarations
        
        ' -------------------------------------------------------------------
        ' ... grab source code name and full file name and path details.
        sFileName = xCodeInfo.FileName
        sSourceName = xCodeInfo.Name
        
        ' -------------------------------------------------------------------
        ' ... read the code info's MembersStringArray property to access
        ' ... its member names.
        Set xMemberArray = xCodeInfo.MembersStringArray
        lngMemberCount = xMemberArray.Count
        
        For lngMemberLoop = 1 To lngMemberCount ' ... note: StringArray is Base 1.
            
            
            ' -------------------------------------------------------------------
            ' ... adding name of source object and its file name to the member data
            ' ... for easy identification of member parent and source file.
            sTmp = xMemberArray(lngMemberLoop) & "|" & sSourceName & "|" & sFileName
            ' -------------------------------------------------------------------
            ParseMemberInfoItem xMemberArray, lngMemberLoop, tMemberInfo
            
            iMemberID = tMemberInfo.Index
            
            tQuickMember = xCodeInfo.QuickMember(iMemberID)
            sTmpDesc = tQuickMember.Attribute
            ReplaceChar sTmpDesc, "|", "!"
            ReplaceChars sTmpDesc, "VB_Description = ", ""
            RemoveQuotes sTmpDesc
            sTmp = sTmp & "|" & sTmpDesc
            ' -------------------------------------------------------------------
            msaMembersData.AddItemString sTmp
            
        Next lngMemberLoop
        
        Set xMemberArray = Nothing
        
        sTmp = vbNullString
        sFileName = vbNullString
        sSourceName = vbNullString
        
        If mLoadRTF Then
            ' -------------------------------------------------------------------
            ' ... create an RTF file
            sTmpFileName = xCodeInfo.Name & ".rtf"
            ' -------------------------------------------------------------------
            xString = BuildRTFString(xCodeInfo.GetDecsAndCode)
    '        xString.ToFile sTmpFolder & "\" & sTmpFileName
            ' -------------------------------------------------------------------
            ' ... new properties to stringworker.
            xString.Name = sTmpFolder & "\" & sTmpFileName ' xCodeInfo.Name
            xString.Tag = xCodeInfo.FileName
        End If
        
        ' -------------------------------------------------------------------
        Set moCodeInfoArray(lngFileIndex) = xCodeInfo
        Set mswCodeRTFArray(lngFileIndex) = xString
        ' -------------------------------------------------------------------
        
        lngFileIndex = lngFileIndex + 1
        ' -------------------------------------------------------------------
        
        Set xCodeInfo = Nothing
        Set xString = Nothing
        
        sTmpFileName = vbNullString
        sCurrentFile = vbNullString
        
    Next lngFileLoop
    
    mUBoundCodeArray = UBound(moCodeInfoArray)
    RaiseEvent MembersAndRTFArrays(moCodeInfoArray, mswCodeRTFArray)
    
    ' -------------------------------------------------------------------
    mVBPFilesLoaded = True
    ' -------------------------------------------------------------------
    pOK = True
    pErrMsg = vbNullString
    
ResumeError:
    
    On Error Resume Next
    
    If msaMembersData.Count Then
        ' -------------------------------------------------------------------
        ' ... sort the member names array by member name.
        msaMembersData.Sort
    End If
    
    lngFileIndex = 0&
    lngFilesCount = 0&
    lngFileLoop = 0&
    
    sCurrentFile = vbNullString
    sTmpFileName = vbNullString
    
    Set saFilesArray = Nothing
    
Exit Sub

ErrHan:

    pOK = False
    pErrMsg = Err.Description
    Debug.Print "VBProject.pLoadCodeFiles.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "release resources."
' release resources
    pRelease
End Sub