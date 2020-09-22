VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ------------------------------------------------------------
' Name:         Project
' Purpose:      container for all source files in a project.
' Author:       Dave Carter.
' Date:         Sunday 15 January 2012
' ------------------------------------------------------------

Option Explicit

Private moVBP As VBPInfo
Private mxCodeFiles() As CodeInfo
Private mInitialised As Boolean
Private mbDerivedMembers As Boolean
Private mbDerivedAPIs As Boolean
Private miSourceCount As Long


' -------------------------------------------------------------------

Event ProcessFile(ByVal ItemIndex As Long, ByVal MaxItems As Long, ByVal ItemName As String)
Attribute ProcessFile.VB_Description = "Event saying a source file is being processed, giving its index in the vbp, the no. of source files in the vbp and the name of the source code file."
Event ProcessMember(ByVal MemberIndex As Long, ByVal MaxMembers As Long, ByVal MemberName As String)
Attribute ProcessMember.VB_Description = "Event, declaring the member of a source code file is being processed, with its member position and name and the total no. of members in the source code file."

Private mAttributeDelimiter As String
Private msaMembersData As StringArray           ' list of project's members
Private moAPILibs As StringArray

' -------------------------------------------------------------------

Property Get MembersStringArray() As StringArray
    Set MembersStringArray = New StringArray
    If mbDerivedMembers = False Then pGenerateMembersStringArray
    If Not msaMembersData Is Nothing Then Set MembersStringArray = msaMembersData
End Property

Property Get APIsStringArray() As StringArray
    Set APIsStringArray = New StringArray
    If mbDerivedAPIs = False Then pGenerateAPIReport
    If Not moAPILibs Is Nothing Then Set APIsStringArray = moAPILibs
End Property

Private Sub pGenerateAPIReport()

Dim xMemberArray As StringArray
Dim lngMemberCount As Long
Dim lngMemberLoop As Long
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
Dim iLoop As Long

    On Error GoTo ErrHan:
        
    If mbDerivedAPIs Then Exit Sub
    Set moAPILibs = New StringArray
    If mInitialised = False Then Exit Sub
    If miSourceCount <= 0 Then Exit Sub
    
    For iLoop = 1 To miSourceCount
        With mxCodeFiles(iLoop)
'            .Declarations
            RaiseEvent ProcessFile(iLoop, miSourceCount, .Name)
            Set xMemberArray = .APIStringArray
            lngMemberCount = xMemberArray.Count
            For lngMemberLoop = 1 To lngMemberCount
                sTmp = xMemberArray(lngMemberLoop)
                RaiseEvent ProcessMember(lngMemberLoop, lngMemberCount, "[API] " & sTmp)
                moAPILibs.AddItemString sTmp
                sTmp = vbNullString
            Next lngMemberLoop
            Set xMemberArray = Nothing
        End With
    Next iLoop
    mbDerivedAPIs = True

ResumeError:
    On Error GoTo 0
    
Exit Sub
ErrHan:

    Debug.Print "frmAPIReport.pGenerateAPIReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pGenerateAPIReport:

Private Sub pGenerateMembersStringArray()
Dim xMemberArray As StringArray
Dim lngMemberLoop As Long
Dim lngMemberCount As Long
Dim tMemberInfo As MemberInfo
Dim tQuickMember As QuickMemberInfo
Dim iMemberID As Long
Dim sTmp As String
Dim sTmpDesc As String
Dim sFileName As String
Dim sSourceName As String
Dim iLoop As Long
Dim sCodeType As String

    If mbDerivedMembers Then Exit Sub
    Set msaMembersData = New StringArray
    If mInitialised = False Then Exit Sub
    If miSourceCount <= 0 Then Exit Sub
    For iLoop = 1 To miSourceCount
        With mxCodeFiles(iLoop)
            sFileName = .FileName
            sSourceName = .Name
            sCodeType = .MemberType
            RaiseEvent ProcessFile(iLoop, miSourceCount, "[Members] " & sSourceName)
            ' -------------------------------------------------------------------
            ' ... read the code info's MembersStringArray property to access
            ' ... its member names.
            Set xMemberArray = .MembersStringArray
            lngMemberCount = xMemberArray.Count
            ' -------------------------------------------------------------------
            For lngMemberLoop = 1 To lngMemberCount ' ... note: StringArray is Base 1.
                ' -------------------------------------------------------------------
                ' ... adding name of source object and its file name to the member data
                ' ... for easy identification of member parent and source file.
                sTmp = xMemberArray(lngMemberLoop) & mAttributeDelimiter & sSourceName & mAttributeDelimiter & sFileName & mAttributeDelimiter & " " & mAttributeDelimiter & sCodeType
                xMemberArray(lngMemberLoop) = sTmp
                ' -------------------------------------------------------------------
                ParseMemberInfoItem xMemberArray, lngMemberLoop, tMemberInfo
                ' -------------------------------------------------------------------
                iMemberID = tMemberInfo.Index
                tMemberInfo.ParentTypeString = sCodeType & " "
                ' -------------------------------------------------------------------
                tQuickMember = .QuickMember(iMemberID)
                RaiseEvent ProcessMember(lngMemberLoop, lngMemberCount, "[Member] " & tQuickMember.Name)
                sTmpDesc = tQuickMember.Attribute
                ReplaceChar sTmpDesc, mAttributeDelimiter, "!"
                ReplaceChars sTmpDesc, "VB_Description = ", ""
                RemoveQuotes sTmpDesc
                sTmp = sTmp & mAttributeDelimiter & sTmpDesc
                ' -------------------------------------------------------------------
                msaMembersData.AddItemString sTmp
                ' -------------------------------------------------------------------
            Next lngMemberLoop
            sTmp = vbNullString
            Set xMemberArray = Nothing
        End With
    Next iLoop
    mbDerivedMembers = True
End Sub

Sub CreateSourceFilesReports()
Dim iLoop As Long
Dim sTmp As String
    If mInitialised = False Then Exit Sub
    If miSourceCount <= 0 Then Exit Sub
    For iLoop = 1 To miSourceCount
        RaiseEvent ProcessFile(iLoop, miSourceCount, "[Report] " & mxCodeFiles(iLoop).Name)
        sTmp = mxCodeFiles(iLoop).QuickCodeReport
        sTmp = vbNullString
    Next iLoop
End Sub

Sub CreateSourceFilesRTF()
Dim iLoop As Long
Dim sTmp As String
    If mInitialised = False Then Exit Sub
    If miSourceCount <= 0 Then Exit Sub
    For iLoop = 1 To miSourceCount
        RaiseEvent ProcessFile(iLoop, miSourceCount, "[RTF] " & mxCodeFiles(iLoop).Name)
        sTmp = mxCodeFiles(iLoop).TheTextAsRTF
        sTmp = vbNullString
    Next iLoop
End Sub

Sub ParseSourceFiles()
Dim iLoop As Long
    If mInitialised = False Then Exit Sub
    If miSourceCount <= 0 Then Exit Sub
    For iLoop = 1 To miSourceCount
        RaiseEvent ProcessFile(iLoop, miSourceCount, "[Parse] " & mxCodeFiles(iLoop).Name)
        mxCodeFiles(iLoop).Declarations
    Next iLoop
End Sub

Sub LoadFromVBPInfo(pVBPInfo As VBPInfo)
    pInit
    Set moVBP = pVBPInfo
    pPostLoad
End Sub

Sub LoadByProjectFileName(pVBPFileName As String)
    pInit
    moVBP.ReadVBP pVBPFileName
    pPostLoad
End Sub

Property Get Project() As VBPInfo
    If mInitialised = False Then Exit Property
    Set Project = moVBP
End Property

Property Get SourceCodeObjectByFileName(pFileName As String) As CodeInfo
Dim iLoop As Long
    If mInitialised = False Then Exit Property
    If miSourceCount <= 0 Then Exit Property
    For iLoop = 1 To miSourceCount
        If mxCodeFiles(iLoop).FileName = pFileName Then
            Set SourceCodeObjectByFileName = mxCodeFiles(iLoop)
            Exit For
        End If
    Next iLoop
End Property

Private Sub pPostLoad()
    mInitialised = moVBP.Initialised
    If mInitialised Then
        mAttributeDelimiter = moVBP.AttributeDelimiter
        pLoadSourceFiles
    End If
End Sub

Private Sub pLoadSourceFiles()
Dim iLoop As Long
Dim xFilesData As StringArray
Dim sFile As String
Dim xTmp As StringArray
    If mInitialised = False Then Exit Sub
    miSourceCount = moVBP.FilesData.Count
    If miSourceCount Then
        ReDim moCodeFiles(miSourceCount)
        ReDim mxCodeFiles(miSourceCount)
        Set xFilesData = moVBP.FilesData
        For iLoop = 1 To miSourceCount
            Set xTmp = xFilesData.ItemAsStringArray(iLoop, mAttributeDelimiter)
            sFile = xTmp(4) ' file name
            Set xTmp = Nothing
            Set mxCodeFiles(iLoop) = New CodeInfo
            RaiseEvent ProcessFile(iLoop, miSourceCount, "[Reading] " & sFile)
            mxCodeFiles(iLoop).ReadCodeFile sFile
        Next iLoop
    End If
    Set xFilesData = Nothing
End Sub

Private Sub pInit()
    pRelease
    Set moVBP = New VBPInfo
    Set msaMembersData = New StringArray
End Sub

Private Sub pRelease()
    Set moVBP = Nothing
    Set msaMembersData = Nothing
    Set moAPILibs = Nothing
    Erase mxCodeFiles
    miSourceCount = 0
    mbDerivedMembers = False
    mbDerivedAPIs = False
    mAttributeDelimiter = vbNullString
End Sub

Private Sub Class_Terminate()
    pRelease
End Sub
