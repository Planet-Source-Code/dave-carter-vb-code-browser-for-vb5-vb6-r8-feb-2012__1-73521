VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class to generate a report on a project in plain text."
' what?
'  a class to write text reports on a Visual Basic (5/6) Project.
'  a bit half-arsed (i just got some bad news).
' why?
'  because it could be helpful.
' when?
'
' how?
'
' who?
'  d.c.

Option Explicit

Private mInitialised As Boolean
Private moVBPInfo As VBPInfo
Private mAttributeDelimiter As String
Private Const cPad As Long = 3
Private Const cDist As Long = 36

Public Sub Clear()
Attribute Clear.VB_Description = "Clears resources and resets module fields to empty."

    On Error GoTo ErrHan:

    mInitialised = False
    Set moVBPInfo = Nothing
    
Exit Sub
ErrHan:

    Debug.Print "VBProjTree.Clear", Err.Number, Err.Description
    Err.Clear
    Resume Next
    
End Sub ' ... Clear.

Public Sub Init(ByRef pVBPInfo As VBPInfo, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute Init.VB_Description = "Initialise the vbp search class with a vbpinfo instance."

'... Parameters.
'    R__ pVBPInfo: VBPInfo           ' ... A VBPInfo instance loaded with data.

Dim bOK As Boolean
Dim sErrMsg As String

    On Error GoTo ErrHan:
    
    bOK = Not pVBPInfo Is Nothing
    If bOK = False Then
        Err.Raise vbObjectError + 1000, , "VBP Info object not instanced."
    Else
        bOK = pVBPInfo.Initialised
        If bOK = False Then
            Err.Raise vbObjectError + 1000, , "VBP Info object not initialised."
        End If
    End If
    
    If mInitialised Then Clear
    
    Set moVBPInfo = pVBPInfo
    mAttributeDelimiter = moVBPInfo.AttributeDelimiter
    
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
    Debug.Print "ProjectReport.Init", Err.Number, Err.Description
    Resume ErrResume:


End Sub ' ... Init.

Public Sub GenerateAPIReport(prReportString As String, Optional pType As Long = 0)

Dim oTmpA As StringArray
Dim oTmpAPI As StringArray
Dim oCodeInfo As CodeInfo
Dim xString As SBuilder ' StringWorker
Dim sRepString As String
Dim sTmp As String
Dim lngLoop As Long
Dim lngCount As Long
Dim lngTCount As Long
Dim lngAPILoop As Long
Dim tDataInfo As DataInfo
Dim lngFileLoop As Long
Dim oOutput As StringArray
Dim lngPos As Long
Dim lngFound As Long

    On Error GoTo ErrHan:
    
    Set xString = New SBuilder ' StringWorker
    Set oOutput = New StringArray
    oOutput.DuplicatesAllowed = False
    oOutput.Sortable = True
    
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
                    
                    xString.AppendAsLine modStrings.PadStrings("", oCodeInfo.Name & " * " & lngTCount, 6, 2)
                    
                    For lngAPILoop = 1 To lngTCount
                        
                        lngPos = lngPos + 1
                        
                        sRepString = oTmpAPI(lngAPILoop)
                        
                        lngFound = InStr(1, sRepString, "|")
                        If lngFound > 0 Then sRepString = Left$(sRepString, lngFound - 1)
                        
                        oOutput.AddItemString sRepString
                        
                        sRepString = modStrings.PadStrings(Format$(lngPos, "0000"), sRepString, 6, 2)
                        xString.AppendAsLine sRepString
                        
                    Next lngAPILoop
                    
                    xString.AppendAsLine ""
                    
                End If
                
                Set oCodeInfo = Nothing
                
            Next lngLoop
        End If
        
    Next lngFileLoop

ResumeError:

    If Not xString Is Nothing Then
        xString.PrefixWith "No. of Items: " & lngPos & vbNewLine, , True
        prReportString = xString
        Set xString = Nothing
    End If
    
    If Not oOutput Is Nothing Then
        If pType = 1 Then
            oOutput.Sort
            ' -------------------------------------------------------------------
            ' ... use new stringarray to index sorted items for numbering.
            Set oTmpAPI = New StringArray
            lngTCount = oOutput.Count
            sTmp = "No. of items: " & lngTCount
            oTmpAPI.AddItemString sTmp
            oTmpAPI.AddItemString ""
            
            If lngTCount > 0 Then
                For lngAPILoop = 1 To lngTCount
                    sTmp = oOutput(lngAPILoop)
                    sRepString = modStrings.PadStrings(Format$(lngAPILoop, "0000"), sTmp, 6, 2)
                    oTmpAPI.AddItemString sRepString
                Next lngAPILoop
            End If
            prReportString = oTmpAPI.ToString("", , , True)
        End If
        Set oOutput = Nothing
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
    sRepString = vbNullString
    
Exit Sub

ErrHan:

    Debug.Print "ProjectReport.GenerateAPIReport.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Public Sub GenerateVBPReport(prReportString As String, _
                    Optional pOK As Boolean = False, _
                    Optional pErrMsg As String = vbNullString)

Dim oCodeInfo As CodeInfo
Dim xString As SBuilder ' StringWorker
Dim sRepString As String
Dim sTmp As String
Dim oTmpA As StringArray
Dim lngLoop As Long
Dim lngCount As Long
Dim lngTCount As Long
Dim tDataInfo As DataInfo
Dim lngFileLoop As Long
Dim lngFileSize As Long
Dim lngTFileSize As Long
Dim sFileSize As String
'Const cSmallCDist As Long = 18
Dim lngComLines As Long
Dim lngEmptyLines As Long
Dim lngAttLines As Long
Dim lngActLines As Long
Dim lngSubs As Long
Dim lngFuncs As Long
Dim lngProps As Long
Dim xTmpString As SBuilder ' StringWorker
Dim sTmpFolder As String
Dim sVBPFolder As String

Dim lngMemberDecLines As Long   ' ... v7/8
Dim lngVarLineCount As Long
Dim lngConstLineCount As Long
Dim lngAPILineCount As Long
Dim lngEventLineCount As Long
Dim lngTypeLineCount As Long
Dim lngEnumLineCount As Long
Dim lngCNTRLStatLineCount As Long
Dim lngDecLineCount As Long
Dim lngNonExecLineCount As Long
Dim lngTotalLines As Long       ' ... v7/8
Dim lngHdrLines As Long ' v8

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    
    If mInitialised = False Then Err.Raise vbObjectError + 1000, "", "The Project Class (a VBPInfo instance) has not been initialised yet for this function to work:" & vbNewLine & "Please call Init with a valid VBPInfo instance and then try again."
    ' -------------------------------------------------------------------
    Set xString = New SBuilder ' StringWorker
    Set xTmpString = New SBuilder ' StringWorker
    
    sVBPFolder = moVBPInfo.FilePath ' v6 added, to help provide relative paths to source files rather than whole path and name.
    
    sTmp = modStrings.PadStrings("VB 5/6 Project Report:", Format$(Date, "short date"), cDist, cPad)
    xString.AppendAsLine sTmp
    xString.AppendAsLine ""
    
    sTmp = modStrings.PadStrings("Project Title:", moVBPInfo.Title, cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Project Name:", moVBPInfo.ProjectName, cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Project Type:", moVBPInfo.ProjectType, cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Project Description:", moVBPInfo.Description, cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Project Version:", moVBPInfo.Version, cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    
    sTmp = modStrings.PadStrings("File Name:", moVBPInfo.FileName, cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("File Location:", moVBPInfo.FilePath, cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("File Size:", modGeneral.GetFileLength(moVBPInfo.FileNameAndPath), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    xString.AppendAsLine ""
    
    sTmp = modStrings.PadStrings("No. of References:", Format$(moVBPInfo.CountReferences, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("No. of Components:", Format$(moVBPInfo.CountComponents, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    
    lngCount = moVBPInfo.CountForms
    lngTCount = lngTCount + lngCount
    sTmp = modStrings.PadStrings("No. of Forms:", Format$(lngCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    lngCount = moVBPInfo.CountModules
    lngTCount = lngTCount + lngCount
    sTmp = modStrings.PadStrings("No. of Modules:", Format$(lngCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    lngCount = moVBPInfo.CountClasses
    lngTCount = lngTCount + lngCount
    sTmp = modStrings.PadStrings("No. of Classes:", Format$(lngCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    lngCount = moVBPInfo.CountUserControls
    If lngCount > 0 Then
        lngTCount = lngTCount + lngCount
        sTmp = modStrings.PadStrings("No. of User Controls:", Format$(lngCount, cNumFormat), cDist, cPad)
        xString.AppendAsLine sTmp
    End If
    
    ' -------------------------------------------------------------------
    ' v6, added, property pages & designers (e.g. data environment / data report).
    lngCount = moVBPInfo.CountPropertyPages
    If lngCount > 0 Then
        lngTCount = lngTCount + lngCount
        sTmp = modStrings.PadStrings("No. of Property Pages:", Format$(lngCount, cNumFormat), cDist, cPad)
        xString.AppendAsLine sTmp
    End If
    
    lngCount = moVBPInfo.CountDesigners
    If lngCount > 0 Then
        lngTCount = lngTCount + lngCount
        sTmp = modStrings.PadStrings("No. of Designers:", Format$(lngCount, cNumFormat), cDist, cPad)
        xString.AppendAsLine sTmp
    End If
    ' -------------------------------------------------------------------
    xString.AppendAsLine ""
    
    sTmp = modStrings.PadStrings("Total Code Files:", Format$(lngTCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    
    For lngFileLoop = 1 To 8 ' 6 ' ... v6, incremented to 8 to include property pages & designers.
        
        ' ... this should be 3 to 8 because references and components are not parsed by CodeInfo
        ' ... loop thu' source code files, summing line property values
        
        Select Case lngFileLoop
            Case 1: sTmp = "References":        Set oTmpA = moVBPInfo.ReferencesData
            Case 2: sTmp = "Components":        Set oTmpA = moVBPInfo.ComponentsData
            Case 3: sTmp = "Forms":             Set oTmpA = moVBPInfo.FormsData
            Case 4: sTmp = "Modules":           Set oTmpA = moVBPInfo.ModulesData
            Case 5: sTmp = "Classes":           Set oTmpA = moVBPInfo.ClassesData
            Case 6: sTmp = "User Controls":     Set oTmpA = moVBPInfo.UserControlsData
            Case 7: sTmp = "Property Pages":    Set oTmpA = moVBPInfo.PropertyPagesData ' v6.
            Case 8: sTmp = "Designers":         Set oTmpA = moVBPInfo.DesignersData ' v6.
        End Select
        
        lngCount = oTmpA.Count
        
        If lngCount > 0 Then
            
            xTmpString.AppendAsLine modStrings.PadStrings(sTmp, Format$(lngCount, cNumFormat), cDist, cPad)
            
            For lngLoop = 1 To lngCount
            
                modGeneral.ParseDataInfoItem oTmpA, lngLoop, tDataInfo, mAttributeDelimiter
                
                sFileSize = modGeneral.GetFileLength(tDataInfo.ExtraInfo, lngFileSize)
                lngTFileSize = CLng(CLng(lngTFileSize) + CLng(lngFileSize))
                                
                sTmpFolder = tDataInfo.ExtraInfo
                If Left$(sTmpFolder, Len(sVBPFolder)) = sVBPFolder Then
                    sTmpFolder = ".." & Mid$(sTmpFolder, Len(sVBPFolder) + 1)
                End If
                
                sTmp = modStrings.PadStrings("  " & tDataInfo.Name, sTmpFolder & IIf(lngFileLoop > 2, "     Size: " & sFileSize, ""), cDist, cPad)
                xTmpString.AppendAsLine sTmp
                
                If lngFileLoop > 2 Then
                
                    ' ... not interested in references and components in this bit
                    ' ... because they can't be read by a CodeInfo Object.
                    
                    Set oCodeInfo = New CodeInfo
                    oCodeInfo.ReadCodeFile tDataInfo.ExtraInfo
                    
                    If oCodeInfo.Initialised Then
                        
                        oCodeInfo.Declarations ' ... read the declarations section of the source file as well, v7/8.
                        
                        lngComLines = lngComLines + oCodeInfo.CountCommentLines
                        lngEmptyLines = lngEmptyLines + oCodeInfo.EmptyLineCount
                        lngAttLines = lngAttLines + oCodeInfo.CountAttributeLines
                        
                        ' -------------------------------------------------------------------
                        ' ... v7/8, number of lines given to method / member declarations including End Statement.
                        lngMemberDecLines = lngMemberDecLines + oCodeInfo.MemberDeclarationLineCount
                        lngAPILineCount = lngAPILineCount + oCodeInfo.APIDeclarationLineCount
                        lngEventLineCount = lngEventLineCount + oCodeInfo.EventDeclarationLineCount
                        lngTypeLineCount = lngTypeLineCount + oCodeInfo.TypeDeclarationLineCount
                        lngEnumLineCount = lngEnumLineCount + oCodeInfo.EnumDeclarationLineCount
                        lngCNTRLStatLineCount = lngCNTRLStatLineCount + oCodeInfo.ControlStatementCount
                        
                        lngConstLineCount = lngConstLineCount + oCodeInfo.ConstDeclarationLineCount
                        lngVarLineCount = lngVarLineCount + oCodeInfo.VarDeclarationLineCount
                        lngDecLineCount = lngDecLineCount + oCodeInfo.DeclarativeLineCount
                        lngNonExecLineCount = lngNonExecLineCount + oCodeInfo.NonExecutingLineCount
                        lngActLines = lngActLines + oCodeInfo.ActualLineCount
                        lngTotalLines = lngTotalLines + oCodeInfo.GetLineCount
                        
                        lngSubs = lngSubs + oCodeInfo.CountOfSubs
                        lngFuncs = lngFuncs + oCodeInfo.CountOfFunctions
                        lngProps = lngProps + oCodeInfo.CountOfProperties
                        ' -------------------------------------------------------------------
                        lngHdrLines = lngHdrLines + oCodeInfo.GetHeaderLineCount ' v8
                    
                    End If
                
                End If
                
                Set oCodeInfo = Nothing
                
            Next lngLoop
        
        End If
        
        xTmpString.AppendAsLine ""
        
    Next lngFileLoop
    
    sTmp = modStrings.PadStrings("Total Bytes:", Format$(lngTFileSize, cNumFormat) & " Bytes", cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    
    sTmp = modStrings.PadStrings("Number of Subs:", Format$(lngSubs, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Number of Functions:", Format$(lngFuncs, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Number of Properties:", Format$(lngProps, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    
    sTmp = modStrings.PadStrings("Total Number of Methods/Members:", Format$(lngProps + lngFuncs + lngSubs, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine ""
    xString.AppendAsLine "Line Stats in Editor:*"
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Line Count:", Format$(lngTotalLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    ' -------------------------------------------------------------------
    ' ... v7/8
    sTmp = modStrings.PadStrings(" Method Declaration Lines:", Format$(lngMemberDecLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Variable Declaration Lines:", Format$(lngVarLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Constant Declaration Lines:", Format$(lngConstLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" API Declaration Lines:", Format$(lngAPILineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Event Declaration Lines:", Format$(lngEventLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Type Declaration Lines:", Format$(lngTypeLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
        
    sTmp = modStrings.PadStrings(" Enum Declaration Lines:", Format$(lngEnumLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine "-------------------------------------------------------------------"
        
    sTmp = modStrings.PadStrings("Total Declarative Lines:", Format$(lngDecLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine "-------------------------------------------------------------------"
        
    sTmp = modStrings.PadStrings("Total Control Statement Lines:**", Format$(lngCNTRLStatLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp ' & " [lines with statements, not number of control statements]"
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings(" Commented Lines:", Format$(lngComLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Empty Lines:", Format$(lngEmptyLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" VB Attribute Lines:", Format$(lngAttLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
        
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Non-Executing Lines:", Format$(lngNonExecLineCount, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Code Line Count:", Format$(lngActLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    xString.AppendAsLine ""
    
    xString.AppendAsLine "Header Line Stats:***" ' v8
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Header Line Count:", Format$(lngHdrLines, cNumFormat), cDist, cPad)
    xString.AppendAsLine sTmp
    
    xString.AppendAsLine "-------------------------------------------------------------------"
    
    xString.AppendAsLine ""
    
    If xTmpString.Length > 0 Then
        xString.AppendAsLine xTmpString.TheString
    End If
    ' -------------------------------------------------------------------
    
    xString.AppendAsLine ""
'    xString.AppendAsLine ""
    
    sTmp = "*   Editor Line Stats. Values and Totals, are, at best, approximate and no frx files have been included in total file size calculations."
    xString.AppendAsLine sTmp
    
    sTmp = "**  Statement Control Line Stats is the number of lines with statements, not the number of control statements."
    xString.AppendAsLine sTmp
    
    sTmp = "*** Header Line Stats refers to the lines making up the top section of code files hidden from the Editor."
    xString.AppendAsLine sTmp
    sTmp = "    (this section includes the vb attributes of the file and any designer related declarations such as Controls and Menus)."
    xString.AppendAsLine sTmp ' v8
    xString.AppendAsLine ""
'
'    sTmp = "For entertainment only;"
'    xString.AppendAsLine sTmp ' v8
'    sTmp = "No Warrantees, guarantees or nothing."
'    xString.AppendAsLine sTmp ' v8
'    sTmp = "Stats may be arguably apocryphal, indefensibly inaccurate or just wildy wrong"
'    xString.AppendAsLine sTmp ' v8
'    xString.AppendAsLine ""
    
    sTmp = "For entertainment only;"
    xString.AppendAsLine sTmp ' v8
    sTmp = "No Warrantees, No Guarantees, No Nothing."
    xString.AppendAsLine sTmp ' v8
    sTmp = "Stats may be perplexingly apocryphal, wildy inaccurate or simply, wrong"
    xString.AppendAsLine sTmp ' v8
    xString.AppendAsLine ""
    
    sRepString = xString
    
    pOK = True
    ' -------------------------------------------------------------------
    
ResumeError:

    If Not xString Is Nothing Then
        Set xString = Nothing
    End If
    
    If Not xTmpString Is Nothing Then
        Set xTmpString = Nothing
    End If
    
    If Not oCodeInfo Is Nothing Then
        Set oCodeInfo = Nothing
    End If
    
    If Not oTmpA Is Nothing Then
        Set oTmpA = Nothing
    End If
    
    prReportString = sRepString

    sRepString = vbNullString
    sTmp = vbNullString
    
Exit Sub
        
ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    sRepString = vbNullString
    Debug.Print "ProjectReport.GenerateVBPReport.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:

End Sub

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a boolean describing whether the tree view has been set up etc."
    
    Initialised = mInitialised
    
End Property

Private Sub Class_Terminate()
    Clear
End Sub
