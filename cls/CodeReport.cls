VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class to display Code File information."
' what?
'  basic class to provide simple reporting of source file
' why?
'  handy to have around.
' when?
'  when there is a valid CodeInfo instance and a need to get an interface or report
'  from it.
' how?
'  this is a little obscure:
'  essentially one should instance a new CodeReport
'  then set its CodeReader property to a valid CodeInfo object
'  to be ready to generate a report or an interface.
'  Call GenerateInterface to return the interface description
'  Call ReportString to return a StringWorker instance with the report data.
'  e.g.
'   Dim x As CodeReport
'
'       Set x = New CodeReport
'       Set x.CodeReader = mCodeInfo
'       ' ... Print the interface to the code info object.
'       Debug.Print x.GenerateInterface
'
'   Dim xS as StringWorker
'       Set xS = x.ReportString
'       ' ... Print the report string to the code info object.
'       Debug.Print xS.ToString
'       Set xS = Nothing
'
'       Set x = Nothing
' who?
'  d.c.

Option Explicit

' S   pGenerateQuickReport.   ' ... Private method to initialise the Report's Compilation.
' S   pInit.             ' ... Private method to clean up any existing resources before doing stuff..
' S   pRelease.          ' ... Private method to release any existing resources.

' P   CodeReader: CodeInfo.  ' ... Sets the CodeInfo Class Instance from which to create a report.
' P   ReportString: StringWorker.  ' ... Returns a StringWorker with the Report Info.

Private moCodeReader As CodeInfo ' ... private field for property CodeReader.
Private m_ReportString As SBuilder ' StringWorker ' ... private field for property ReportString.
Private mAttributeDelimiter As String
Private Const cPad As Long = 3
Private Const cDist As Long = 36

' Note:
'   Busker's version, 1:
'   Report;
'   Whatever info is desired, prepare it within the pGenerateQuickReport method.
'   I've just made up a few things I thought handy in a report and presented it
'   in whatever way seemed ok at the time.
'   Just add and format as you see fit, use the xString StringWorker dot AppendAsLine method to
'   add new lines to the output stringworker.

Public Function GenerateInterface(Optional ByRef pIncludeAttributes As Boolean = True, _
                                  Optional ByRef pAllMembers As Boolean = False) As String
Attribute GenerateInterface.VB_Description = "Attempts to create an interface for a class e.g. all the exposed member names without their code.  pAllMembers, when True, includes Private as well as Public methods."

Dim xString As SBuilder ' StringWorker
Dim tQInfo As QuickMemberInfo
Dim lngLoop As Long
Dim lngCount As Long
Dim sDec As String
Dim sEndString As String
Dim sAttributes As String
Dim lngAttsLoop As Long
Dim lngAccessor As Long
Dim sAttribs() As String

Const cEnd As String = "End "

    On Error GoTo ErrHan:
        
    If moCodeReader Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Unable to generate interface without a valid CodeInfo object to parse."
    End If

    Let lngCount = moCodeReader.MemberCount
    Set xString = New SBuilder ' StringWorker
    xString.AppendAsLine ""
    xString.AppendAsLine "' Interface based upon " & moCodeReader.Name & IIf(pAllMembers = False, ", Public/Friend", ", All") & " members."
    If Len(moCodeReader.Description) > 0 Then
        xString.AppendAsLine "' " & moCodeReader.Description
    End If
    xString.AppendAsLine ""
    For lngLoop = 0 To lngCount - 1
        tQInfo = moCodeReader.QuickMember(lngLoop)
        lngAccessor = tQInfo.Accessor
        If lngAccessor <> 2 Or pAllMembers = True Then
            sDec = tQInfo.Declaration
            sEndString = cEnd & Choose(tQInfo.Type, "Sub", "Function", "Property") & " ' ... " & tQInfo.Name & "."
            xString.AppendAsLine sDec
            If pIncludeAttributes Then
                sAttributes = tQInfo.Attribute
                If Len(sAttributes) Then
                    modStringArrays.SplitString sAttributes, sAttribs, "|"
                    For lngAttsLoop = 0 To UBound(sAttribs)
                        xString.AppendAsLine "Attribute " & tQInfo.Name & "." & sAttribs(lngAttsLoop)
                    Next lngAttsLoop
                End If
            End If
            xString.AppendAsLine ""
            xString.AppendAsLine sEndString
            xString.AppendAsLine ""
        End If
    Next lngLoop

ResumeError:
    
    On Error GoTo 0
    If Not xString Is Nothing Then
        GenerateInterface = xString
        Set xString = Nothing
    End If
    
Exit Function

ErrHan:

    Debug.Print "CodeReport.GenerateInterface.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Private Sub pGenerateQuickReport(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pGenerateQuickReport.VB_Description = "Private method to initialise the Report's Compilation."

' Sub:             pGenerateQuickReport
' Description:     Private method to initialise the Report's Compilation.

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim lngMCount As Long
Dim lngLoop As Long
Dim lngTmpCount As Long
Dim sTmpLeft As String
Dim sTmpRight As String
Dim sTmp As String

Dim oSubs As StringArray
Dim oFuncs As StringArray
Dim oProps As StringArray
Dim oTmpA As StringArray ' v5
Dim oMembers As StringArray ' v5
Dim lngMemberIndex As Long ' v5
Dim lngMType As Long
Dim tQuickMem As QuickMemberInfo
Dim sTmpDec As String
Dim sTmpValue As String
Dim sTmpPType As String
Dim lngAccessor As Long
Dim sAccessor As String

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
Dim lngComLines As Long
Dim lngEmptyLines As Long
Dim lngAttLines As Long
Dim lngActLines As Long

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    If m_ReportString Is Nothing Then
        Err.Raise vbObjectError + 1000, , "No Report String available to write the report into."
    End If
    ' -------------------------------------------------------------------
    If moCodeReader Is Nothing Then
        Err.Raise vbObjectError + 1000, , "No Code Reader to Read From."
    End If
    ' -------------------------------------------------------------------
        
    lngMCount = moCodeReader.MemberCount
    
    sTmp = modStrings.PadStrings(moCodeReader.MemberType & ":", moCodeReader.Name, 30, 3)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings("Description:", moCodeReader.Description, 30, 3)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = moCodeReader.FileName
    If Len(sTmp) Then
        
        m_ReportString.AppendAsLine ""
        sTmp = modStrings.PadStrings("File Name:", sTmp, 30, 3)
        m_ReportString.AppendAsLine sTmp
        
        sTmp = GetFileLength(moCodeReader.FileName)
        sTmp = modStrings.PadStrings("File Size:", sTmp, 30, 3)
        m_ReportString.AppendAsLine sTmp
        
    End If
    
    m_ReportString.AppendAsLine ""
    
    If lngMCount > 0 Then
        
        sTmp = modStrings.PadStrings("Member Count:", Format$(lngMCount, cNumFormat), 30, 3)
        m_ReportString.AppendAsLine sTmp
    
        lngTmpCount = moCodeReader.CountOfSubs
        If lngTmpCount > 0 Then
            sTmp = modStrings.PadStrings("Subs:", Format$(lngTmpCount, cNumFormat), 30, 3) ' v2
            m_ReportString.AppendAsLine sTmp
        End If
        
        lngTmpCount = moCodeReader.CountOfFunctions
        If lngTmpCount > 0 Then
            sTmp = modStrings.PadStrings("Functions:", Format$(lngTmpCount, cNumFormat), 30, 3) ' v2
            m_ReportString.AppendAsLine sTmp
        End If
        
        lngTmpCount = moCodeReader.CountOfProperties
        If lngTmpCount > 0 Then
            sTmp = modStrings.PadStrings("Properties:", Format$(lngTmpCount, cNumFormat), 30, 3) ' v2
            m_ReportString.AppendAsLine sTmp
        End If
        
        m_ReportString.AppendAsLine ""
        
    End If
    
'    lngTmpCount = moCodeReader.GetLineCount
'    sTmp = modStrings.PadStrings("Line Count:", Format$(lngTmpCount, cNumFormat) & " ... includes attributes.", 30, 3)
'    m_ReportString.AppendAsLine sTmp
'
'    sTmp = modStrings.PadStrings("Commented Lines:", Format$(moCodeReader.CountCommentLines, cNumFormat), 30, 3) ' v2
'    m_ReportString.AppendAsLine sTmp
'
'    sTmp = modStrings.PadStrings("Empty Lines:", Format$(moCodeReader.EmptyLineCount, cNumFormat), 30, 3) ' v2
'    m_ReportString.AppendAsLine sTmp
'
'    sTmp = modStrings.PadStrings("VB Attribute Lines:", Format$(moCodeReader.CountAttributeLines, cNumFormat), 30, 3) ' v2
'    m_ReportString.AppendAsLine sTmp
'
'    sTmp = modStrings.PadStrings("Declarative Lines:", Format$(moCodeReader.DeclarativeLineCount, cNumFormat), 30, 3) ' v2
'    m_ReportString.AppendAsLine sTmp ' v8
'
'    sTmp = modStrings.PadStrings("Actual Lines:", Format$(lngTmpCount - moCodeReader.CountCommentLines - moCodeReader.EmptyLineCount - moCodeReader.CountAttributeLines, cNumFormat), 30, 3) ' v2
'    m_ReportString.AppendAsLine sTmp
'
'    m_ReportString.AppendAsLine ""
'    lngTmpCount = moCodeReader.APIStringArray.Count
'    If lngTmpCount > 0 Then
'        sTmp = modStrings.PadStrings("APIs:", Format$(lngTmpCount, cNumFormat), 30, 3) ' v2
'        m_ReportString.AppendAsLine sTmp
'    End If
'    lngTmpCount = moCodeReader.ConstantsStringArray.Count
'    If lngTmpCount > 0 Then
'        sTmp = modStrings.PadStrings("Constants:", Format$(lngTmpCount, cNumFormat), 30, 3) ' v2
'        m_ReportString.AppendAsLine sTmp
'    End If
    
    
' -------------------------------------------------------------------
' -------------------------------------------------------------------

    lngComLines = lngComLines + moCodeReader.CountCommentLines
    lngEmptyLines = lngEmptyLines + moCodeReader.EmptyLineCount
    lngAttLines = lngAttLines + moCodeReader.CountAttributeLines
    
    ' -------------------------------------------------------------------
    ' ... v7/8, number of lines given to method / member declarations including End Statement.
    lngMemberDecLines = lngMemberDecLines + moCodeReader.MemberDeclarationLineCount
    lngAPILineCount = lngAPILineCount + moCodeReader.APIDeclarationLineCount
    lngEventLineCount = lngEventLineCount + moCodeReader.EventDeclarationLineCount
    lngTypeLineCount = lngTypeLineCount + moCodeReader.TypeDeclarationLineCount
    lngEnumLineCount = lngEnumLineCount + moCodeReader.EnumDeclarationLineCount
    lngCNTRLStatLineCount = lngCNTRLStatLineCount + moCodeReader.ControlStatementCount
    
    lngConstLineCount = lngConstLineCount + moCodeReader.ConstDeclarationLineCount
    lngVarLineCount = lngVarLineCount + moCodeReader.VarDeclarationLineCount
    lngDecLineCount = lngDecLineCount + moCodeReader.DeclarativeLineCount
    lngNonExecLineCount = lngNonExecLineCount + moCodeReader.NonExecutingLineCount
    lngActLines = lngActLines + moCodeReader.ActualLineCount
    lngTotalLines = lngTotalLines + moCodeReader.GetLineCount
    
'    lngSubs = lngSubs + moCodeReader.CountOfSubs
'    lngFuncs = lngFuncs + moCodeReader.CountOfFunctions
'    lngProps = lngProps + moCodeReader.CountOfProperties
    ' -------------------------------------------------------------------
    lngHdrLines = lngHdrLines + moCodeReader.GetHeaderLineCount ' v8


    m_ReportString.AppendAsLine ""
    m_ReportString.AppendAsLine "Line Stats in Editor:*"
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Line Count:", Format$(lngTotalLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    ' -------------------------------------------------------------------
    ' ... v7/8
    sTmp = modStrings.PadStrings(" Method Declaration Lines:", Format$(lngMemberDecLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Variable Declaration Lines:", Format$(lngVarLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Constant Declaration Lines:", Format$(lngConstLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" API Declaration Lines:", Format$(lngAPILineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Event Declaration Lines:", Format$(lngEventLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Type Declaration Lines:", Format$(lngTypeLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
        
    sTmp = modStrings.PadStrings(" Enum Declaration Lines:", Format$(lngEnumLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
        
    sTmp = modStrings.PadStrings("Total Declarative Lines:", Format$(lngDecLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
        
    sTmp = modStrings.PadStrings("Total Control Statement Lines:**", Format$(lngCNTRLStatLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp ' & " [lines with statements, not number of control statements]"
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings(" Commented Lines:", Format$(lngComLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" Empty Lines:", Format$(lngEmptyLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    sTmp = modStrings.PadStrings(" VB Attribute Lines:", Format$(lngAttLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
        
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Non-Executing Lines:", Format$(lngNonExecLineCount, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Code Line Count:", Format$(lngActLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    m_ReportString.AppendAsLine ""
    
    m_ReportString.AppendAsLine "Header Line Stats:***" ' v8
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    sTmp = modStrings.PadStrings("Total Header Line Count:", Format$(lngHdrLines, cNumFormat), cDist, cPad)
    m_ReportString.AppendAsLine sTmp
    
    m_ReportString.AppendAsLine "-------------------------------------------------------------------"
    
    m_ReportString.AppendAsLine ""

' -------------------------------------------------------------------
' -------------------------------------------------------------------

    Set oSubs = New StringArray
    Set oFuncs = New StringArray
    Set oProps = New StringArray
    
    Set oMembers = moCodeReader.MembersStringArray
    
    For lngLoop = 0 To lngMCount - 1
    
        Set oTmpA = oMembers.ItemAsStringArray(lngLoop + 1, "|")
                
        lngMemberIndex = oTmpA.ItemAsNumberValue(2)
        
        tQuickMem = moCodeReader.QuickMember(lngMemberIndex)
        
        lngAccessor = tQuickMem.Accessor
        
        sAccessor = "+  "
        If lngAccessor > 1 Then
            sAccessor = "   "
            If lngAccessor = 3 Then
                sAccessor = "~  "
            End If
        End If
        
        lngMType = tQuickMem.Type
        
        Select Case lngMType
        
            Case 1  ' ... Sub.
                sTmpLeft = sAccessor & tQuickMem.Name & ":"
                sTmpRight = Mid$(tQuickMem.Attribute, 18)
                sTmp = modStrings.PadStrings(sTmpLeft, sTmpRight, 30, 3)
                
                oSubs.AddItemString sTmp
            
            Case 2, 3 ' ... Function, Property respectively.
            
                If lngMType = 2 Then ' ... Function.
                    
                    sTmpLeft = sAccessor & tQuickMem.Name & ": " & tQuickMem.ValueType
                    sTmpRight = Mid$(tQuickMem.Attribute, 18)
                    sTmp = modStrings.PadStrings(sTmpLeft, sTmpRight, 30, 3)
                    
                    oFuncs.AddItemString sTmp
                    
                ElseIf lngMType = 3 Then ' ... Property.
                
                    sTmpValue = vbNullString
                    sTmpDec = tQuickMem.Declaration
                    If InStr(1, sTmpDec, " Get ") > 0 Then
                        sTmpValue = tQuickMem.ValueType
                        sTmpPType = "(Get)"
                    Else
                        If InStr(1, sTmpDec, " Let ") > 0 Then
                            sTmpPType = "(Let)"
                        Else
                            sTmpPType = "(Set)"
                        End If
                    End If
                    
                    If InStr(1, tQuickMem.Attribute, "VB_UserMemId = 0") > 0 Then
                        sAccessor = Left$(sAccessor, 1) & "! "
                    End If
                    
                    sTmpLeft = sAccessor & tQuickMem.Name & sTmpPType & ": " & sTmpValue
                    sTmpRight = Mid$(tQuickMem.Attribute, 18)
                    sTmp = modStrings.PadStrings(sTmpLeft, sTmpRight, 30, 3)
                    
                    oProps.AddItemString sTmp
                
                End If
                
        End Select
        
    Next lngLoop
    
    If oSubs.Count > 0 Then
        m_ReportString.AppendAsLine ""
        m_ReportString.AppendAsLine "== Subs: " & CStr(oSubs.Count) '& " ===="
        m_ReportString.AppendAsLine oSubs.ToString("", "", , True)
    End If
    
    If oFuncs.Count > 0 Then
        m_ReportString.AppendAsLine ""
        m_ReportString.AppendAsLine "== Functions: " & CStr(oFuncs.Count) '& " ===="
        m_ReportString.AppendAsLine oFuncs.ToString("", "", , True)
    End If
    
    If oProps.Count > 0 Then
        m_ReportString.AppendAsLine ""
        m_ReportString.AppendAsLine "== Properties: " & CStr(oProps.Count) '& " ===="
        m_ReportString.AppendAsLine oProps.ToString("", "", , True)
    End If
        
' -------------------------------------------------------------------
' -------------------------------------------------------------------
    m_ReportString.AppendAsLine ""
'    m_ReportString.AppendAsLine ""
    
    sTmp = "*   Editor Line Stats. Values and Totals, are, at best, approximate and no frx files have been included in total file size calculations."
    m_ReportString.AppendAsLine sTmp
    
    sTmp = "**  Statement Control Line Stats is the number of lines with statements, not the number of control statements."
    m_ReportString.AppendAsLine sTmp
    
    sTmp = "*** Header Line Stats refers to the lines making up the top section of code files hidden from the Editor."
    m_ReportString.AppendAsLine sTmp
    sTmp = "    (this section includes the vb attributes of the file and any designer related declarations such as Controls and Menus)."
    m_ReportString.AppendAsLine sTmp ' v8
    m_ReportString.AppendAsLine ""
'
'    sTmp = "For entertainment only;"
'    m_ReportString.AppendAsLine sTmp ' v8
'    sTmp = "No Warrantees, guarantees or nothing."
'    m_ReportString.AppendAsLine sTmp ' v8
'    sTmp = "Stats may be arguably apocryphal, indefensibly inaccurate or just wildy wrong"
'    m_ReportString.AppendAsLine sTmp ' v8
'    m_ReportString.AppendAsLine ""
    
    sTmp = "For entertainment only;"
    m_ReportString.AppendAsLine sTmp ' v8
    sTmp = "No Warrantees, No Guarantees, No Nothing."
    m_ReportString.AppendAsLine sTmp ' v8
    sTmp = "Stats may be perplexingly apocryphal, wildy inaccurate or simply, wrong"
    m_ReportString.AppendAsLine sTmp ' v8
    m_ReportString.AppendAsLine ""

' -------------------------------------------------------------------
' -------------------------------------------------------------------



    pOK = True
    
ErrResume:

    On Error GoTo 0
    
    sTmp = vbNullString
    sTmpLeft = vbNullString
    sTmpRight = vbNullString
    
    Set oSubs = Nothing
    Set oFuncs = Nothing
    Set oProps = Nothing
    
Exit Sub
ErrHan:
    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "CodeReport.pGenerateQuickReport", Err.Number, Err.Description
    Resume ErrResume:

End Sub ' ... pGenerateQuickReport.

Private Sub pRelease(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pRelease.VB_Description = "Private method to release any existing resources."

' Sub:             pRelease
' Description:     Private method to release any existing resources.

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:
    
    If Not m_ReportString Is Nothing Then
        Set m_ReportString = Nothing
    End If
    
    If Not moCodeReader Is Nothing Then
        Set moCodeReader = Nothing
    End If
    
    Let pErrMsg = vbNullString
    Let pOK = True

ErrResume:

Exit Sub
ErrHan:
    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "CodeReport.pRelease", Err.Number, Err.Description
    Resume ErrResume:

End Sub ' ... pRelease.

Private Sub pInit(Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
Attribute pInit.VB_Description = "Private method to clean up any existing resources before doing stuff.."

' Sub:             pInit
' Description:     Private method to clean up any existing resources before doing stuff..

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

    On Error GoTo ErrHan:
    
    pRelease pOK, pErrMsg
    
    If pOK = True Then
        
        Set m_ReportString = New SBuilder ' StringWorker
        Let pErrMsg = vbNullString
        Let pOK = True
    
    End If
ErrResume:

Exit Sub
ErrHan:
    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "CodeReport.pInit", Err.Number, Err.Description
    Resume ErrResume:
End Sub ' ... pInit.

Public Property Get ReportString() As SBuilder ' StringWorker
Attribute ReportString.VB_Description = "Returns a StringWorker with the Report Info."

' Property Get:    ReportString
' Type:            StringWorker.
' Description:     Returns a StringWorker with the Report Info.

Dim bOK As Boolean
Dim sErrMsg As String
    On Error GoTo ErrHan:
    pGenerateQuickReport bOK, sErrMsg
    If bOK = False Then
        Set ReportString = New SBuilder ' StringWorker
    Else
        Set ReportString = m_ReportString
    End If
    
Exit Property
ErrHan:
    Debug.Print "CodeReport.ReportString", Err.Number, Err.Description
End Property ' ... ReportString: StringWorker.

Public Property Set CodeReader(ByRef pNewValue As CodeInfo)
Attribute CodeReader.VB_Description = "Sets the CodeInfo Class Instance from which to create a report."

' Property Set:    CodeReader
' Type:            CodeInfo.
' Description:     Sets the CodeInfo Class Instance from which to create a report.
    
    On Error GoTo ErrHan:
    If Not pNewValue Is Nothing Then
        pInit   ' ... must come ahead of setting the new CodeReader because it releases any current one.
        Set moCodeReader = pNewValue
    End If
Exit Property
ErrHan:
    Debug.Print "CodeReport.CodeReader", Err.Number, Err.Description
End Property ' ... CodeReader: CodeInfo.

Private Sub Class_Terminate()
    pRelease
End Sub
