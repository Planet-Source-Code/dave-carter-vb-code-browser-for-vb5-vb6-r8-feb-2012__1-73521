VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DevHelpGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A draft class to generate list of all public interfaces in a project in HTML as quick reference."

' what?
'  a class to generate a HTML Help File geared toward
'  developer type help for a project.
' why?
'  for the sake of easier documentation and information
'  about the interfaces and members of a project.
' when?
'  at any stage of development but more likely come release time.
' how?
'  Declare a new DevHelpGen, pass it a project file (vbp)
'  and go.
' who?
'  d.c.

' Note:
'       The html help project is intended to act as a draft.
'       It will work to a template and some blanks such as
'       parameter and function descriptions and uses will
'       need to be filled in to add the extra value required
'       to make the document of any real use.
'       The idea is that information may be more easily shared
'       in a disparate and dynamic development environment.
'       The important information is in the public interfaces
'       where sub systems can be plugged in to add extra functionality.
'       This will draft a template of your public interfaces
'       and give you a chance to describe how to use them
'       so other developers can get to work with them more quickly.
'       The html help project is not intended to be a general
'       user help document and is not suitable in such a role.

Option Explicit

' ... interface requirements.

' create HTML   Topic
'               Table of Contents
'               Context Help references
'               Style Sheet

' compile a html help project

' [delete content files on completed compilation]

' name / location / overwrite files

' ... may be looking at two classes
'   ... one that manages a vbp and
'   ... one that manages a class.

' ... Topic
'   ... a topic will identify the name and type of an Interface and its parent class.
'   ... parameters will be listed with name, type in/out and description
'   ... return values will be noted with name, type and description
'   ... remarks placeholder
'   ... example placeholder

'   ... tags will be placed within the topic file to identify
'   ... the various sections so creating an after editor could be easy.
'   ... the tags will be within the <!-- and --> of the html comment.

' ... Class/Source File
'   ... a class will have a name, type, description and belong to a project.
'   ... it will group public, private and friend members together to
'   ... provide a sorted, tabled list of members, parameters, return values and description
'   ... this list shall act as a table of contents for accessing the member's detail page
'   ... a class will have a general context of use or scenarios where it can be useful

' ... HTML Help Project
'   ... an HTML Help Project will provide a formal draft help manual
'   ... for other developers to learn how to use your components etc.
'   ... it shall create and organise the neccessary files and attempt
'   ... to compile them into a distributable CHM file.

' ... Editor
'   ... an editor can be written that will allow the topic files to be updated
'   ... and there will be an instruction to recompile the help project after editing.

' -------------------------------------------------------------------

' ... given a vbp, this class can then run through each of the component source files
' ... and then generate a set of topic files for each, when this is done the table
' ... of contents can be written and the project compiled using the HTML Help program.

' -------------------------------------------------------------------

Private mInitialised As Boolean
Private moVBPInfo As VBPInfo
Private mAttributeDelimiter As String

Private Sub pGenerateTopicFiles(Optional ByRef pOK As Boolean = False, _
                                Optional pErrMsg As String = vbNullString)
Attribute pGenerateTopicFiles.VB_Description = "draft routine to read vbp files and generate html help topics for their methods / members."


' ... draft routine to read vbp files and generate html help topics for their methods / members.

Dim bDefaultProperty As Boolean

Dim lngSourceCount As Long
Dim lngFileLoop As Long
Dim lngFileLen As Long
Dim lngMemberCount As Long
Dim lngMemberLoop As Long
Dim lngMemParamCount As Long
Dim lngParamLoop As Long

Dim sCurrentFile As String
Dim sFileLen As String
Dim sMemberParamString As String
Dim sParam As String
Dim sParamOutput As String
Dim sSyntax As String
Dim sHelpTopicFile As String
Dim sTmpFolder As String
Dim sTmpDesc As String

Dim sMemberParamsArray() As String

Dim xQuickMemInfo As QuickMemberInfo
Dim xMemInfo As MemberInfo
Dim xFileInfo As FileNameInfo
Dim tDataInfo As DataInfo
Dim tParamInfo As ParamInfo

Dim xCodeInfo As CodeInfo
Dim soTmpFiles As StringArray
Dim soTmpMember As StringArray
Dim soMemberProperties As StringArray
Dim xString As SBuilder ' StringWorker
Dim xHeader As SBuilder ' StringWorker

Dim bUniqueProperty As Boolean
Dim bParentIsObject As Boolean
Dim sParentObject As String
Dim xConstsString As SBuilder ' StringWorker
Dim soConstsArray As StringArray
Dim xAPIsString As SBuilder ' StringWorker
Dim soAPIsArray As StringArray

    On Error GoTo ErrHan:
    
    Screen.MousePointer = vbHourglass
    
    ' -------------------------------------------------------------------
    If mInitialised = False Then Err.Raise vbObjectError + 1000, , "Class has not been successfully initialised"
    If moVBPInfo Is Nothing Then Err.Raise vbObjectError + 1000, , "VBP Info Object not instanced"
    If moVBPInfo.Initialised = False Then Err.Raise vbObjectError + 1000, , "VBP Info Object not initialised"
    ' -------------------------------------------------------------------
    lngSourceCount = moVBPInfo.FilesData.Count
    If lngSourceCount = 0 Then Err.Raise vbObjectError + 1000, , "No Files to process"
    ' -------------------------------------------------------------------
    sTmpFolder = moVBPInfo.FilePath & "\hhlp"
    If Dir$(sTmpFolder, vbDirectory) = "" Then MkDir sTmpFolder
    ' -------------------------------------------------------------------
    ' ... grab the vbp info's files data string array for faster access to its data.
    Set soTmpFiles = moVBPInfo.FilesData
    ' -------------------------------------------------------------------
    
    Set xConstsString = New SBuilder ' StringWorker
    Set xAPIsString = New SBuilder ' StringWorker
    ' set up and begin the index page header ----------------------------
    Set xHeader = New SBuilder ' StringWorker
    xHeader.AppendAsLine "<HTML><TITLE>" & pHTMLEncode(moVBPInfo.ProjectName) & " Dev. Dictionary: Index</TITLE><BODY><PRE>"
    xHeader.AppendAsLine "<B>" & pHTMLEncode(moVBPInfo.ProjectName) & "</B>"
    xHeader.AppendAsLine pHTMLEncode(moVBPInfo.Description)
    xHeader.AppendAsLine "<HR>"
    xHeader.AppendAsLine "Project Interfaces"
    xHeader.AppendAsLine "<HR>"
    xHeader.AppendAsLine "<A HREF='" & moVBPInfo.ProjectName & "_Constants.htm'>Constants</A>, " & "<A HREF='" & moVBPInfo.ProjectName & "_APIs.htm'>APIs</A>"
    xHeader.AppendAsLine "<BR><HR>Source File Members<HR>" ', True, True
    xHeader.AppendAsLine "<TABLE>"
    ' -------------------------------------------------------------------
    ' ... loop thru the files data from the vbp
    For lngFileLoop = 1 To lngSourceCount
        
        ' -------------------------------------------------------------------
        ' ... parse each item as a DataInfo structure and get the file name referenced in ExtraInfo member.
        modGeneral.ParseDataInfoItem soTmpFiles, lngFileLoop, tDataInfo
        ' -------------------------------------------------------------------
        
        sCurrentFile = tDataInfo.ExtraInfo
        ' -------------------------------------------------------------------
        
        If Len(sCurrentFile) > 0 Then
            
            ' ... see that the file name references a valid file.
            sFileLen = modGeneral.GetFileLength(sCurrentFile, lngFileLen)
            ' -------------------------------------------------------------------
            If lngFileLen > 0 Then
                ' -------------------------------------------------------------------
                ' ... parse the file name into a FileInfo structure for easy access to file path related info.
                modFileName.ParseFileNameEx sCurrentFile, xFileInfo
                ' -------------------------------------------------------------------
                sHelpTopicFile = sTmpFolder & "\" & xFileInfo.FileName & ".htm"
                ' ... set up a new CodeInfo instance and read the source file.
                Set xCodeInfo = New CodeInfo
                xCodeInfo.ReadCodeFile sCurrentFile
                xCodeInfo.Declarations
                ' -------------------------------------------------------------------
                xConstsString.AppendAsLine xCodeInfo.ConstsString
                xAPIsString.AppendAsLine xCodeInfo.APIStringArray.ToString("", "", , True)
                ' -------------------------------------------------------------------
                bParentIsObject = CBool(xCodeInfo.MemberType <> "Module")
                sParentObject = IIf(bParentIsObject, "<I>Object</I>.", "")
                ' set up and begin html page for source code object -----------------
                Set xString = New SBuilder ' StringWorker
                xString.AppendAsLine "<HTML><TITLE>" & pHTMLEncode(moVBPInfo.ProjectName & " Dev. Dictionary: " & xCodeInfo.MemberType & ": " & xCodeInfo.Name) & "</TITLE><BODY><PRE>"
                xString.AppendAsLine "Project: " & pHTMLEncode(moVBPInfo.ProjectName & ";  Type: " & moVBPInfo.ProjectType & "; v" & moVBPInfo.Version)
                xString.AppendAsLine "Public / Friend Members of:"
                xString.AppendAsLine xCodeInfo.MemberType & ": <B>" & pHTMLEncode(xCodeInfo.Name) & "</B>"
                ' -------------------------------------------------------------------
                If Len(xCodeInfo.Description) > 0 Then _
                    xString.AppendAsLine pHTMLEncode(xCodeInfo.Description)
                ' if code info thinks its ok, press on ------------------------------
                If xCodeInfo.Initialised Then
                    ' -------------------------------------------------------------------
                    lngMemberCount = xCodeInfo.MemberCount
                    ' -------------------------------------------------------------------
                    If lngMemberCount > 0 Then
                        ' update the index page to reference this new html page -------------
                        xHeader.AppendAsLine "<TR>"
                        xHeader.AppendAsLine "<TD><A HREF='" & pHTMLEncode(xFileInfo.FileName & ".htm") & "'><B>" & pHTMLEncode(xCodeInfo.Name) & "</B></A></TD><TD>" & pHTMLEncode(xCodeInfo.Description) & "</TD></TR>"
                        ' -------------------------------------------------------------------
                        Set soMemberProperties = New StringArray
                        ' -------------------------------------------------------------------
                        Set soTmpMember = xCodeInfo.MembersStringArray
                        ' loop thru' members ------------------------------------------------
                        For lngMemberLoop = 1 To lngMemberCount
                            ' parse member into a MemberInfo Type -------------------------------
                            modGeneral.ParseMemberInfoItem soTmpMember, lngMemberLoop, xMemInfo
                            ' -------------------------------------------------------------------
                            With xMemInfo
                                
                                ' -------------------------------------------------------------------
                                ' ... CodeInfo::QuickMember.
                                xQuickMemInfo = xCodeInfo.QuickMember(.Index)
                                ' -------------------------------------------------------------------
                                If xQuickMemInfo.Accessor = 2 Then ' ... if private then skip
                                    GoTo SkipPrivateMember:
                                    
                                End If
                                ' -------------------------------------------------------------------
                                bUniqueProperty = True
                                ' try handle properties so write only once --------------------------
                                If xQuickMemInfo.Type = 3 Then
                                    If soMemberProperties.Exists(xQuickMemInfo.Name) = False Then
                                        soMemberProperties.AddItemString xQuickMemInfo.Name
                                    Else
                                        bUniqueProperty = False
                                        xString.AppendAsLine ""
                                        xString.AppendAsLine "<!-- SyntaxReverseProperty -->"
                                        xString.AppendAsLine "Reverse Property"
                                        If InStr(1, xQuickMemInfo.Declaration, "Property Set ") > 0 Then
                                            xString.AppendAsLine "Syntax: Set " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B> = x" & xQuickMemInfo.ValueType
                                        Else
                                            xString.AppendAsLine "Syntax: " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B> = x" & xQuickMemInfo.ValueType ' & " ' ... Let/Set"
                                        End If
                                        xString.AppendAsLine ""
                                    End If
                                End If
                                ' -------------------------------------------------------------------
                                ' ... Parameters String from Declaration.
                                sMemberParamString = xCodeInfo.ExtractParamsString(xQuickMemInfo.Declaration)
                                ' -------------------------------------------------------------------
                                ' ... Parameters Array from Parameters String, above.
                                xCodeInfo.GetDecParams sMemberParamString, sMemberParamsArray, lngMemParamCount
                                ' -------------------------------------------------------------------
                                ' ... Start of output.
                                If bUniqueProperty Then
                                    xString.AppendAsLine "<HR>"
                                    xString.AppendAsLine "<!-- Name -->"
                                    xString.AppendAsLine "<B>" & pHTMLEncode(.Name) & "</B>" & " <!-- Type -->" & pHTMLEncode(.TypeAsString)
                                End If
                                ' -------------------------------------------------------------------
                                bDefaultProperty = False
                                ' -------------------------------------------------------------------
                                If bUniqueProperty Then
                                    sTmpDesc = xQuickMemInfo.Attribute
                                    If InStr(1, sTmpDesc, "VB_UserMemId = 0") > 0 Then
                                        bDefaultProperty = True
                                        sTmpDesc = modStrings.Replace(sTmpDesc, "VB_UserMemId = 0", "")
                                        If Right$(sTmpDesc, 1) = "|" Then sTmpDesc = Left$(sTmpDesc, Len(sTmpDesc) - 1)
                                    End If
                                    ' ... Descriptive Attribute, not necessarily a description
                                    ' ... could be any of the Attribute Tags VB adds inside a method.
                                    If Len(sTmpDesc) > 0 Then
                                        If Left$(sTmpDesc, 17) = "VB_Description = " Then
                                            sTmpDesc = Mid$(sTmpDesc, 18)
                                        End If
                                        RemoveQuotes sTmpDesc
                                        xString.AppendAsLine "<!-- Description -->" 'Description"
                                        xString.AppendAsLine pHTMLEncode(sTmpDesc)
                                        If bDefaultProperty Then
                                            xString.AppendAsLine "<!-- DefaultProperty -->*** <i>Default Property</i> ***"
                                        End If
                                        xString.AppendAsLine ""
                                    End If
                                End If
                                
                                ' -------------------------------------------------------------------
                                ' ... Return Data Type (if any).
                                If Len(.ValueType) > 0 Then
                                    xString.AppendAsLine "<!-- Returns -->"
                                    xString.AppendAsLine "<B>Returns</B>: " & .ValueType
                                End If
                                
                                ' -------------------------------------------------------------------
                                sSyntax = vbNullString
                                sParamOutput = vbNullString
                                
                                ' -------------------------------------------------------------------
                                If lngMemParamCount > 0 Then
                                    
                                    ' ... Parameters for output and syntax rule.
                                    For lngParamLoop = 0 To lngMemParamCount - 1
                                    
                                        ' -------------------------------------------------------------------
                                        ' ... loop through parameter Members Array
                                        sParam = sMemberParamsArray(lngParamLoop)
                                        
                                        ' -------------------------------------------------------------------
                                        ' ... Parameter to ParamInfo Structure.
                                        modGeneral.ParseParamInfoItem sParam, tParamInfo
                                        ' -------------------------------------------------------------------
                                        With tParamInfo
                                            If Len(sParamOutput) > 0 Then sParamOutput = sParamOutput & vbNewLine
                                            sParamOutput = sParamOutput & pHTMLEncode(IIf(.IsByRef, "In/Out: ", "In:     ") & PadStrings(.Name, .Type, 20, 2))
                                            If .IsOptional Then
                                                If Len(.DefaultValue) > 0 Then
                                                    sParamOutput = sParamOutput & vbTab & " [default = " & .DefaultValue & "]"
                                                End If
                                            End If
                                            If Len(sSyntax) Then sSyntax = sSyntax & ", "
                                            sSyntax = sSyntax & IIf(.IsOptional, "[<I>" & pHTMLEncode(.Name) & "</I>]", pHTMLEncode(.Name))
                                        End With
                                    Next lngParamLoop
                                End If
                                
                                ' -------------------------------------------------------------------
                                If bUniqueProperty Then
                                    xString.AppendAsLine "<!-- Syntax -->"
                                    If Len(sSyntax) > 0 Then
                                        ' ... Output Syntax.
                                        If xQuickMemInfo.Type > 1 Then
                                            'xString.AppendAsLine "Syntax: x" & xQuickMemInfo.ValueType & " = <B>" & pHTMLEncode(.Name) & "</B>(" & sSyntax & ")"
                                            If xQuickMemInfo.Type = 2 Then ' function
                                                xString.AppendAsLine "Syntax: x = " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B>(" & sSyntax & ")"
                                            Else ' property
                                                xString.AppendAsLine "Syntax: x = " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B>(" & sSyntax & ")"
                                            End If
                                        Else ' sub
                                            xString.AppendAsLine "Syntax: " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B>(" & sSyntax & ")"
                                        End If
                                    Else
                                        If xQuickMemInfo.Type > 1 Then
                                            If xQuickMemInfo.Type = 2 Then
                                                xString.AppendAsLine "Syntax: x = " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B>"
                                            Else
                                                xString.AppendAsLine "Syntax: x = " & sParentObject & "<B>" & pHTMLEncode(.Name) & "</B>"
                                            End If
                                        End If
                                    End If
                                End If
                                ' -------------------------------------------------------------------
                                If Len(sParamOutput) > 0 Then
                                    ' ... Output Parameters.
                                    xString.AppendAsLine "<!-- Parameters -->"
                                    xString.AppendAsLine "<B>Parameters</B>"
                                    xString.AppendAsLine sParamOutput
                                End If
                                
                                ' -------------------------------------------------------------------
                                ' ... Remarks.
                                If bUniqueProperty Then
                                    xString.AppendAsLine "<!-- Remarks -->"
                                    xString.AppendAsLine "<B>Remarks</B>" & vbNewLine
                                Else
                                    xString.AppendAsLine ""
                                End If
                                
                            End With
SkipMemberAdded:
SkipPrivateMember:
                        Next lngMemberLoop
                        
                        xString.AppendAsLine "<HR>"
                        xString.AppendAsLine "</PRE></BODY></HTML>"
                        xString.WriteToFile sHelpTopicFile
                    
                    End If
                    
                Else
                
                    Debug.Print sCurrentFile & ": file not successfully read by CodeInfo instance."
                    
                End If
                
            Else
            
                Debug.Print sCurrentFile & ": no file found at location specified."
                
            End If
            
            
        End If
        
    Next lngFileLoop
    
    xHeader.AppendAsLine "</TABLE>"
    xHeader.AppendAsLine "</HTML></BODY></PRE>"
    xHeader.WriteToFile sTmpFolder & "\" & moVBPInfo.ProjectName & "_HelpIndex.htm"
    
    OpenWebPage sTmpFolder & "\" & moVBPInfo.ProjectName & "_HelpIndex.htm"
    
    pOK = True
    
ResumeError:

    Set soConstsArray = New StringArray
    
    soConstsArray.Sortable = True
    soConstsArray.FromString xConstsString.TheString, vbCrLf, True, False
    
    If soConstsArray.Count Then
    
        soConstsArray.Sort
        
        xConstsString.TheString = soConstsArray.ToString("", "", , True)
        xConstsString.PrefixWith "Module Level Constant Declarations found in " & moVBPInfo.ProjectName & vbNewLine, , True
        xConstsString.TheString = modEncode.BuildHTMLString(xConstsString.TheString, , , , moVBPInfo.ProjectName & " Dev. Dictionary: Module Level Constants")
        xConstsString.WriteToFile sTmpFolder & "\" & moVBPInfo.ProjectName & "_Constants.htm"
        
    End If
    
    Set soAPIsArray = New StringArray
    
    soAPIsArray.Sortable = True
    soAPIsArray.FromString xAPIsString.TheString, vbCrLf, True, False
    
    If soAPIsArray.Count Then
        
        soAPIsArray.Sort
        
        xAPIsString.TheString = soAPIsArray.ToString("", "", , True)
        xAPIsString.PrefixWith "API Declarations found in " & moVBPInfo.ProjectName & vbNewLine, , True
        xAPIsString.TheString = modEncode.BuildHTMLString(xAPIsString.TheString, , , , moVBPInfo.ProjectName & " Dev. Dictionary: API Declarations")
        xAPIsString.WriteToFile sTmpFolder & "\" & moVBPInfo.ProjectName & "_APIs.htm"
    
    End If
    
    Screen.MousePointer = vbDefault
    
Exit Sub

ErrHan:

    pOK = False
    pErrMsg = Err.Description
    
    Debug.Print "DevHelpGen.pGenerateTopicFiles.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears resources and resets module fields to empty."

    On Error GoTo ErrHan:

    mInitialised = False
    Set moVBPInfo = Nothing
    
Exit Sub
ErrHan:

    Debug.Print "DevHelpGen.Clear", Err.Number, Err.Description
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
    pGenerateTopicFiles
    
Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "DevHelpGen.Init", Err.Number, Err.Description
    Resume ErrResume:

End Sub ' ... Init.

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a boolean describing whether the tree view has been set up etc."
    
    Initialised = mInitialised
    
End Property

Private Sub Class_Terminate()
    Clear
End Sub

Private Function pHTMLEncode(ByRef pTheString As String) As String
Attribute pHTMLEncode.VB_Description = "draft attempt to encode a string for html."

' ... draft attempt to encode a string for html.

Dim bString() As Byte
Dim sWord As String
Dim sTmp As String
Dim lngLoop As Long
Dim lngChar As Long

    On Error GoTo ErrHan:
    If Len(pTheString) = 0 Then
        pHTMLEncode = pTheString
        Exit Function
    End If
    
    bString = pTheString
    
'    sTmp = String$(Len(pTheString) * 5, 0)
    
    For lngLoop = 0 To UBound(bString) Step 2
        
        lngChar = bString(lngLoop)
        
        Select Case lngChar
            Case 34, 38, 39, 60, 62 ' ( " & ' < > ) HTML Reserved Chars.
                sWord = sWord & "&#" & CStr(lngChar) & ";"
            Case Else
                sWord = sWord & Chr$(lngChar)
        End Select
        
        sTmp = sTmp & sWord
        
        sWord = vbNullString
        
    Next lngLoop
    
'    sTmp = Left$(sTmp, InStr(1, sTmp, Chr$(0)) - 1)
    ' -------------------------------------------------------------------
    pTheString = sTmp
    pHTMLEncode = sTmp
    ' -------------------------------------------------------------------

ResumeError:
    
    Erase bString
    sTmp = vbNullString
    lngChar = 0&
    lngLoop = 0&

Exit Function

ErrHan:

    Debug.Print "DevHelpGen.pHTMLEncode.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function