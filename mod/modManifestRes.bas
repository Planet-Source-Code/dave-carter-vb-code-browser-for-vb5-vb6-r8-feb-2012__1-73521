Attribute VB_Name = "modManifestRes"
Attribute VB_Description = "A module to help create and compile a manifest file into a compiled resource file."
Option Explicit

' Requires:
'           cOptions
'           StringWorker
'           VBPInfo
'           modGeneral
'           modFileName
'           modStrings

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Sub GenerateManifestResource(ByVal pvVBPFullFileName As String, _
                                    ByRef prTheResourceFile As String, _
                           Optional ByVal pIncludeHookText As Boolean = True, _
                           Optional ByRef pOK As Boolean = False, _
                           Optional ByRef pErrMsg As String = vbNullString)
Attribute GenerateManifestResource.VB_Description = "Attempts to create a compiled Resource File containing a Manifest describing a VB Project."

' Sub:             GenerateManifestResource
' Description:     Attempts to create a compiled resource file containing a manifest describing a vb project.

    
' ... open a new vbpinfo
' ... read vbp info
'   ... command32 name / if not project name
'   ... project description.

' ... create manifest file ensuring length divisible by 4
' ... make sure not to overwrite an existing one.

' ... run rc to create the resource file
' ... make sure not to overwrite existing one.

' ... return the name of the resource file created

'... Parameters.
'    V__ pvVBPFullFileName: String   ' ... The full path and name of a VBP file.
'    R__ prTheResourceFile: String   ' ... Returns the full path and name of the Compiled Resource file.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
Dim xOptions As cOptions
Dim tFileNameInf As FileNameInfo
Dim sRequiredFile As String
Dim lngFSize As Long
Dim xString As SBuilder ' StringWorker
Dim oVBP As VBPInfo
Dim sExeName As String
Dim sRCFileName As String
Dim sResName As String

Dim sTmpFolder As String

Dim sDesc As String
Dim sCommand As String
Dim sRCExe As String
Dim sManifestPath As String
Dim sWinTmpFolder As String
Dim iAnswer As VbMsgBoxResult

' Note:
'   Busker's version, 1.

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    ' ... make sure we have RC.exe and RCDLL.dll
    Set xOptions = New cOptions
    xOptions.Read bOK, sErrMsg
    ' -------------------------------------------------------------------
    ' ... make sure the vbp is valid.
    If bOK = True Then
        sRequiredFile = pvVBPFullFileName
        modGeneral.GetFileLength sRequiredFile, lngFSize
        bOK = lngFSize > 0
    Else
        Err.Raise vbObjectError + 1000, , "The VBP could not be located."
    End If
    ' -------------------------------------------------------------------
    ' ... grab options path to VB6 IDE directory (probably need to include VB5 IDE dir as well).
    modFileName.ParseFileNameEx xOptions.PathToVB6, tFileNameInf
    ' -------------------------------------------------------------------
    ' ... check we have the RC.exe file.
    If bOK = True Then
        sRequiredFile = tFileNameInf.Path & "\Wizards\RC.exe"
        sRCExe = sRequiredFile
        modGeneral.GetFileLength sRequiredFile, lngFSize
        bOK = lngFSize > 0
    Else
        Err.Raise vbObjectError + 1000, , "RC.exe could not be located."
    End If
    ' -------------------------------------------------------------------
    ' ... check we have the RCDLL.dll file.
    If bOK = True Then
        sRequiredFile = tFileNameInf.Path & "\Wizards\RCDLL.dll"
        modGeneral.GetFileLength sRequiredFile, lngFSize
        bOK = lngFSize > 0
    Else
        Err.Raise vbObjectError + 1000, , "RCDLL.dll could not be located."
    End If
    ' -------------------------------------------------------------------
    
    If bOK = True Then
        ' -------------------------------------------------------------------
        ' ... post validation, processing.
        
        Set oVBP = New VBPInfo
        
        oVBP.ReadVBP pvVBPFullFileName, bOK, sErrMsg
        
        If bOK = True Then
            
            If oVBP.HasResource Then
                Err.Raise vbObjectError + 1000, , "Project: " & oVBP.Title & ", already has a Resource File" & _
                    vbNewLine & oVBP.ResFileNameAndPath & vbNewLine & "Code Browser will not attempt to overwrite this resource reference!"
                
            End If
            ' -------------------------------------------------------------------
            
            sWinTmpFolder = modGeneral.GetTempPath
            sWinTmpFolder = Left$(sWinTmpFolder, 3)
            
            sTmpFolder = oVBP.FilePath & "\tmpRes"
            If Dir$(sTmpFolder, vbDirectory) = "" Then
                MkDir sTmpFolder
            Else
                iAnswer = MsgBox("The Temp Resource Folder " & vbNewLine & sTmpFolder & vbNewLine & "already exists." & vbNewLine & "Would you like to continue and overwrite any files?", vbQuestion + vbYesNo, "Manifest Resource Compiler")
                If iAnswer = vbNo Then Err.Raise vbObjectError + 1000, , "Temp Resource Folder exists and operation was cancelled."
            End If
            
            ' -------------------------------------------------------------------
            sExeName = oVBP.ExeName32
            If Len(sExeName) = 0 Then
                sExeName = oVBP.ProjectName & ".exe"
            End If
            sManifestPath = sExeName
            
            sDesc = oVBP.Title
            ' -------------------------------------------------------------------
            ' ... manifest file text.
            Set xString = New SBuilder ' StringWorker
            
            xString.AppendAsLine "<?xml version=" & WrapInQuoteChars("1.0") & " encoding=" & WrapInQuoteChars("UTF-8") & " standalone=" & WrapInQuoteChars("yes") & " ?>"
            xString.AppendAsLine "<assembly xmlns=" & WrapInQuoteChars("urn:schemas-microsoft-com:asm.v1") & " manifestVersion=" & WrapInQuoteChars("1.0") & ">"
            xString.AppendAsLine "    <assemblyIdentity"
            xString.AppendAsLine "        version = " & WrapInQuoteChars("1.0.0.0")
            xString.AppendAsLine "        processorArchitecture = " & WrapInQuoteChars("X86")
            xString.AppendAsLine "        name = " & WrapInQuoteChars(sExeName)
            xString.AppendAsLine "        type=" & WrapInQuoteChars("win32") & " />"
            xString.AppendAsLine "    <description>" & sDesc & "</description>"
            xString.AppendAsLine "    <dependency>"
            xString.AppendAsLine "        <dependentAssembly>"
            xString.AppendAsLine "            <assemblyIdentity"
            xString.AppendAsLine "                type=" & WrapInQuoteChars("win32")
            xString.AppendAsLine "                name = " & WrapInQuoteChars("Microsoft.Windows.Common-Controls")
            xString.AppendAsLine "                version = " & WrapInQuoteChars("6.0.0.0")
            xString.AppendAsLine "                processorArchitecture = " & WrapInQuoteChars("X86")
            xString.AppendAsLine "                publicKeyToken = " & WrapInQuoteChars("6595b64144ccf1df")
            xString.AppendAsLine "                language=" & WrapInQuoteChars("*") & " />"
            xString.AppendAsLine "        </dependentAssembly>"
            xString.AppendAsLine "    </dependency>"
            xString.Append "</assembly>"
            
            lngFSize = xString.Length
            
            Select Case lngFSize Mod 4
                ' ... ensure manifest text output length is divisible by 4
                Case 1, 2, 3
                    xString.Append Space$(4 - (lngFSize Mod 4))
            End Select
            ' -------------------------------------------------------------------
                        
            sExeName = sWinTmpFolder & sExeName & ".manifest"
            sRCFileName = sWinTmpFolder & oVBP.ProjectName & ".rc"

            sResName = sWinTmpFolder & oVBP.ProjectName & ".res"
            
            ' -------------------------------------------------------------------
            ' ... write manifest file to temp folder.
            xString.WriteToFile sExeName, True
            ' -------------------------------------------------------------------
            xString.DeleteAll
            ' -------------------------------------------------------------------
            ' ... write rc file to temp folder.
            sManifestPath = sWinTmpFolder & sManifestPath & ".manifest"
            
            xString.TheString = "1   24  " & sManifestPath & vbNewLine
            xString.WriteToFile sRCFileName
            ' -------------------------------------------------------------------
            ' ... generate command string for RC.Exe to process and create resource file.
            sCommand = modStrings.WrapInQuoteChars(sRCExe) & " /r /fo " & modStrings.WrapInQuoteChars(sResName) & " " & modStrings.WrapInQuoteChars(sRCFileName)
                        
            lngFSize = Shell(sCommand, vbHide)
            
            ' -------------------------------------------------------------------
            ' ... let the program catch up with itself, could still be writing otherwise.
            DoEvents
            Sleep 1000
            DoEvents
            ' -------------------------------------------------------------------
            
            ' -------------------------------------------------------------------
            ' ... cleaning up temp files.
            Kill sExeName
            Kill sRCFileName
            ' -------------------------------------------------------------------
            
            If Dir$(sResName) <> "" Then
                
                If pIncludeHookText Then
                    
                    ' -------------------------------------------------------------------
                    ' ... text file with initcommoncontrolsVB declaration and method.
                    ' -------------------------------------------------------------------
                    
                    xString.DeleteAll
                    ' -------------------------------------------------------------------
                    
                    xString.AppendAsLine ""
                    xString.AppendAsLine "Private Type tagInitCommonControlsEx"
                    xString.AppendAsLine "   lngSize As Long"
                    xString.AppendAsLine "   lngICC As Long"
                    xString.AppendAsLine "End Type"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "Private Declare Sub InitCommonControls Lib " & modStrings.WrapInQuoteChars("comctl32.dll") & " ()"
                    xString.AppendAsLine "Private Declare Function InitCommonControlsEx Lib " & modStrings.WrapInQuoteChars("comctl32.dll") & " (iccex As tagInitCommonControlsEx) As Boolean"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "Private Const ICC_USEREX_CLASSES = &H200"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "Public Sub InitCommonControlsVB()"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "Dim X As tagInitCommonControlsEx"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "    On Error Resume Next"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "    With X"
                    xString.AppendAsLine "        .lngSize = LenB(X)"
                    xString.AppendAsLine "        .lngICC = ICC_USEREX_CLASSES"
                    xString.AppendAsLine "    End With"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "    InitCommonControlsEx X"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "    If Err Then"
                    xString.AppendAsLine "        Err.Clear"
                    xString.AppendAsLine "        On Error GoTo 0"
                    xString.AppendAsLine "        InitCommonControls"
                    xString.AppendAsLine "    End If"
                    xString.AppendAsLine ""
                    xString.AppendAsLine "End Sub"
                    xString.AppendAsLine ""
                    
                    ' -------------------------------------------------------------------
                    
                    xString.WriteToFile sTmpFolder & "\InitCommonControlsVB.txt"
                
                End If
                
                ' -------------------------------------------------------------------
                
                FileCopy sResName, sTmpFolder & "\" & oVBP.ProjectName & ".res"
                Kill sResName
                
                ' -------------------------------------------------------------------
                sResName = sTmpFolder & "\" & oVBP.ProjectName & ".res"
                
                ' -------------------------------------------------------------------
                prTheResourceFile = sResName
                ' -------------------------------------------------------------------
                
                MsgBox "Manifest Compiled into Resource file:" & vbNewLine & sResName, vbInformation, "Resource Created"
                
            Else
            
                MsgBox "Manifest Not Compiled into Resource file:" & vbNewLine & sResName, vbExclamation, "Resource Not Created"
            
            End If
            
        End If
    
    End If
    
    ' -------------------------------------------------------------------
    
    Let sErrMsg = vbNullString
    Let bOK = True

ErrResume:
    
    If Not xString Is Nothing Then
        Set xString = Nothing
    End If
    
    If Not oVBP Is Nothing Then
        Set oVBP = Nothing
    End If
    
    If Not xOptions Is Nothing Then
        Set xOptions = Nothing
    End If
    
    Let pErrMsg = sErrMsg
    Let pOK = bOK

Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "modManifestRes.GenerateManifestResource", Err.Number, Err.Description
    Resume ErrResume:

End Sub ' ... GenerateManifestResource.
