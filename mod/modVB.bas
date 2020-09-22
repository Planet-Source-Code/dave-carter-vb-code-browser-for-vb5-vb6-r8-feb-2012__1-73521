Attribute VB_Name = "modVB"
Attribute VB_Description = "A module with some vb specific text parsing stuff."

' what?
'  a module dedicated to reading vb information.
' why?
'  portability, these methods will become requisite to higher functions.
' when?
'  reading vb source files from disk.
' how?
'
' who?
'  d.c.

Option Explicit

' Requires
'   modReader.ReadFile
'   modStrings.RemoveQuotes
'   modStrings.InstrRevChar

'Private Enum vbpFileObjTypeEnum
'
'    evbNotSet = -1
'
'    evbModule
'    evbClass
'    evbForm
'    evbUserControl
'
'    evbDateEnvironment
'    evbReport
'
'End Enum

'Private Type vbpFileObjInfo
'
'    Name As String
'    FileName As String
'    RelativePath As String
'    Type As vbpFileObjTypeEnum
'
'End Type

' -------------------------------------------------------------------


Public Type DeclarationInfo

' Structure:   DeclarationInfo.
' Description: A structure to describe various properties of a VB Member's Declaration.

    Name As String                  ' ... The Member's Name.
    Accessor As Long                ' ... The Member's Accessor e.g. Public, Private, Friend.
    AccessorAsString As String      ' ... The Member's Accessor as a string e.g. Public, Private, Friend.
    MemberType As Long              ' ... The Member's Type as a Long e.g. Sub=0, Function=1,Property=2
    MemberTypeAsString As String    ' ... The Member's Type as a String e.g. Sub, Function, Property
    ValueTypeAsString As String     ' ... The Member's Value Data Type as a String e.g. String, Long, Boolean etc
    ParamString As String           ' ... The Member's Declaration Parameters (without parenthesis) as a string.
    PropertyType As Long            ' ... When the Member is a Property, describes if it is a Get=0, Let=1 or Set=2.
    PropertyTypeAsString As String  ' ... When the Member is a Property, describes if it is a Get, Let or Set.
    ReturnsSomething As Boolean     ' ... Describes whether the Declaration Returns anything, True for Yes, False for No.
'    Syntax As String                ' ... Returns a syntax statement for using the Declaration.

End Type ' ... DeclarationInfo.

' -------------------------------------------------------------------

Public Type QuickVBPInfo

' Structure:   QuickVBPInfo.
' Description: A structure to describe various properties from a vb6 vbp file.

    Type As String              ' ... Describes the Type of Program (Standard EXE, ActiveX DLL / EXE or User Control).
    Name As String              ' ... Describes the Name of the VB6 Project.
    Description As String       ' ... Describes the Description of the VB6 Project.
    HelpFile As String          ' ... Describes the Name and Location of the project's Help File.
    
    ' -------------------------------------------------------------------
    MajorVersion As String
    MinorVersion As String
    RevisionVersion As String
    ' -------------------------------------------------------------------
'    ' ... to keep things simple the onous is on the client to parse the following strings for their respective values.
'
'    Objects As String           ' ... Returns a vbCRLF delimited string describing the Objects found in the vbp.
'    References As String        ' ... Returns a vbCRLF delimited string describing the References found in the vbp.
'    Forms As String             ' ... Returns a vbCRLF delimited string describing the Forms found in the vbp.
'    Classes As String           ' ... Returns a vbCRLF delimited string describing the Classes found in the vbp.
'    Modules As String           ' ... Returns a vbCRLF delimited string describing the Modules found in the vbp.
'    UserControls As String      ' ... Returns a vbCRLF delimited string describing the UserControls found in the vbp.
    
    ' -------------------------------------------------------------------
    ProjectPath As String       ' ... Sets / Returns the path of the VB6 Project file being read.
    ProjectFileName As String   ' ... Sets / Returns the name of the vbp being read.
    ' -------------------------------------------------------------------
    
    ResourceFile As String
    Version As String
    
End Type ' ... QuickVBPInfo.

Public Type QuickMemberInfo
    Accessor As Long
    Attribute As String
    Declaration As String
    EditorLineStart As Long
    Index As Long
    Name As String
    LineCount As Long
    LineStart As Long
    LineEnd As Long
    Type As Long
    ValueType As String
End Type

' -------------------------------------------------------------------
' ... v6, structure for vbp references and objects
' -------------------------------------------------------------------
' ... we can expect the following
' ...   a GUID              ' ... last char is }
' ...   a version number    ' ... if Ref last char is #, if Obj last char is ;
' ...   a file name         ' ... might include absolute or relative path or just file name
' ...   a description       ' ... not necessarily provided
' ... in that order.

Public Type ReferenceInfo

'' ------------------------------------------------------------------- ' ... give it a miss.
'    Exists As Boolean           ' ... result of trying to find the file referenced,
'                                ' ... not in Registry but as a file.
' -------------------------------------------------------------------
    Guid As String
    VersionNumber As String     ' ... formatted as Major/Minor#Revision, note # delimiter.
    FileName As String          ' ... includes extension if exists.
    FileNameAndPath As String
    FileDescription As String
' -------------------------------------------------------------------
    FilePath As String          ' ... no terminating folder delimiter (\)
                                ' ... derived so may be empty if file not found.
    FileExtension As String     ' ... derived extension of file name.
' -------------------------------------------------------------------
    ' ... a bit over the top?
    Reference As Boolean        ' ... when true indicates item is a Reference.
    Component As Boolean        ' ... when true indicates item is a component.
' -------------------------------------------------------------------

' Note: file path related members could return absolute or relative path and even nothing.
'       There is no attempt to derive the full path from a relative one, this can
'       be done later.
'       Just because we derived a path doesn't mean it exists at this point,
'       we is just parsing info from line text.

End Type    ' ... ReferenceInfo


' -------------------------------------------------------------------

Sub ParseDec(pDec As String, ByRef pDecInfo As DeclarationInfo)

Dim sb() As Byte
Dim i As Long
Dim iChar As Long
Dim iLen As Long
Dim sTmpDec As String
Dim sTmp As String
Dim sW() As Byte
'Dim tDecInfo As DeclarationInfo
Dim j As Long
Dim iBegin As Long
Dim iEnd As Long
Dim iSum As Long
Dim bParamsDone As Boolean

Dim bInQuotes As Boolean
Dim iUB As Long

'    sTmpDec = "Public Function Dec(Optional pDec As String = vbNullString) As String"
'    sTmpDec = "Public Property Set Fred(Optional pDec As String = vbNullString) As String"
''    sTmpDec = "Public Property Get Fred() As String"
''    sTmpDec = "Public Property Let Fred(Value As string)"
'    sTmpDec = "Public Sub Init(ByRef pCodeInfo As CodeInfo, ByRef pTreeview As TreeView, Optional ByRef pImageList As ImageList, Optional ByVal pIncHeadCount As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)"
'    sTmpDec = "Public Init(ByRef pCodeInfo As CodeInfo, ByRef pTreeview As TreeView, Optional ByRef pImageList As ImageList, Optional ByVal pIncHeadCount As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean"
    sTmpDec = pDec
    
    sb = sTmpDec
    
    pDecInfo.Accessor = -1
    pDecInfo.AccessorAsString = "Public"
    
    pDecInfo.MemberType = 0
    pDecInfo.MemberTypeAsString = "Sub"
    
    pDecInfo.PropertyType = 0
    pDecInfo.PropertyTypeAsString = "Get"
    
    iUB = UBound(sb)
    
    For i = 0 To iUB Step 2
        iChar = sb(i)
        Select Case iChar
            Case 32, 40
                If iChar = 32 Then
                    sTmp = StrConv(sW, vbUnicode)
                    iLen = 0
                    ReDim sW(iLen)
                    Select Case sTmp
                        Case "Public", "Private", "Friend" ' 7 chars, ilen = 7?
                            ' ... accessor
                            pDecInfo.AccessorAsString = sTmp
                            If sTmp = "Private" Then
                                pDecInfo.Accessor = 1 ' 0
                            ElseIf sTmp = "Public" Then
                                pDecInfo.Accessor = 2
                            Else
                                pDecInfo.Accessor = 3
                            End If
                        Case "Function", "Property" ' 8 chars, ilen = 8?
                            ' ... type
                            pDecInfo.ReturnsSomething = True
                            pDecInfo.MemberTypeAsString = sTmp
                            If sTmp = "Function" Then
                                pDecInfo.MemberType = 1
                                pDecInfo.PropertyTypeAsString = vbNullString
                            Else
                                pDecInfo.MemberType = 2
                            End If
                        Case "Const"
                            pDecInfo.ReturnsSomething = True
                            pDecInfo.MemberTypeAsString = sTmp
                            pDecInfo.MemberType = 5
                        Case "Let", "Set" ' 3 chars, ilen = 3
                            ' ... property type
                            pDecInfo.ReturnsSomething = False
                            pDecInfo.PropertyTypeAsString = sTmp
                            If sTmp = "Let" Then
                                pDecInfo.PropertyType = 1
                            Else
                                pDecInfo.PropertyType = 2
                            End If
                        Case "As" ' 2 chars, ilen = 2
                            ' ... return or property data type
                            ' ... ripping the parameters string, earlier, should have
                            ' ... taken us to the end of the declaration so this As
                            ' ... represents the value or property type of the method / member
                            iLen = 0
                            ReDim sW(iLen)
                            For j = i + 2 To iUB Step 2
                                iChar = sb(j)
                                Select Case iChar
                                    Case 39, 32
                                        j = iUB
                                        Exit For
                                    Case Else
                                        ReDim Preserve sW(iLen)
                                        sW(iLen) = iChar
                                        If j + 1 < iUB Then
                                            iLen = iLen + 1
                                        End If
                                End Select
                            Next j
                            i = j + 2
                            If pDecInfo.ReturnsSomething Then
                                pDecInfo.ValueTypeAsString = StrConv(sW, vbUnicode)
                            End If
                    Case Else
                        pDecInfo.Name = sTmp
                    End Select
                Else
                    ' -------------------------------------------------------------------
                    ' ... parameters
                    iSum = 1
                    iBegin = (i / 2 + 1)
                    pDecInfo.Name = StrConv(sW, vbUnicode)
                    For j = i + 2 To iUB Step 2
                        iChar = sb(j)
                        Select Case iChar
                            Case 34:
                                bInQuotes = Not bInQuotes
                            Case 40, 41
                                If bInQuotes = False Then
                                    If iChar = 40 Then
                                        iSum = iSum + 1
                                    Else
                                        iSum = iSum - 1
                                        If iSum = 0 Then
                                            iEnd = (j / 2 + 1)
                                            If iEnd > iBegin Then
                                                ' ... extract params string
                                                pDecInfo.ParamString = Mid$(sTmpDec, iBegin + 1, iEnd - iBegin - 1)
                                            End If
                                            iLen = 0
                                            ReDim sW(iLen)
                                            i = j + 2
                                            bParamsDone = True
                                            Exit For
                                        End If
                                    End If
                                End If
                        End Select
                    Next j
                    bInQuotes = False
                End If
            Case Else
                ReDim Preserve sW(iLen)
                sW(iLen) = iChar
                iLen = iLen + 1
        End Select
    Next i
    
'    With pDecInfo
'        Debug.Print .Name, .ValueTypeAsString
'        Debug.Print .Accessor, .AccessorAsString
'        Debug.Print .MemberType, .MemberTypeAsString
'        Debug.Print .PropertyType, .PropertyTypeAsString
'        Debug.Print .ParamString
'    End With
    
End Sub



' -------------------------------------------------------------------

Sub pDec(Optional pDec As String = vbNullString) ',ByRef pDecInfo As DeclarationInfo)

Dim sb() As Byte
Dim i As Long
Dim iChar As Long
Dim iLen As Long
Dim sTmpDec As String
Dim sTmp As String
Dim sW() As Byte
Dim tDecInfo As DeclarationInfo
Dim j As Long
Dim iBegin As Long
Dim iEnd As Long
Dim iSum As Long
Dim bParamsDone As Boolean

Dim bInQuotes As Boolean
Dim iUB As Long

    sTmpDec = "Public Function Dec(Optional pDec As String = vbNullString) As String"
    sTmpDec = "Public Property Set Fred(Optional pDec As String = vbNullString) As String"
'    sTmpDec = "Public Property Get Fred() As String"
'    sTmpDec = "Public Property Let Fred(Value As string)"
    sTmpDec = "Public Sub Init(ByRef pCodeInfo As CodeInfo, ByRef pTreeview As TreeView, Optional ByRef pImageList As ImageList, Optional ByVal pIncHeadCount As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)"
    sTmpDec = "Public Init(ByRef pCodeInfo As CodeInfo, ByRef pTreeview As TreeView, Optional ByRef pImageList As ImageList, Optional ByVal pIncHeadCount As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean"
    
    sb = sTmpDec
    
    tDecInfo.Accessor = 1
    tDecInfo.AccessorAsString = "Public"
    
    tDecInfo.MemberType = 0
    tDecInfo.MemberTypeAsString = "Sub"
    
    tDecInfo.PropertyType = 0
    tDecInfo.PropertyTypeAsString = "Get"
    
    iUB = UBound(sb)
    
    For i = 0 To iUB Step 2
        iChar = sb(i)
        Select Case iChar
            Case 32, 40
                If iChar = 32 Then
                    sTmp = StrConv(sW, vbUnicode)
                    iLen = 0
                    ReDim sW(iLen)
                    Select Case sTmp
                        Case "Private", "Friend" ' 7 chars, ilen = 7?
                            ' ... accessor
                            tDecInfo.AccessorAsString = sTmp
                            If sTmp = "Private" Then
                                tDecInfo.Accessor = 0
                            Else
                                tDecInfo.Accessor = 2
                            End If
                        Case "Function", "Property" ' 8 chars, ilen = 8?
                            ' ... type
                            tDecInfo.ReturnsSomething = True
                            tDecInfo.MemberTypeAsString = sTmp
                            If sTmp = "Function" Then
                                tDecInfo.MemberType = 1
                                tDecInfo.PropertyTypeAsString = vbNullString
                            Else
                                tDecInfo.MemberType = 2
                            End If
                        Case "Let", "Set" ' 3 chars, ilen = 3
                            ' ... property type
                            tDecInfo.ReturnsSomething = False
                            tDecInfo.PropertyTypeAsString = sTmp
                            If sTmp = "Let" Then
                                tDecInfo.PropertyType = 1
                            Else
                                tDecInfo.PropertyType = 2
                            End If
                        Case "As" ' 2 chars, ilen = 2
                            ' ... return or property data type
                            ' ... ripping the parameters string, earlier, should have
                            ' ... taken us to the end of the declaration so this As
                            ' ... represents the value or property type of the method / member
                            iLen = 0
                            ReDim sW(iLen)
                            For j = i + 2 To iUB Step 2
                                iChar = sb(j)
                                Select Case iChar
                                    Case 39, 32
                                        j = iUB
                                        Exit For
                                    Case Else
                                        ReDim Preserve sW(iLen)
                                        sW(iLen) = iChar
                                        If j + 1 < iUB Then
                                            iLen = iLen + 1
                                        End If
                                End Select
                            Next j
                            i = j + 2
                            If tDecInfo.ReturnsSomething Then
                                tDecInfo.ValueTypeAsString = StrConv(sW, vbUnicode)
                            End If
                    End Select
                Else
                    ' -------------------------------------------------------------------
                    ' ... parameters
                    iSum = 1
                    iBegin = (i / 2 + 1)
                    tDecInfo.Name = StrConv(sW, vbUnicode)
                    For j = i + 2 To iUB Step 2
                        iChar = sb(j)
                        Select Case iChar
                            Case 34:
                                bInQuotes = Not bInQuotes
                            Case 40, 41
                                If bInQuotes = False Then
                                    If iChar = 40 Then
                                        iSum = iSum + 1
                                    Else
                                        iSum = iSum - 1
                                        If iSum = 0 Then
                                            iEnd = (j / 2 + 1)
                                            If iEnd > iBegin Then
                                                ' ... extract params string
                                                tDecInfo.ParamString = Mid$(sTmpDec, iBegin + 1, iEnd - iBegin - 1)
                                            End If
                                            iLen = 0
                                            ReDim sW(iLen)
                                            i = j + 2
                                            bParamsDone = True
                                            Exit For
                                        End If
                                    End If
                                End If
                        End Select
                    Next j
                    bInQuotes = False
                End If
            Case Else
                ReDim Preserve sW(iLen)
                sW(iLen) = iChar
                iLen = iLen + 1
        End Select
    Next i
    
    With tDecInfo
        Debug.Print .Name, .ValueTypeAsString
        Debug.Print .Accessor, .AccessorAsString
        Debug.Print .MemberType, .MemberTypeAsString
        Debug.Print .PropertyType, .PropertyTypeAsString
        Debug.Print .ParamString
    End With
    
End Sub

Sub ParseDeclarationInfo(ByRef pTheDeclaration As String, _
                         ByRef pDecInfo As DeclarationInfo, _
                Optional ByRef pOK As Boolean = False, _
                Optional ByRef pErrMsg As String = vbNullString)
Attribute ParseDeclarationInfo.VB_Description = "Tries to parse a string as a Declaration into a DeclarationInfo Structure."
' ... Tries to parse a string as a Declaration into a DeclarationInfo Structure.
Dim sBytes() As Byte
Dim i As Long
Dim iChar As Long
Dim iType As Long
Dim tDecInfo As DeclarationInfo
Dim iStart As Long
Dim iLen As Long

    On Error GoTo ErrHan:
    
    If Len(pTheDeclaration) = 0 Then Err.Raise vbObjectError + 1000, , "Not a valid Declaration for parsing"
    
    sBytes = pTheDeclaration
    
    For i = 0 To UBound(sBytes) Step 2
    
        iChar = sBytes(i)
        
        Select Case iChar
            
            Case 40 ' ... open parenthesis, end of accessor type and name, start of params
                
            Case 32 ' ... space, a word break
                
                
            
            Case Else
            
        End Select
    
    Next i
    
ResumeError:

Exit Sub

ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "modVB.ParseDeclarationInfo.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... ParseDeclarationInfo:

Sub ParseReferenceInfo(ByVal pTheObjRefString As String, _
                       ByRef pRefInfo As ReferenceInfo, _
              Optional ByRef pOK As Boolean = False, _
              Optional ByRef pErrMsg As String = vbNullString, _
              Optional ByVal pWithDebug As Boolean = False)
Attribute ParseReferenceInfo.VB_Description = "Parses a VBP Reference/Object Line into a ReferenceInfo Type/Structure."

Dim xRefInfo As ReferenceInfo

Dim sTmp As String

Dim sVersion() As Byte

Dim lngFound As Long
Dim lngLoop As Long
Dim lngCount As Long
Dim lngChar As Long

Dim bIsRef As Boolean

  
    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... bit of validation.
    If Trim$(Len(pTheObjRefString)) = 0 Then
        Err.Raise vbObjectError + 1000, , "Reference String not Initialised"
    End If
    
    sTmp = pTheObjRefString

    ' -------------------------------------------------------------------
    ' ... is reference or object (bIsRef = True / False respectively)
    bIsRef = Left$(LCase$(sTmp), 9) = "reference"
    
    If bIsRef = False Then
        If Left$(LCase$(sTmp), 6) <> "object" Then
            Err.Raise vbObjectError + 1000, , "Reference String invalid, neither Reference nor Object"
        End If
        xRefInfo.Component = True
    Else
        xRefInfo.Reference = True
    End If
    
    ' -------------------------------------------------------------------
    ' ... first split on equals sign.
    lngFound = InStr(1, sTmp, "=")
    If lngFound = 0 Then
        Err.Raise vbObjectError + 1000, , "Reference String invalid, no Equals sign ' = '"
    End If
    sTmp = Mid$(sTmp, lngFound + 1)
    
    ' -------------------------------------------------------------------
    ' ... second split on GUID.
    lngFound = InStr(1, sTmp, "}")
    If lngFound = 0 Then
        Err.Raise vbObjectError + 1000, , "Reference String invalid, no GUID end ' } '"
    End If
    ' -------------------------------------------------------------------
    xRefInfo.Guid = Left$(sTmp, lngFound)
    ' -------------------------------------------------------------------
    
    sTmp = Mid$(sTmp, lngFound + 1)
    ' -------------------------------------------------------------------
    ' ... third split, on Version.
    If Asc(Left$(sTmp, 1)) <> 35 Then
        Err.Raise vbObjectError + 1000, , "Reference String invalid, no start to Version ' # '"
    End If
    sTmp = Mid$(sTmp, 2)
    sVersion = sTmp
    
    ' ... looking for two # marks and then either another # (ref) or ; (obj)
    
    ' ... first mark.
    For lngLoop = 0 To UBound(sVersion) Step 2
        lngChar = sVersion(lngLoop)
        If lngChar = 35 Then
            Exit For
        End If
    Next lngLoop
    
    ' ... second mark.
    For lngLoop = lngLoop + 2 To UBound(sVersion) Step 2
        lngChar = sVersion(lngLoop)
        If bIsRef Then
            If lngChar = 35 Then Exit For ' ... e.g. #
        Else
            If lngChar = 59 Then Exit For ' ... e.g. ;
        End If
    Next lngLoop
    
    lngCount = lngLoop / 2 + 1
    xRefInfo.VersionNumber = Left$(sTmp, lngCount - 1)
    
    ' -------------------------------------------------------------------
    sTmp = Mid$(sTmp, lngCount + 1)
    ' -------------------------------------------------------------------
    
    ' ... fourth split on file name, then description.
    lngFound = InStr(1, sTmp, "#")
    If lngFound > 0 Then
        xRefInfo.FileName = Trim$(Left$(sTmp, lngFound - 1))
        xRefInfo.FileDescription = Mid$(sTmp, lngFound + 1)
    Else
        xRefInfo.FileName = Trim$(sTmp)
    End If
    
    ' -------------------------------------------------------------------
    ' ... file info.
    sTmp = xRefInfo.FileName
    
    lngFound = modStrings.InStrRevChar(sTmp, "\")
    
    If lngFound > 0 Then
        ' ... presume not just a file name, includes rel or abs path
        xRefInfo.FileNameAndPath = sTmp
        xRefInfo.FileName = Mid$(sTmp, lngFound + 1)
        xRefInfo.FilePath = Left$(sTmp, lngFound - 1)
    End If
    
    ' ... going for extension.
    lngFound = modStrings.InStrRevChar(sTmp, ".")
    If lngFound > 0 Then
        xRefInfo.FileExtension = Mid$(sTmp, lngFound + 1)
    End If
    
    ' -------------------------------------------------------------------
    
    With xRefInfo
        
        If pWithDebug Then
            Debug.Print IIf(xRefInfo.Reference, "Reference", "Component")
            Debug.Print .Guid
            Debug.Print .VersionNumber
            Debug.Print .FileName
            Debug.Print .FileDescription
            Debug.Print .FilePath
            Debug.Print .FileNameAndPath
            Debug.Print .FileExtension
        End If
        
        pRefInfo.Component = .Component
        pRefInfo.FileDescription = .FileDescription
        pRefInfo.FileExtension = .FileExtension
        pRefInfo.FileName = .FileName
        pRefInfo.FileNameAndPath = .FileNameAndPath
        pRefInfo.FilePath = .FilePath
        pRefInfo.Guid = .Guid
        pRefInfo.Reference = .Reference
        pRefInfo.VersionNumber = .VersionNumber
        
    End With
'    ' ... Object/Reference added as Name : Type : Exists : GUID (help sort by name) & If Ref, space & dervied path to file.
'    sTmpAdd = sRight & cBang & IIf(bIsRef = True, "0", "1") & "|1|" & sLeft & IIf(bIsRef, " " & sRefPath, "")
'    If bIsRef Then
'        moRefs.AddItemString sTmpAdd
'    Else
'        moObjects.AddItemString sTmpAdd
'    End If

    pOK = True
    
ResumeError:
    
    On Error GoTo 0
    sTmp = vbNullString
    Erase sVersion
    lngFound = 0&
    lngLoop = 0&
    lngCount = 0&
    lngChar = 0&

Exit Sub

ErrHan:
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "modVB.ParseReferenceInfo.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:


End Sub

Sub ReadVBFilePath(ByVal pTheProjectPath As String, ByRef pTheFilePath As String, Optional pFileFound As Boolean = False)
Attribute ReadVBFilePath.VB_Description = "Tries to return a full path and filename for a file referenced in a VB Project file."

Dim lngCount As Long
Dim bDoing As Boolean
Dim lngStart As Long
Dim lngLength As Long
Dim sTmpPath As String
Const cFind As String = "..\"
Const cCharSlash As String = "\"
Const cDriveSig As String = ":\"

' -------------------------------------------------------------------
' ... helper to generate file path from project path and named file's relative path.
' ... pFileFound returns true if the file is accessible.
' -------------------------------------------------------------------
' Note:
'       This was written to target file references found in Visual Basic Project files (.VBP/.VBG)
'       and is probably of little use otherwise.
' -------------------------------------------------------------------
' v7, note:
'       looks like i've fairly buggered things up in here
'       hacked another bug when the file path has more subfolders than
'       the project folder, the hack may cause other bugs, good luck.
' -------------------------------------------------------------------
'       pretty weird logic going on in here, not quite sure where I was at when
'       I knocked it up, a different planet perhaps?
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If Mid$(pTheFilePath, 2, 2) = cDriveSig Then
        ' ... absolute path on different drive, no change.
    Else
        sTmpPath = pTheProjectPath
        If Left$(pTheFilePath, 3) = cFind Then      ' ... indicates a relative path not within project's directory or sub-directory thereof.
            ' ... count the number of instances of ..\ in pTheFilePath
            ' ... use the result later to derive a path based upon the project folder.
            lngStart = 1
            lngLength = Len(cFind)
            Do
                bDoing = False
                If Mid$(pTheFilePath, lngStart, lngLength) = cFind Then
                    bDoing = True
                    lngCount = lngCount + 1
                    lngStart = lngStart + lngLength
                End If
            Loop While bDoing = True
            
            ' ... it could be that lngCount remains above 0 and thereby cause a bit of a never ending loop
            ' ... see below for escape hack.
            Do While lngCount > 0
                lngStart = modStrings.InStrRevChar(sTmpPath, cCharSlash) ' (note dependency on modStrings).
                If lngStart Then
                    sTmpPath = Left$(sTmpPath, lngStart - 1)
                    pTheFilePath = Mid$(pTheFilePath, 4)
                    lngCount = lngCount - 1
                Else
                    ' -------------------------------------------------------------------
                    ' ... v7
                    ' ... we've found the start of the project directory but
                    ' ... the file path is still prefixed with ' ..\ ' * lngCount
                    ' ... e.g. if lngCount = 2 the file path reads ..\..\[file]
                    ' ...      if lngCount = 3 the file path reads ..\..\..\[file] etc.
                    ' ... so remove the leading ..\ from pTheFilePath.
                    ' -------------------------------------------------------------------
                    If lngCount > 0 Then
                        Do While Left$(pTheFilePath, 3) = "..\"
                            pTheFilePath = Mid$(pTheFilePath, 4)
                            lngCount = lngCount - 1
                        Loop
                        ' -------------------------------------------------------------------
                        ' ... make sure we can escpape the loop
                        ' ... if we buggered up then the file won't be found.
                        If lngCount <> 0 Then lngCount = 0
                        ' -------------------------------------------------------------------
                    End If
                    ' -------------------------------------------------------------------
                End If
            Loop
            
        End If
        
        If Right$(sTmpPath, 1) <> cCharSlash Then sTmpPath = sTmpPath & cCharSlash
        pTheFilePath = sTmpPath & pTheFilePath
        
    End If
    
    pFileFound = Dir$(pTheFilePath, vbNormal) <> ""          ' ... this test could call the Drive not ready prompt.
                                                                ' ... thanks to RDe & B.McK.
'    Debug.Assert pFileFound                                     ' ... stop in ide run mode if file not found.
    
' -------------------------------------------------------------------
' ... Note: This method takes into account the following possible path values;
'           1. Different drive, absolute e.g. Z:\Code\VBProject\Form1.frm
'           2. Same drive, same folder as vbp e.g. Form1.frm
'           3. Same drive, off vbp folder e.g. \Forms\Form1.frm
'           4. Same drive, entirely different folder to vbp e.g. ..\..\DraftCode\Form1.frm
' -------------------------------------------------------------------
Exit Sub
ErrHan: ' ... Mid$ will error if invalid values are provided.
    Debug.Print "modVB.ReadVBFilePath.Error: " & Err.Description
End Sub ' ... ReadVBFilePath:


Public Sub ReadVBName(ByVal TheFile As String, _
                      ByRef TheName As String, _
             Optional ByVal pIsForm As Boolean = False, _
             Optional ByRef pType As Integer = 0, _
             Optional ByRef pOK As Boolean = False, _
             Optional ByRef pErrMsg As String = vbNullString)
Attribute ReadVBName.VB_Description = "Returns the Name Attribute of a VB Source file and can test for Form Type when pIsFOrm is True."

Dim sText As String
Dim bOK As Boolean
Dim sErrMsg As String
Dim lngStart As Long
Dim lngEnd As Long
Dim lngSkip As Long
Dim sFind As String

Const cFind As String = "Attribute VB_Name ="
Const cMDISig As String = "Begin VB.MDIForm"
Const cFormSig As String = "Begin VB.Form"
Const cMDIChildSig As String = "MDIChild        =   -1  'True"
' -------------------------------------------------------------------
' ... helper to retrieve the name of a form/usercontrol from its file
' ... because it is not written in the vbp entry.  it uses the Attibute VB_Name = "[name]"
' ... as the source to extracting the desired value.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    
    sText = modReader.ReadFile(TheFile, bOK, sErrMsg)
    
    If bOK = True Then
    
        sFind = cFind
        
        lngStart = InStr(1, sText, vbCrLf & sFind)
        
        If lngStart > 0 Then
            
            ' ... expecting quotes wrapping value.
            lngEnd = InStr(lngStart + 1, sText, vbCrLf)
            lngSkip = Len(sFind)
            lngStart = lngSkip + lngStart + 2 + 1
            
            TheName = Mid$(sText, lngStart, lngEnd - lngStart - 1)
            
            modStrings.RemoveQuotes TheName
        
        Else
            
            TheName = vbNullString
            pErrMsg = "The name of the object could not be retrieved."
        
        End If
        

        If pIsForm Then
            
            If InStr(1, sText, cMDISig) > 0 Then
                ' ... an mdi form.
                pType = 1
            
            Else
                
                If InStr(1, sText, cFormSig) > 0 Then
                    
                    ' ... normal form, but is it an MDIChild?
                    pType = 2
                    
                    If InStr(1, sText, cMDIChildSig) > 0 Then
                        pType = 3
                    End If
                
                End If
            
            End If
        
        End If
    
    End If
    
    bOK = Len(TheName) > 0

ResumeError:
    
    sText = vbNullString
    pOK = bOK
' -------------------------------------------------------------------
' Note: a shortcut for forms and user controls is to access the "Begin VB.[object type]" where
'       object type could be MDIForm, Form or UserControl.  This information should be found
'       on the second line in the vbp data and is not wrapped in quotes and this suggests
'       that processing should be faster if this method to retrieve the name is used instead.
' -------------------------------------------------------------------
Exit Sub
ErrHan:
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "modVB.ReadVBName.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub ' ... ReadVBName:

Public Function ExtractParamString(ByVal pDeclaration As String, _
                          Optional ByRef pOK As Boolean = False, _
                          Optional ByRef pErrMsg As String = vbNullString) As String
Attribute ExtractParamString.VB_Description = "Returns the Parameter portion of a Method / Event Declaration without the opening and closing parenthesis."

' Function:         ExtractParamString
' Returns:          String

' ... Extracts the Parameters to a declaration, the bit between the open and close brackets.

' ... v8, new method
' ... step through the chars of the declaration and count and compare
' ... the number of open and closed brackets found.

' ... This method does require a compilable declaration in order to work correctly

Dim bBytes() As Byte

Dim i As Long
Dim iChar As Long
Dim iBegin As Long
Dim iEnd As Long
Dim iSum As Long

Dim bInQuotes As Boolean

' -------------------------------------------------------------------

' ... Note: Doesn't return the opening and closing parenthesis.

' ... ? ExtractParamString("Public Sub Init(ByRef pCodeInfo As CodeInfo, ByRef pTreeview As TreeView, Optional ByRef pImageList As ImageList, Optional ByVal pIncHeadCount As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)")
' ... = ByRef pCodeInfo As CodeInfo, ByRef pTreeview As TreeView, Optional ByRef pImageList As ImageList, Optional ByVal pIncHeadCount As Boolean = True, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString

' ... ? ExtractParamString("Public Sub QuickSortOnStringArray(ByRef MyArray() As String, ByVal lngLBound As Long, ByVal lngUbound As Long)")
' ... = ByRef MyArray() As String, ByVal lngLBound As Long, ByVal lngUbound As Long

' ... ? ExtractParamString("Public Sub QuickSortOnStringArray(ByRef MyArray() As String, ByVal lngLBound() As Long, ByVal lngUbound() As Long)")
' ... = ByRef MyArray() As String, ByVal lngLBound() As Long, ByVal lngUbound() As Long

' ... ? ExtractParamString("Public Function StringThing(Optional ByRef MyString As String = " & Chr(34) & "Some text () value" & Chr(34) & ", Optional ByVal lngLBound As Long = 0) As String()")
' ... = Optional ByRef MyString As String = "Some text () value", Optional ByVal lngLBound As Long = 0

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If Len(pDeclaration) Then
        
        bBytes = pDeclaration
        
        For i = 0 To UBound(bBytes) Step 2
            
            iChar = bBytes(i)
            
            Select Case iChar
                
                Case 34: bInQuotes = Not bInQuotes
                    ' ... toggle in quotes.
                    
                Case 40, 41
                    ' ... ( or ) respectively
                    
                    If bInQuotes = False Then
                    
                        If iChar = 40 Then
                            
                            ' ... increment sum
                            iSum = iSum + 1
                            
                            If iBegin = 0 Then
                                ' ... capture start
                                iBegin = (i / 2 + 1)
                            End If
                        
                        ElseIf iChar = 41 Then
                            
                            ' ... decrement sum
                            iSum = iSum - 1
                            
                            If iSum = 0 Then
                                
                                ' ... capture end and quit loop
                                iEnd = (i / 2 + 1)
                                Exit For
                            
                            End If
                        
                        End If
                    
                    End If
            
            End Select
        
        Next i
                
        If iEnd > iBegin + 1 Then
        
            ' ... attempt to extract the params bit only.
            ExtractParamString = Mid$(pDeclaration, iBegin + 1, iEnd - iBegin - 1)
            
        End If
        
    End If
    
    pOK = True
    pErrMsg = vbNullString
    
ResumeError:
    
    Erase bBytes
    iBegin = 0
    iEnd = 0
    iSum = 0
    i = 0
    
Exit Function

ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "modVB.ExtractParamString.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function


 Sub pGetDecParams(pParamString As String, _
                          pParamsArray() As String, _
                 Optional pCount As Long = 0)

' ... Returns a string array (via pParamsArray) of the parameters found in the input parameters string.
' -------------------------------------------------------------------
' Note:     This method expects a single line string e.g. no carriage returns.
'           This can be achieved by reading the QuickMember Declaration property.
' -------------------------------------------------------------------

Dim sParamString As String
Dim bParams() As Byte
Dim lngLoop As Long
Dim bInQuotes As Boolean
Dim lngChar As Long
Dim lngPos As Long
Dim lngStartPos As Long
Dim lngLength As Long
Dim sParameter As String

    
    ' -------------------------------------------------------------------
    pCount = 0
    
    If Len(pParamString) = 0 Then
        Exit Sub
    End If
    ' -------------------------------------------------------------------
    sParamString = pParamString
    ' -------------------------------------------------------------------
    bParams = sParamString
    ' -------------------------------------------------------------------
    lngStartPos = 1
    
    ' -------------------------------------------------------------------
    ' ... loop thru' the params string grabbing each comma delimited param.
    ' ... if in a quote and get comma then ignore.
    For lngLoop = LBound(bParams) To UBound(bParams) Step 2
                        
        lngPos = lngPos + 1             ' ... refers to the actual postion currently being read.
        lngLength = lngLength + 1       ' ... refers to the length of the string to be extracted from the source as a single parameter.
        
        lngChar = bParams(lngLoop)
        
        If lngChar = 34 Then            ' ... quote e.g. "
            bInQuotes = Not bInQuotes
        End If
    
        If lngChar = 44 Then            ' ... comma e.g. , (param delimiter)
            If bInQuotes = False Then
                ' -------------------------------------------------------------------
                sParameter = Trim$(Mid$(sParamString, lngStartPos, lngLength - 1))
                ' -------------------------------------------------------------------
                ReDim Preserve pParamsArray(pCount)
                pParamsArray(pCount) = sParameter
                lngStartPos = lngPos + 1
                pCount = pCount + 1
                lngLength = 0
            End If
        End If
                
    Next lngLoop
    
    ' -------------------------------------------------------------------
    ' ... read trailing string if any.
    If lngStartPos < lngPos Then
        sParameter = Trim$(Mid$(sParamString, lngStartPos))
        ReDim Preserve pParamsArray(pCount)
        pParamsArray(pCount) = sParameter
        pCount = pCount + 1
    End If
    
    Erase bParams
    sParamString = vbNullString
    sParameter = vbNullString
    
End Sub ' ... pGetDecParams:

Public Sub GetMemberSyntax(ByRef pMemberDeclaration As String, ByRef rReturn As String)
Attribute GetMemberSyntax.VB_Description = "Tries to compile a syntax statement (to rReturn) for a method / member from its declaration."
' ... Tries to compile a syntax statement (to rReturn) for a method / member from its declaration.
Dim sParams() As String
Dim sTmpParams As String
Dim iParamCount As Long
Dim tParamInfo As ParamInfo
Dim i As Long
Dim sTmpParam As String
Dim sTmpReturn As String

    On Error GoTo ErrHan:
    
    rReturn = vbNullString
    
    If Len(pMemberDeclaration) = 0 Then Exit Sub
    
    sTmpParams = ExtractParamString(pMemberDeclaration)
    
    pGetDecParams sTmpParams, sParams, iParamCount
    
    If iParamCount Then
        For i = 0 To iParamCount - 1
            ParseParamInfoItem sParams(i), tParamInfo
            sTmpParam = tParamInfo.Name
            If tParamInfo.IsByRef Then
                sTmpParam = "ByRef " & sTmpParam ' "< " & sTmpParam ' & "><"
            Else
                sTmpParam = "ByVal " & sTmpParam ' "> " & sTmpParam '& ">"
            End If
            sTmpParam = sTmpParam & ": " & tParamInfo.Type
            If tParamInfo.IsOptional Then
                sTmpParam = "[" & sTmpParam & " = " & tParamInfo.DefaultValue & " ]"
            End If
            If Len(sTmpReturn) Then
                sTmpParam = ", " & sTmpParam
            End If
            sTmpReturn = sTmpReturn & sTmpParam
        Next i
    End If
    
ResumeError:
    rReturn = sTmpReturn
Exit Sub

ErrHan:

    Debug.Print "modVB.GetMemberSyntax.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub


Private Sub pExtractParameters(pDeclaration As String, pTheParameters As String)
Dim bBytes() As Byte
Dim i As Long
Dim bIQ As Boolean
Dim iBegin As Long
Dim iEnd As Long
Dim iSum As Long

    On Error GoTo ErrHan:
    pTheParameters = vbNullString
    If Len(pDeclaration) = 0 Then Exit Sub
    bBytes = pDeclaration
    For i = 0 To UBound(bBytes) Step 2 ' notice the non existent mbcs handling
        Select Case bBytes(i)
            Case 34: bIQ = Not bIQ
            Case 40, 41
                If bIQ Then GoTo JumpNext:
                If bBytes(i) = 40 Then
                    iSum = iSum + 1
                    If iBegin = 0 Then
                        iBegin = (i / 2 + 1)
                    End If
                Else
                    iSum = iSum - 1
                    If iSum = 0 Then
                        iEnd = (i / 2 + 1)
                        Exit For
                    End If
                End If
        End Select
JumpNext:
    Next i
    If iEnd > iBegin + 1 Then
        pTheParameters = Mid$(pDeclaration, iBegin + 1, iEnd - iBegin - 1)
    End If
ResumeError:
    Erase bBytes
    iBegin = 0: iEnd = 0: iSum = 0: i = 0
Exit Sub
ErrHan:
    Debug.Print "CR2.pExtractParameters.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
End Sub

'
'Public Function LeftOfComment(ByVal pTheText As String, Optional ByVal pCommentChar As String = "'", Optional pTrimRight As Boolean = True) As String
'
'' ... get the left side of a line of text from the first comment char.
'' ... and optionally Right Trim the return value.
'' ... expects a single char comment character although any character will do, not just the apostrophe.
'' ... attempts to ignore comment chars within quotes.
'' ... reads bytes rather than characters.
'
'Dim lngPos As Long
'Dim lngLoop As Long
'Dim lngText As Long
'Dim lngComment As Long
'Dim bInQuotes As Boolean
'
'Dim lngCommentChar As Long
'
'Dim lngCurrentChar As Long
'Dim bText() As Byte
'
'Const cQuoteChar As Long = 34
'
'    LeftOfComment = pTheText
'
'    ' ... validation.
'
'    lngText = Len(pTheText)
'    lngComment = Len(pCommentChar)
'
'    If lngComment <> 1 Then Exit Function
'    If lngText = 0 Then Exit Function
'
'    lngPos = InStr(1, pTheText, pCommentChar)
'
'    If lngPos = 0 Then Exit Function
'
'    ' ... end validation.
'
'    lngPos = 0
'    lngCommentChar = Asc(pCommentChar)
'
'    bText = pTheText
'    ' ... loop the bytes looking for the comment char.
'    For lngLoop = 0 To UBound(bText) Step 2
'
'        lngPos = lngPos + 1
'        lngCurrentChar = bText(lngLoop)
'
'        Select Case lngCurrentChar
'            ' ... check if quote char and update bInQuotes accordingly.
'            Case cQuoteChar: bInQuotes = Not bInQuotes
'            Case lngCommentChar:
'                ' ... if not in quotes and is match then found pos (-1).
'                If Not bInQuotes Then Exit For
'
'        End Select
'
'    Next lngLoop
'
'    ' ... extract left side.
'    LeftOfComment = Left$(pTheText, lngPos - 1)
'    If pTrimRight Then
'        ' ... trim right side if required.
'        LeftOfComment = RTrim$(LeftOfComment)
'    End If
'
'    ' ... clean up.
'    Erase bText
'    lngPos = 0&
'    lngText = 0&
'    lngComment = 0&
'    lngCommentChar = 0&
'
'End Function
'
