Attribute VB_Name = "modVBCB"
Option Explicit

' -------------------------------------------------------------------
Public Type MenuInfo

' ... items with needs to be tristate;
' ...   default, inused, properties are not
' ...   written to the designer
' ...   checked and index are really is checkable and is in menu array
' ...   when they are not listed they must be ignored
' ...   enabled can default to true if not listed


' ... a structure to describe various attributes of a menu item.

    Caption As String           ' ... caption property
    Checked As Boolean          ' ... checked property, needs to be tristate
    Enabled As Boolean          ' ... enabled property, again, needs to be tristate
    HelpContextID As Long       ' ... help context id property
    ID As Long                  ' ... unique index within entire menu
    Index As Long               ' ... index within menu array, needs to be tristate
    IsMainItem As Boolean       ' ... is this a main parent e.g. File, Admin, Help ...
    IsParent As Boolean         ' ... indicates whether the item has children / is a parent
    IsSeparator As Boolean      ' ... true when caption is "-"
    NestLevel As Long           ' ... the indent level of the menu item
    MethodLineNumber As Long    ' ... the line number of the method in the source code
    MethodName As String          ' ... derived from menu name & "_Click"
    Name As String              ' ... the name property
    NegotiatePosition As Long   ' ... the negotiate position property
    ParentID As Long            ' ... the id of the menu's parent
    ShortCut As String          ' ... the short cut property
    WindowList As Boolean       ' ... the window list property
    
End Type
' -------------------------------------------------------------------

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
    Debug.Assert pFileFound                                     ' ... stop in ide run mode if file not found.
    
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




Public Function ParseVBMenuEx(Optional ByVal pHeader As String, _
                   Optional ByVal pDelimiter As String = vbCrLf) As String

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... parse a vb menu into a string of data rows
    ' ... as we read the menu string we can note that
    ' ... we will have three types of item coming at us
    ' ... a Begin VB.Menu [menu name]
    ' ... a menu Property, e.g. Caption
    ' ... an End to represent the end of the current menu item.
    ' ... an index is required to identify items in the rows of menu data.
    ' ... every time we come across a Begin VB.Menu the index / count will increment.
    ' ... every Begin VB.Menu is terminated by an End (or is meant to)
    ' ... every menu item will have at least one property
    ' ... but may have more than one; the limit will be the number
    ' ... of properties available to the menu editor.
    ' ... a calculation is required to derive the end of a menu
    ' ... current no. of begins - current number of ends
    ' ... where the result is zero we will find a group header menu
    ' ... where the result is > 1 a new sub menu has been defined.
    ' ... where the result is = 1 a new sub menu is being closed.
    
    ' ... i want to trap all the info for each item so that I can
    ' ... identify individual members and their parentage as well as
    ' ... extract sub menu info about a menu item.
    ' ... these will help me write the menu to a tree view and a
    ' ... dynamic pop-up menu on the fly.
    
    ' ... the definition of menus has tab indentation per level
    ' ... and this can help identify the level to which an item belongs
    ' ... but not the item.
    
    
Dim iSum As Long
Dim sLines() As String
Dim iLines As Long
Dim sTmp As String
Dim i As Long
Dim sName As String
Dim iIndex As Long
Dim iParent As Long
Dim iParents() As Long
Dim sCaption As String
Dim iPIndex As Long
Dim xMenuInfo() As MenuInfo
Dim iCurrIndex As Long

Dim sTest As String

sTest = sTest & "Version 5#"
sTest = sTest & vbNewLine & "Begin VB.UserControl ucMenus"
sTest = sTest & vbNewLine & "   BorderStyle = 1        'Fixed Single"
sTest = sTest & vbNewLine & "   ClientHeight = 1125"
sTest = sTest & vbNewLine & "   ClientLeft = 0"
sTest = sTest & vbNewLine & "   ClientTop = 0"
sTest = sTest & vbNewLine & "   ClientWidth = 825"
sTest = sTest & vbNewLine & "   InvisibleAtRuntime = -1   'True"
sTest = sTest & vbNewLine & "   ScaleHeight = 1125"
sTest = sTest & vbNewLine & "   ScaleWidth = 825"
sTest = sTest & vbNewLine & "   Begin VB.Menu mnuProject"
sTest = sTest & vbNewLine & "      Caption = " & Chr$(34) & "Project" & Chr$(34)
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPLoadProject"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Load Project" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPSep1a"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "-" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPOpen"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Open" & Chr$(34)
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNewWindow"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in New Window" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPOpenFolder"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "Containing Folder" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNotePad"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in Text Editor" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB5" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB6" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "   End"
sTest = sTest & vbNewLine & "   Begin VB.Menu mnuProject2"
sTest = sTest & vbNewLine & "      Caption = " & Chr$(34) & "Project2" & Chr$(34)
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPLoadProject2"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Load Project2" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPSep1a2"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "-2" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPOpen2"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Open2" & Chr$(34)
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNewWindow2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in New Window2" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPOpenFolder2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "Containing Folder2" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNotePad2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in Text Editor2" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE22"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB52" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB62" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "      End"

sTest = sTest & vbNewLine & "   End"
sTest = sTest & vbNewLine & "End"

    
    pHeader = sTest
    
    ReDim iParents(999)
    ReDim xMenuInfo(999)
    
    SplitString pHeader, sLines, pDelimiter, iLines
    
    If iLines Then
        
        For i = 0 To iLines - 1
        
            sTmp = LCase$(LTrim$(sLines(i)))
'            Debug.Print sTmp
            If Left$(sTmp, 14) = "begin vb.menu " Then
                
                ' ... beginning a new menu item
                ' ... index and name available
                ' ... need to dervive parent index
                ' ... increase level/sum by one
                ' ... the parent
                
                
                iIndex = iIndex + 1         ' ... running index of menu items
                sName = Mid$(sTmp, 15)      ' ... name of menu item
                iSum = iSum + 1             ' ... increment level
                With xMenuInfo(iIndex)
                    .Name = sName
                    .MethodName = sName & "_Click"
                    .ID = iIndex
                    .NestLevel = iSum
                    If iSum = 1 Then
                        .ParentID = 0
                    ElseIf iSum = 0 Then
                        Debug.Print "Hmmm"
                    Else
'                        .ParentID = xMenuInfo(iSum - 1).ID
                        .ParentID = xMenuInfo(iIndex - 1).ID
                    End If
                    
                End With
                iCurrIndex = iIndex
'                Debug.Print iIndex, iSum, Space$(iSum); sName
            
            ElseIf Left$(sTmp, 3) = "end" Then
                
                ' ... ending a menu item, not necessarily the last one
                ' ... reduce level/sum indicator by one
                ' ... when sum is zero, a main menu item is ended
                ' ... and this means that the parent defaults to 0
                
                ' ... when sum is one it means that a sub item
                ' ... to a main menu has been closed
                ' ... and this means that the parent of the next item
                ' ... is this same main menu item.
                
                
                
                iSum = iSum - 1
                If iSum = 0 Then '11
'                    Debug.Print "Here"
'                    Exit For
                ElseIf iSum = 1 Then
'                    Debug.Print iPIndex, iParents(iSum)
                End If
                
                iCurrIndex = iCurrIndex - 1 ' ... current parent index please
                
            ElseIf Left$(sTmp, 7) = "caption" Then
            
                ' ... inbetween a begin & end menu of the current menu
                ' ... the next line could be another begin, another property
                ' ... or and end
                
                sCaption = LTrim$(Mid$(sTmp, 8))
                If Left$(sCaption, 2) = "= " Then
                    sCaption = LTrim$(Mid$(sCaption, 3))
                End If
                
                With xMenuInfo(iCurrIndex) ' (iIndex)
                    .Caption = sCaption
                    
                    Debug.Print .ID, .NestLevel, .ParentID, .Name, .Caption
                    
                End With
                
'                Debug.Print "index: " & iIndex, "level: " & iSum, "parent: x", "name: " & sName, "caption: " & sCaption
                
            End If
                
        Next i
        
    
    End If

ResumeError:

Exit Function

ErrHan:

    Debug.Print "modVB.ParseVBMenuEx.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Private Function TestMenu() As String
Dim sTest As String

sTest = sTest & "Version 5#"
sTest = sTest & vbNewLine & "Begin VB.UserControl ucMenus"
sTest = sTest & vbNewLine & "   BorderStyle = 1        'Fixed Single"
sTest = sTest & vbNewLine & "   ClientHeight = 1125"
sTest = sTest & vbNewLine & "   ClientLeft = 0"
sTest = sTest & vbNewLine & "   ClientTop = 0"
sTest = sTest & vbNewLine & "   ClientWidth = 825"
sTest = sTest & vbNewLine & "   InvisibleAtRuntime = -1   'True"
sTest = sTest & vbNewLine & "   ScaleHeight = 1125"
sTest = sTest & vbNewLine & "   ScaleWidth = 825"
sTest = sTest & vbNewLine & "   Begin VB.Menu mnuProject"
sTest = sTest & vbNewLine & "      Caption = " & Chr$(34) & "Project" & Chr$(34)
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPLoadProject"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Load Project" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPSep1a"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "-" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPOpen"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Open" & Chr$(34)
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNewWindow"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in New Window" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPOpenFolder"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "Containing Folder" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNotePad"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in Text Editor" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB5" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB6" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "   End"
sTest = sTest & vbNewLine & "   Begin VB.Menu mnuProject2"
sTest = sTest & vbNewLine & "      Caption = " & Chr$(34) & "Project2" & Chr$(34)
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPLoadProject2"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Load Project2" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPSep1a2"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "-2" & Chr$(34)
sTest = sTest & vbNewLine & "      End"
sTest = sTest & vbNewLine & "      Begin VB.Menu mnuPOpen2"
sTest = sTest & vbNewLine & "         Caption = " & Chr$(34) & "Open2" & Chr$(34)
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNewWindow2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in New Window2" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPOpenFolder2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "Containing Folder2" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPNotePad2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "in Text Editor2" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE22"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB52" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "         Begin VB.Menu mnuPIDE2"
sTest = sTest & vbNewLine & "            Caption = " & Chr$(34) & "VB62" & Chr$(34)
sTest = sTest & vbNewLine & "         End"
sTest = sTest & vbNewLine & "      End"

sTest = sTest & vbNewLine & "   End"
sTest = sTest & vbNewLine & "End"

TestMenu = sTest
sTest = vbNullString

End Function

Public Sub TrimLeftSpaces(ByRef pString As String, Optional ByRef pSpaceCount As Long = 0)

Dim sBytes() As Byte
Dim i As Long
Dim c As Long
Dim j As Long

    pSpaceCount = 0
    
    If Len(pString) = 0 Then Exit Sub

    sBytes = pString
        
    For i = 0 To UBound(sBytes) Step 2
        c = sBytes(i)
        If c = 32 Then
            j = j + 1
        Else
            Exit For
        End If
    Next i
    
    If j > 0 Then
        pString = Mid$(pString, j + 1)
        pSpaceCount = j
    End If
    
'    Debug.Print pString, j

    Erase sBytes
    i = 0&
    j = 0&
    c = 0&
    
End Sub

Public Function ParseVBMenu(Optional ByVal pHeader As String, _
                   Optional ByVal pDelimiter As String = vbCrLf, _
                   Optional ByRef pMenuMethods As String = vbNullString) As String
Attribute ParseVBMenu.VB_Description = "Tries to return a vbCrLf delimited string of Menu related data from the Header Section of a source file."

' ... pHeader is the header section of a source file.
' ... pDelimiter is the delimiter to use to split the header into lines.
' ... pMenuMethods returns a string containing the names of the menu click events.

' ... pHeader is optional for testing purposes only!
' ... if it is empty, some test data is used.

' ... The Return Value is a vbCrLf Line Delimited string describing various attributes
' ...   of each menu item encountered.

Dim sLines() As String
Dim iParents() As Long

Dim sD As String            ' ... output, field delimiter.
Dim sCaption As String
Dim sName As String
Dim sTmp As String
Dim sReturn As String

Dim i As Long
Dim j As Long
Dim iIndex As Long
Dim iLevel As Long
Dim iLines As Long
Dim iParent As Long
Dim iLastLevel As Long
Dim iParentIndex As Long

' ... primitive menu parser
' ... outputting lines of menu descriptors
' ... e.g.
'    Print ParseVBMenu
'    Index         Level         ParentID     Name                 Caption
'     1             1             0            mnuProject           "Project"
'     2             2             1            mnuPLoadProject      "Load Project"
'     3             2             1            mnuPSep1a            "-"
'     4             2             1            mnuPOpen             "Open"
'     5             3             4            mnuPNewWindow        "in New Window"
'     6             3             4            mnuPOpenFolder       "Containing Folder"
'     7             3             4            mnuPNotePad          "in Text Editor"
'     8             3             4            mnuPIDE2             "VB5"
'     9             3             4            mnuPIDE              "VB6"
'     10            1             0            mnuProject2          "Project2"
'     11            2             10           mnuPLoadProject2     "Load Project2"
'     12            2             10           mnuPSep1a2           "-2"
'     13            2             10           mnuPOpen2            "Open2"
'     14            3             13           mnuPNewWindow2       "in New Window2"
'     15            3             13           mnuPOpenFolder2      "Containing Folder2"
'     16            3             13           mnuPNotePad2         "in Text Editor2"
'     17            3             13           mnuPIDE22            "VB52"
'     18            3             13           mnuPIDE2             "VB62"

' ... busked my way through this a bit too clueless especially as to how to get the parent to a child
' ... got a bug stepping backwards in level to previous parent.
' ... K, so the idea is that a string descriptor results which can then
' ... be split on vbCrLf and parsed into a menu info structure.

    On Error GoTo ErrHan:

    If Len(pHeader) = 0 Then
        pHeader = TestMenu          ' ... just grabbing some test data.
    End If
    
    pMenuMethods = vbNullString
    
    If Len(pHeader) = 0 Then Exit Function
    
    ReDim iParents(99)
    
    SplitString pHeader, sLines, pDelimiter, iLines
    
    If iLines Then
    
        sD = Chr$(1)
        
        For i = 0 To iLines - 1
            
            sTmp = sLines(i)
            
            TrimLeftSpaces sTmp
            
            If Left$(LCase$(sTmp), 14) = "begin vb.menu " Then
                
                iParentIndex = iIndex       ' ... new parent index if level changed
                
                iIndex = iIndex + 1         ' ... running index of menu items
                sName = Mid$(sTmp, 15)      ' ... name of menu item
                sName = Trim$(sName)
                iLevel = iLevel + 1         ' ... increment level
                
                If iLevel > iLastLevel Then
                    iParents(iLastLevel) = iParentIndex
                    iParent = iParents(iLastLevel)
                    iLastLevel = iLevel
                ElseIf iLevel = iLastLevel Then
                    iParent = iParent
                ElseIf iLevel < iLastLevel Then
                    iLastLevel = iLevel
                End If
            
            Else
                
                If iIndex > 0 Then
                
                    If Left$(LCase$(sTmp), 3) = "end" Then
                    
                        iLevel = iLevel - 1
                        If iLevel < 0 Then iLevel = 0
                        If iLevel = 0 Then
                            iParent = 0
                        End If
                    
                    ElseIf Left$(LCase$(sTmp), 7) = "caption" Then
                        
                        sCaption = LTrim$(Mid$(sTmp, 8))
                        If Left$(sCaption, 2) = "= " Then
                            sCaption = LTrim$(Mid$(sCaption, 3))
                            RemoveQuotes sCaption
                        End If
                                    
                        ' ... loop through next lines looking for
                        ' ... a new begin or an end capturing any other
                        ' ... properties mentioned using splitstringpair
                        
                        sTmp = vbNullString
                        If Len(sReturn) Then
                            sTmp = vbCrLf
                        End If
                        
                        sTmp = sTmp & CStr(iIndex) & sD & CStr(iLevel) & sD
                        sTmp = sTmp & CStr(iParent) & sD & sName & sD
                        sTmp = sTmp & sCaption
                        
                        sReturn = sReturn & sTmp
                        
                        pMenuMethods = pMenuMethods & " " & sName & "_Click "
                        
                    End If
                
                End If
            End If
        Next i
    End If

ResumeError:
        
    ParseVBMenu = sReturn
    
    Erase iParents
    Erase sLines
    
    sCaption = vbNullString
    sName = vbNullString
    sTmp = vbNullString
    sReturn = vbNullString
    
Exit Function

ErrHan:
    
    Debug.Print "modVB.ParseVBMenu.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
Resume
End Function ' ... ParseVBMenu: String

Public Sub ParseVBMenuItem(pMenuItem As String, pMenuInfo As MenuInfo)

Dim sItems() As String
Dim sTmp As String

Dim j As Long
Dim i As Long
Dim k As Long

Dim xMnu As MenuInfo

    SplitString pMenuItem, sItems, Chr$(1), j
    
    If j Then
    
        For i = 0 To j - 1
            sTmp = sItems(i)
            k = CLng(Val(sTmp))
            With xMnu
                Select Case i
                    Case 0
                        .ID = k
                    Case 1
                        .NestLevel = k
                    Case 2
                        .ParentID = k
                    Case 3
                        .Name = sTmp
                        .MethodName = sTmp & "_Click"
                    Case 4
                        .Caption = sTmp
                End Select
            End With
        Next i
    
    End If
    
    pMenuInfo = xMnu
    
    Erase sItems
    
    sTmp = vbNullString
    
    i = 0
    j = 0
    k = 0
    
End Sub

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
'
''''      Begin VB.Menu mnuPSep1
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuPCopyProject
''''         Caption = "Copy Project"
''''      End
''''      Begin VB.Menu mnuPCopyFile
''''         Caption = "Copy File"
''''      End
''''      Begin VB.Menu mnuPSep2
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuPCompile
''''         Caption = "Compile"
''''         Visible = 0             'False
''''      End
''''      Begin VB.Menu mnuPReports
''''         Caption = "Text Reports"
''''         Begin VB.Menu mnuPQuickReport
''''            Caption = "Quick Report"
''''            Visible = 0             'False
''''         End
''''         Begin VB.Menu mnuPFullReport
''''            Caption = "Project Report"
''''         End
''''         Begin VB.Menu mnuPAPIReport
''''            Caption = "API Report (Actual)"
''''         End
''''         Begin VB.Menu mnuPAPIReportDistinct
''''            Caption = "API Report (Distinct)"
''''         End
''''      End
''''      Begin VB.Menu mnuPDevHelp
''''         Caption = "Dev. Dictionary"
''''      End
''''      Begin VB.Menu mnuTypesEtc
''''         Caption = "API-Constant-Type-Enum"
''''         Begin VB.Menu mnuPAPI
''''            Caption = "API"
''''         End
''''         Begin VB.Menu mnuPConstant
''''            Caption = "Constant"
''''         End
''''         Begin VB.Menu mnuPType
''''            Caption = "Type"
''''         End
''''         Begin VB.Menu mnuPEnum
''''            Caption = "Enum"
''''         End
''''      End
''''      Begin VB.Menu mnuSep3
''''         Caption = "-"
''''         Visible = 0             'False
''''      End
''''      Begin VB.Menu mnuPFindDeps
''''         Caption = "Find Dependencies"
''''         Visible = 0             'False
''''      End
''''      Begin VB.Menu mnuPSep1d
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuPSearch
''''         Caption = "Search Project"
''''      End
''''      Begin VB.Menu mnuPSep1b
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuPRefresh
''''         Caption = "Refresh"
''''      End
''''      Begin VB.Menu mnuPSep1c
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuPClose
''''         Caption = "Close"
''''      End
''''   End
''''   Begin VB.Menu mnuClass
''''      Caption = "Class"
''''      Begin VB.Menu mnuCCopySig
''''         Caption = "Copy Signature"
''''      End
''''      Begin VB.Menu mnuCCopyMethod
''''         Caption = "Copy Method"
''''      End
''''      Begin VB.Menu mnuCReports
''''         Caption = "Reports"
''''         Visible = 0             'False
''''      End
''''      Begin VB.Menu mnuCFullReport
''''         Caption = "Full Report"
''''         Visible = 0             'False
''''      End
''''      Begin VB.Menu mnuCRefresh
''''         Caption = "Refresh"
''''         Visible = 0             'False
''''      End
''''   End
''''   Begin VB.Menu mnuViewer
''''      Caption = "Viewer"
''''      NegotiatePosition = 3   'Right
''''      Begin VB.Menu mnuVShow
''''         Caption = "Show / Hide"
''''         Begin VB.Menu mnuVSep1
''''            Caption = "-"
''''            Visible = 0             'False
''''         End
''''         Begin VB.Menu mnuVProjExp
''''            Caption = "  Project Explorer"
''''            Shortcut        =   {F6}
''''         End
''''         Begin VB.Menu mnuVClassExp
''''            Caption = "  Class Explorer"
''''            Shortcut        =   {F7}
''''         End
''''         Begin VB.Menu mnuVToolBar
''''            Caption = "  ToolBar"
''''            Shortcut        =   {F8}
''''         End
''''         Begin VB.Menu mnuVStatusBar
''''            Caption = "  Status Bar"
''''            Shortcut        =   {F9}
''''         End
''''      End
''''      Begin VB.Menu mnuSep1b
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuSyntaxColuring
''''         Caption = "Syntax Colouring"
''''         Checked = -1            'True
''''      End
''''      Begin VB.Menu mnuVSep2
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuVInterface
''''         Caption = "Interface"
''''         Begin VB.Menu mnuVSep3
''''            Caption = "-"
''''            Visible = 0             'False
''''         End
''''         Begin VB.Menu mnuVIntPubOnly
''''            Caption = "  Public Members Only"
''''         End
''''         Begin VB.Menu mnuVIntAllMembs
''''            Caption = "  All Members"
''''         End
''''      End
''''      Begin VB.Menu mnuRSep1
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuCQuickReport
''''         Caption = "Quick Reporter"
''''      End
''''      Begin VB.Menu mnuRSep2
''''         Caption = "-"
''''      End
''''      Begin VB.Menu mnuRHist
''''         Caption = "History"
''''         Enabled = 0             'False
''''         Begin VB.Menu mnuRHItem
''''            Caption = ""
''''            Index = 0
''''         End
''''      End
''''   End
''''End
'
'
