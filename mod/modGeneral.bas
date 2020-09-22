Attribute VB_Name = "modGeneral"
Attribute VB_Description = "general program stuff"
' what?
'  general app specific stuff.
' why?
'
' when?
'
' how?

Option Explicit

' ... Memory stuff.
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "PSAPI" (ByVal hProcess As Long) As Long

Private Declare Function fShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function fGetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function fGetWinSysDir Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function fGetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' v8
' -------------------------------------------------------------------
' structure to describe elements of a PSC Read Me file
Public Type PSCInfo
    Name As String
    Description As String
    Link As String
    FileName As String
    UnzipFolder As String
    FileDate As Long
End Type
' -------------------------------------------------------------------
' v7/8
' Structure:    VariableInfo
' Description:  A Type / Structure to describe various attributes of a variable.

Public Type VariableInfo
    
    Accessor As Long                    ' ... accessor code
    AccessorAsString As String          ' ... Accessor as a string (derived from numeric accessor code)
    
    Declaration As String               ' ... the actual declaration.
    Name As String                      ' ... the name of the variable.
    Type As String                      ' ... the type of variable, as a string.
    
    LineStart As Long                   ' ... the line number in the source file.
    EditorLineStart As Long             ' ... the line number in the editor.
    
    ObjectWithEvents As Boolean
    
End Type ' ... VariableInfo

' -------------------------------------------------------------------
' v7/8
' Structure:   FindResultInfo.
' Description: A Type / Structure to describe various attributes of a find/search result.

Public Type FindResultInfo

    FileName As String  ' ... Source FIle Name.
    MemberName As String  ' ... Name of Member.
    FindLineText As String ' ... the entire line where the substring is found.
    FilePosition As Long  ' ... Start position of substring in file.
    MemberIndex As Long ' ... the index of the member in the source file's members.
    MemberPosition As Long  ' ... Start position of substring in Member.
    FileLineNumber As Long  ' ... Line position of substring in File.
    MemberLineNumber As Long  ' ... Line position of substring in Member.
    LinePosition As Long  ' ... Start position of substring on Line.
    SearchIndex As Long  ' ... An index for the search result in a list.
    ParentIndex As Long ' ... An index for the parent list.
    
End Type ' ... FindResultInfo.

' -------------------------------------------------------------------

' -------------------------------------------------------------------
' v7
Public Type MemberInfo

' ... this is designed around an existing format for the MembersStringArray in the CodeInfo Class.
' ... here I am adding string versions for both Accessor (e.g. Public, Private) and Type (e.g. Sub, Function)

    Accessor As Long                ' ... accessor code
    AccessorAsString As String      ' ... Accessor as a string (derived from numeric accessor code)
    Index As Long                   ' ... actual index of member in list of members.
    Name As String                  ' ... name of member
    Type As Long                    ' ... Type code
    TypeAsString As String          ' ... Type as a string (derived from numeric type code)
    ValueType As String             ' ... return value of a function or property get
    ' -------------------------------------------------------------------
    ' v8:   these are added so that we can extend the information available
    '       when just reading a member's stringarray item built in CodeInfo.pParseClass.
    LineStart As Long
    ParentName As String
    ParentFileName As String
    ParentTypeString As String
    MethodAttributes As String
    ' -------------------------------------------------------------------
End Type

Public Type ParamInfo
    
' ... this is designed around the rules for parameter declarations to methods/members.

    IsOptional As Boolean
    IsByRef As Boolean
    IsArray As Boolean
    IsParamArray As Boolean
    Name As String
    Type As String
    DefaultValue As String
    
End Type
' -------------------------------------------------------------------

' v6

' -------------------------------------------------------------------

' Enumerator:  ProjectExplorerNodesEnum.
' Description: An enumerator to help describe which nodes can be loaded into a project explorer tree view.

Public Enum ProjectExplorerNodesEnum

    eInfoNodes = 1  ' ... Project Properties
    eObjectsNode = 2  ' ... Refers to References and Components belonging to a project.
    eSourceFiles = 4  ' ... Project Source File members e.g. Forms, Classes, Modules ... etc
    eRelatedDocuments = 8  ' ... Refers to Related Documents added to the project.
    eAllNodes = 15
    
End Enum ' ... ProjectExplorerNodesEnum.

' -------------------------------------------------------------------

Public Type ConstInfo
    Declaration As String
    Name As String
    Scope As String
    Type As String
    Value As String
End Type
' -------------------------------------------------------------------

Public Type DataInfo
    Name As String
    Type As Byte
    Exists As Boolean
    ExtraInfo As String
    Index As Long           ' ... v8
End Type

Public Type APIInfo
    Declaration As String    ' ... v3/4.
    Scope As Byte
    Type As Byte
    Name As String
    Lib As String
    Alias As String
    Parameters As String
    ReturnValue As String
    ParentName As String
    ParentType As String
    LineNo As Long
    EditorLineNo As Long
End Type

    '   File Name | File Folder | Full Member Name | Date | Uncomp. Size | Comp. Size & | Zip Index

Public Type ZipMemberInfo
    FileName As String
    FilePath As String
    FullPathAndName As String
    FileDate As Date
    UnCompSize As Long
    CompSize As Long
    Index As Long
    Encrypted As Boolean
End Type

' -------------------------------------------------------------------
' ... a couple of structures to help resizing forms.
' ... v6, moved here to save repeat declarations.

Public Type CoDim  ' ... co-ordinates / dimensions
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Public Type MinMax ' ... Minimum / Maximum values.
    Min As Single
    max As Single
End Type

' -------------------------------------------------------------------

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVS_EX_FULLROWSELECT As Long = &H20
Private Const LVS_EX_CHECKBOXES As Long = &H4

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' text box api constants
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000
Private Const ECM_FIRST = &H1500                    ' ... Edit control messages.
Private Const EM_SETCUEBANNER = (ECM_FIRST + 1)

Sub SetTextBoxCueBanner(pTxtBxHwnd As Long, pCueText As String)
    
    SendMessageLongW pTxtBxHwnd, EM_SETCUEBANNER, 0, StrPtr(pCueText)
'    Let lResult = SendMessageLongW(txtFind.hwnd, EM_SETCUEBANNER, 0, StrPtr("Enter Search Text"))

End Sub

Sub ParsePSCInfo(ByVal pvReadMeText As String, _
                 ByRef prPSCInfo As PSCInfo)

Dim lngFound As Long
Dim lngNext As Long
Dim lngStart As Long
Dim sFind As String
    
    prPSCInfo.Name = pvReadMeText
    prPSCInfo.Description = pvReadMeText
    prPSCInfo.Link = pvReadMeText
    
    lngStart = 1
    sFind = ": "
    
    lngFound = InStr(lngStart, pvReadMeText, sFind)
    
    If lngFound > 0 Then
        lngStart = lngFound + Len(sFind)
        sFind = vbCrLf
        lngNext = InStr(lngStart, pvReadMeText, sFind)
        
        If lngNext > 0 Then
            prPSCInfo.Name = Mid$(pvReadMeText, lngStart, lngNext - lngStart)
            lngStart = lngNext + 1
            sFind = ": "
            lngFound = InStr(lngStart, pvReadMeText, sFind)
            If lngFound > 0 Then
                lngStart = lngFound + Len(sFind)
                sFind = "This file came from Planet-Source-Code.com...the home millions of lines of source code"
                lngNext = InStr(lngStart, pvReadMeText, sFind)
                If lngNext > 0 Then
                    prPSCInfo.Description = Mid$(pvReadMeText, lngStart, lngNext - lngStart)
                    lngStart = lngNext + 1
                    sFind = "You can view comments on this code/and or vote on it at:"
                    lngFound = InStr(lngStart, pvReadMeText, sFind)
                    If lngFound > 0 Then
                        lngStart = lngFound + Len(sFind)
                        sFind = vbCrLf
                        lngNext = InStr(lngStart, pvReadMeText, sFind)
                        If lngNext > 0 Then
                            prPSCInfo.Link = Mid$(pvReadMeText, lngStart, lngNext - lngStart)
                            prPSCInfo.Link = Trim$(prPSCInfo.Link)
                        End If
                    End If
                End If
            End If
        End If
        
    End If
        
End Sub


Sub ClearMemory()

Dim lngProcHandle As Long
Dim lngRet As Long
' -------------------------------------------------------------------
' Helper: clears up application memory use.
' -------------------------------------------------------------------
    On Error GoTo LogMemory_Err
    lngProcHandle = GetCurrentProcess()
    lngRet = EmptyWorkingSet(lngProcHandle)
LogMemory_Exit:
Exit Sub
LogMemory_Err:
End Sub


Public Sub ScrollRTFBox(pRTBHwnd As Long, Optional pScrollToLine As Long = 0)
Attribute ScrollRTFBox.VB_Description = "Scroll the Lines of a Rich Text Box to the Line Number provided."
    
Const EM_LINESCROLL = &HB6
    
    On Error GoTo ErrHan:
    
    SendMessageLong pRTBHwnd, EM_LINESCROLL, 0&, pScrollToLine

Exit Sub
ErrHan:

    Debug.Print "modGeneral.ScrollRTFBox.Error: " & Err.Number & "; " & Err.Description

End Sub

Public Function GetFirstVisibleLineRTFBox(pRTBHwnd As Long) As Long
Attribute GetFirstVisibleLineRTFBox.VB_Description = "Returns the first visible line of a Rich Text Box or -1."

Dim lngFirstLine As Long

Const EM_GETFIRSTVISIBLELINE = &HCE

    On Error GoTo ErrHan:
    
    If pRTBHwnd Then
        lngFirstLine = SendMessageLong(pRTBHwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    Else
        Err.Raise vbObjectError + 1000, , "No RTF Window Handle available"
    End If
ResumeError:
    
    GetFirstVisibleLineRTFBox = lngFirstLine
    lngFirstLine = 0&
    
Exit Function
ErrHan:

    lngFirstLine = -1
    Debug.Print "modGeneral.GetFirstVisibleLineRTFBox.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Public Function WordWrapRTFBox(pRTBHwnd As Long, Optional pWrap As Boolean = False) As Long
Attribute WordWrapRTFBox.VB_Description = "Toggle Word Wrap on a Rich Text Box, defaults to Off if no instruction provided."

Dim lRet As Long

Const WM_USER = &H400
Const EM_SETTARGETDEVICE = (WM_USER + 72)
    
    On Error GoTo ErrHan:
    
    If pRTBHwnd Then
        lRet = SendMessageLong(pRTBHwnd, EM_SETTARGETDEVICE, 0, IIf(pWrap = True, 0, 1))
    Else
        Err.Raise vbObjectError + 1000, , "No RTF Window Handle available"
    End If

ResumeError:
    
    WordWrapRTFBox = lRet
    lRet = 0&

Exit Function

ErrHan:
    lRet = -1
    Debug.Print "modGeneral.WordWrapRTFBox.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function

Public Function TypeDeclaration(ByVal pTypeDefinition As String) As String

Dim soTmp As StringArray
Dim soMembers As StringArray
Dim lngCount As Long
Dim lngLoop As Long
Dim sTmp As String
Dim sMember As String
Dim sName As String
Dim sDeclaration As String

    On Error GoTo ErrHan:
    
    If Len(pTypeDefinition) = 0 Then
        Err.Raise vbObjectError + 1000, , "Type Definition is empty"
    End If

    Set soTmp = New StringArray
    soTmp.FromString pTypeDefinition, ":"
    
    lngCount = soTmp.Count
    If lngCount <> 2 Then
        Err.Raise vbObjectError + 1000, , "Wrong number of items in top level of Type Definition"
    End If

    sName = soTmp.Item(1)
    Set soMembers = soTmp.ItemAsStringArray(2, ";")

    lngCount = soMembers.Count
    sDeclaration = "Type " & sName
    
    For lngLoop = 1 To lngCount
    
        sTmp = soMembers(lngLoop)
        sDeclaration = sDeclaration & vbNewLine & "    " & sTmp
    
    Next lngLoop
    
    sDeclaration = sDeclaration & vbNewLine & "End Type" & vbNewLine

ResumeError:
    
    TypeDeclaration = sDeclaration
    
    sName = vbNullString
    sDeclaration = vbNullString
    sTmp = vbNullString
    
    lngCount = 0&
    lngLoop = 0&
    
    If Not soTmp Is Nothing Then
        Set soTmp = Nothing
    End If
    
    If Not soMembers Is Nothing Then
        Set soMembers = Nothing
    End If
    
Exit Function

ErrHan:

    Debug.Print "modGeneral.TypeDeclaration.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Public Function EnumDeclaration(ByVal pEnumDefinition As String) As String
Attribute EnumDeclaration.VB_Description = "Returns an Enum Declaration from an in house Enum Definition."

Dim soTmp As StringArray
Dim soMembers As StringArray
Dim lngCount As Long
Dim lngLoop As Long
Dim sTmp As String
Dim sMember As String
Dim sName As String
Dim sDeclaration As String

    On Error GoTo ErrHan:
    
    If Len(pEnumDefinition) = 0 Then
        Err.Raise vbObjectError + 1000, , "Enum Definition is empty"
    End If

    Set soTmp = New StringArray
    soTmp.FromString pEnumDefinition, ":"
    
    lngCount = soTmp.Count
    If lngCount <> 2 Then
        Err.Raise vbObjectError + 1000, , "Wrong number of items in top level of Enum Definition"
    End If

    sName = soTmp.Item(1)
    Set soMembers = soTmp.ItemAsStringArray(2, ";")

    lngCount = soMembers.Count
    sDeclaration = "Enum " & sName
    
    For lngLoop = 1 To lngCount
    
        sTmp = soMembers(lngLoop)
        sDeclaration = sDeclaration & vbNewLine & "    " & sTmp
    
    Next lngLoop
    
    sDeclaration = sDeclaration & vbNewLine & "End Enum" & vbNewLine

ResumeError:
    
    EnumDeclaration = sDeclaration
    
    sName = vbNullString
    sDeclaration = vbNullString
    sTmp = vbNullString
    
    lngCount = 0&
    lngLoop = 0&
    
    If Not soTmp Is Nothing Then
        Set soTmp = Nothing
    End If
    
    If Not soMembers Is Nothing Then
        Set soMembers = Nothing
    End If
    
Exit Function

ErrHan:

    Debug.Print "modGeneral.EnumDeclaration.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Public Function TranslateEnumFunction(ByVal pEnumDefinition As String) As String
Attribute TranslateEnumFunction.VB_Description = "Generates a function to return a string for a given enum member from an internal single line Enum Definition."

Dim sKey As String
Dim sName As String
Dim lngFound As Long
Dim sEnums() As String
Dim lngCount As Long
Dim stext As String
Dim lngLoop As Long
Dim sMember As String
Dim sLeft As String
Dim sRight As String
Dim sFunction As String

' Note: would like to call this from different situations
'       hence writing to module.

    ' -------------------------------------------------------------------
    ' ... generate enum select case.
    ' -------------------------------------------------------------------
    ' ... note: enums key is single string in following format...
    ' ... name : mem1; mem2; mem3, ... etc.
    On Error GoTo ErrHan:
    
    sKey = pEnumDefinition
    
    lngFound = InStr(1, sKey, ":")
    
    If lngFound > 0 Then
        
        sName = Left$(sKey, lngFound - 1)
        
        sFunction = "Function Translate" & sName & "Member(ByVal pMember As " & sName & ") As String" & vbNewLine
'        sFunction = sFunction & "Dim sTmp As String" & vbNewLine & vbNewLine
        
       
        sKey = Mid$(sKey, lngFound + 1)
        
        SplitString sKey, sEnums, ";", lngCount
        ' ... shred the members.
        If lngCount > 0 Then
            
            stext = vbNullString
            lngFound = 0 ' ... reuse as auto increment for valueless enum members.
            
            For lngLoop = 0 To lngCount - 1
                
                If Len(stext) Then stext = stext & vbNewLine
                
                sMember = sEnums(lngLoop)
                
                modStrings.SplitStringPair sMember, "=", sLeft, sRight, True, True
                ' ... sleft = member name, sright = member value.
                
                stext = stext & "        Case " & sLeft
                If Len(sRight) Then
                    stext = stext & " ' ... " & sRight & "."
                    lngFound = lngFound + 1
                Else
                    stext = stext & " ' ... 0 based Index " & lngLoop & IIf(lngFound > 0, " ( minus " & lngFound & " = " & CStr(lngLoop - lngFound) & " )", "") & "."
                End If
                stext = stext & vbNewLine & "            sTmp = " & Chr$(34) & sLeft & Chr$(34)
                
            Next lngLoop
            
            stext = "Dim sTmp As String" & vbNewLine & vbNewLine & "    Select Case pMember" & vbNewLine & vbNewLine & stext & vbNewLine & vbNewLine & "    End Select" & vbNewLine & vbNewLine
            stext = stext & "    Translate" & sName & "Member = sTmp" & vbNewLine & vbNewLine
        
        End If
    
    End If
    
    TranslateEnumFunction = sFunction & stext & "End Function"

ResumeError:
    
    On Error Resume Next
    
    stext = vbNullString
    sKey = vbNullString
    sName = vbNullString
    sMember = vbNullString
    sLeft = vbNullString
    sRight = vbNullString
    
    Erase sEnums
    
    lngFound = 0&
    lngLoop = 0&
    lngCount = 0&
    
Exit Function

ErrHan:

    Debug.Print "modGeneral.TranslateEnumFunction.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function


Public Function SelectCaseFromEnum(ByVal pEnumDefinition As String) As String
Attribute SelectCaseFromEnum.VB_Description = "Generates a Select Case (enum member) Statement Block from an internal single line Enum Definition."

Dim sKey As String
Dim sName As String
Dim lngFound As Long
Dim sEnums() As String
Dim lngCount As Long
Dim stext As String
Dim lngLoop As Long
Dim sMember As String
Dim sLeft As String
Dim sRight As String

' Note: would like to call this from different situations
'       hence writing to module.

    ' -------------------------------------------------------------------
    ' ... generate enum select case.
    ' -------------------------------------------------------------------
    ' ... note: enums key is single string in following format...
    ' ... name : mem1; mem2; mem3, ... etc.
    On Error GoTo ErrHan:
    
    sKey = pEnumDefinition
    
    lngFound = InStr(1, sKey, ":")
    
    If lngFound > 0 Then
        
        sName = Left$(sKey, lngFound - 1)
       
        sKey = Mid$(sKey, lngFound + 1)
        
        SplitString sKey, sEnums, ";", lngCount
        ' ... shred the members.
        If lngCount > 0 Then
            
            stext = vbNullString
            lngFound = 0 ' ... reuse as auto increment for valueless enum members.
            
            For lngLoop = 0 To lngCount - 1
                
                If Len(stext) Then stext = stext & vbNewLine
                
                sMember = sEnums(lngLoop)
                
                modStrings.SplitStringPair sMember, "=", sLeft, sRight, True, True
                ' ... sleft = member name, sright = member value.
                
                stext = stext & "        Case " & sLeft
                If Len(sRight) Then
                    stext = stext & " ' ... " & sRight & "."
                    lngFound = lngFound + 1
                Else
                    stext = stext & " ' ... 0 based Index " & lngLoop & IIf(lngFound > 0, " ( minus " & lngFound & " = " & CStr(lngLoop - lngFound) & " )", "") & "."
                End If
                stext = stext & vbNewLine & "            "
                
            Next lngLoop
            
            stext = "Dim x As " & sName & vbNewLine & vbNewLine & "    Select Case x" & vbNewLine & vbNewLine & stext & vbNewLine & "    End Select"
        
        End If
    
    End If
    
    SelectCaseFromEnum = stext

ResumeError:
    
    On Error Resume Next
    
    stext = vbNullString
    sKey = vbNullString
    sName = vbNullString
    sMember = vbNullString
    sLeft = vbNullString
    sRight = vbNullString
    
    Erase sEnums
    
    lngFound = 0&
    lngLoop = 0&
    lngCount = 0&
    
Exit Function

ErrHan:

    Debug.Print "modGeneral.SelectCaseFromEnum.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function

Public Sub AutosizeColumns(ByVal TargetListView As ListView)
Attribute AutosizeColumns.VB_Description = "Auto Size all the Columns of a List View."

' just found this at
' http://binaryworld.net/Main/CodeDetail.aspx?CodeId=3640&atlanta=software%20development

' ... need a way to avoid resizing hidden columns, those we didn't want seen, else fab.


  Const SET_COLUMN_WIDTH  As Long = 4126
  Const AUTOSIZE_USEHEADER As Long = -2

  Dim lngColumn As Long

  For lngColumn = 0 To (TargetListView.ColumnHeaders.Count - 1)

    Call SendMessage(TargetListView.hwnd, _
        SET_COLUMN_WIDTH, _
        lngColumn, _
        ByVal AUTOSIZE_USEHEADER)

  Next lngColumn

End Sub

Sub LVSortTextCol(pLV As ListView, ByVal pSortKey As Long)
' sort list view by text column
Dim iSortOrder As Long
    On Error GoTo ErrHan:
    ' trap an invalid object error.
    iSortOrder = pLV.SortOrder
    ' reverse current sort order.
    iSortOrder = IIf(iSortOrder = lvwAscending, lvwDescending, lvwAscending)
    pLV.SortOrder = iSortOrder
    pLV.SortKey = pSortKey
    pLV.Sorted = True
Exit Sub
ErrHan:
    Debug.Print "modCtrls.LVSortTextCol.Error: " & Err.Number & "; " & Err.Description
End Sub

Public Sub LVSizeColumn(ByVal pLV As ListView, Optional ByVal pColIndex As Long = 0)
Attribute LVSizeColumn.VB_Description = "Auto Size a Single Column in a List View."

' just found this at
' http://binaryworld.net/Main/CodeDetail.aspx?CodeId=3640&atlanta=software%20development

' ... need a way to avoid resizing hidden columns, those we didn't want seen, else fab.


Const SET_COLUMN_WIDTH  As Long = 4126
Const AUTOSIZE_USEHEADER As Long = -2

    If Not pLV Is Nothing Then
        If pColIndex >= 0 And pColIndex < pLV.ColumnHeaders.Count Then ' - 1 Then
            Call SendMessage(pLV.hwnd, SET_COLUMN_WIDTH, pColIndex, ByVal AUTOSIZE_USEHEADER)
        
        End If
    End If
    
End Sub

Public Function BinaryNumberSearch(ByVal pNumberToFind As Long, ByRef pSourceNumberArray() As Long) As Long
Attribute BinaryNumberSearch.VB_Description = "Binary Search algorithm on an array of longs to find first matching search number within the array, returns -1 if number not found."

' ... try and find a source array index representing the first value larger
' ... than the number to find.
' ... the number to find is a Line Number and the source array is an array of
' ... starting positions of methods and members...
' ... in this way we can find which method a line belongs to.

' ... using a binary search algorithm.
' ... essentially, the idea is to split the source data in two, find which section
' ... top or bottom a number belongs in and then repeat on that section, and so on and so forth.

Dim lngUpper As Long
Dim lngLow As Long
Dim lngPos As Long
Dim lngFindValue As Long

    On Error GoTo ErrHan:
    
    lngPos = -1 ' ... default, not found or error.
    
    ' ... test the source array.
    lngLow = LBound(pSourceNumberArray)
    lngUpper = UBound(pSourceNumberArray)
    
    Do While True
        
        ' ... find the middle of the source array.
        lngPos = (lngLow + lngUpper) / 2
        
        If pSourceNumberArray(lngPos) = pNumberToFind Then
            
            ' ... found it.
            BinaryNumberSearch = lngPos
            Exit Function
        
        End If
        
        ' ... see if its the last to be checked.
        If lngUpper = lngLow + 1 Then
            
            lngLow = lngUpper
        
        Else
            
            If lngPos = lngLow Then
            
                ' ... if we get to the lowest position its not there.
                BinaryNumberSearch = -1
                Exit Function
            
            Else
                
                ' ... determine whether to look in the upper or lower
                ' ... part of the array.
                If pSourceNumberArray(lngPos) > pNumberToFind Then
                    
                    lngUpper = lngPos
                
                Else
                    
                    lngLow = lngPos
                
                End If
            
            End If
            
        End If
    
    Loop


ResumeError:

    BinaryNumberSearch = lngPos

Exit Function

ErrHan:

    lngPos = -1
    Debug.Print "modGeneral.BinaryNumberSearch.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Public Sub LVFullRowSelect(pLVHwnd As Long)
Attribute LVFullRowSelect.VB_Description = "Attempts to instruct a vb5 ListView to implement Full Row Select."

    Call SendMessageLong(pLVHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, LVS_EX_FULLROWSELECT) ' Or LVS_EX_CHECKBOXES)

End Sub

Public Sub SaveFormPosition(ByRef pForm As Form)
Attribute SaveFormPosition.VB_Description = "Attempts to save the position of a form to the registry for later reloading."

    If pForm.WindowState = vbNormal Then
        
        SaveSetting App.Title, pForm.Name, "Top", CLng(pForm.Top)
        SaveSetting App.Title, pForm.Name, "Left", CLng(pForm.Left)
        SaveSetting App.Title, pForm.Name, "Width", CLng(pForm.Width)
        SaveSetting App.Title, pForm.Name, "Height", CLng(pForm.Height)
    
    End If

End Sub

Public Sub LoadFormPosition(ByRef pForm As Form, _
                            ByVal pContainerHeight As Long, _
                            ByVal pContainerWidth As Long)
Attribute LoadFormPosition.VB_Description = "Attempts to read form position data from registry and apply it to the form when loaded."

Dim lngTop As Long
Dim lngLeft As Long
Dim lngWidth As Long
Dim lngHeight As Long

    If pForm.BorderStyle = 2 Then

        lngTop = CLng(Val(GetSetting(App.Title, pForm.Name, "Top", "300")))
        lngLeft = CLng(Val(GetSetting(App.Title, pForm.Name, "Left", "300")))
        lngWidth = CLng(Val(GetSetting(App.Title, pForm.Name, "Width", "9060")))
        lngHeight = CLng(Val(GetSetting(App.Title, pForm.Name, "Height", "7305")))
    
        If lngTop + lngHeight > pContainerHeight Then
            lngTop = 300
        End If
    
        If lngLeft + lngWidth > pContainerWidth Then
            lngLeft = 300
        End If

        pForm.Move lngLeft, lngTop, lngWidth, lngHeight
            
    End If

End Sub

Public Sub CentreForm(ByRef frm As Form)
Attribute CentreForm.VB_Description = "Attempts to centre an MDI Child Form in its parent mdi container."
Dim RCT As RECT
Dim x As Variant, y As Variant
    With frm
        If .MDIChild Then Call GetClientRect(GetParent(.hwnd), RCT)
        x = (((RCT.Right - RCT.Left) * Screen.TwipsPerPixelX) - .Width) / 2
        y = (((RCT.Bottom - RCT.Top) * Screen.TwipsPerPixelY) - .Height) / 2
        .Move x, y
    End With
End Sub

Public Sub InitCommonControlsVB()
Attribute InitCommonControlsVB.VB_Description = "Attempt to install Themes into the program (themes may require a manifest file)."

Dim x As tagInitCommonControlsEx

    On Error Resume Next

    With x
        .lngSize = LenB(x)
        .lngICC = ICC_USEREX_CLASSES
    End With

    InitCommonControlsEx x

    If Err Then
        Err.Clear
        On Error GoTo 0
        InitCommonControls
    End If

End Sub


Public Function GetFileLength(pTheFileName As String, Optional prFileLength As Long) As String
Attribute GetFileLength.VB_Description = "Attempts to return the length of a file as a string e.g. 1,024 bytes.  prFileLength returns the size as a Long.  If file size > Long file size returns 0 and function returns "" > 2Gb?""."
Dim lngLen As Long
Dim SLen As String
    
    On Error GoTo ErrHan:
    SLen = "N/A"
    If Dir$(pTheFileName) <> "" Then
        lngLen = FileLen(pTheFileName)
        SLen = Format$(lngLen, cNumFormat) & " bytes"
    End If
ErrRes:
    GetFileLength = SLen
    prFileLength = lngLen
Exit Function
ErrHan:
    ' ... could error above if file size is > Long value.
    SLen = " > 2Gb?"
    Resume ErrRes:
End Function

Function FormatBytes(ByVal ByteSize As Double) As String 'Original code Ben White
    Dim x As Long, y As Long, z As Double
    Dim sBytes As String, sUnits As String
    Do: x = x + 1
       z = 2 ^ (x * 10)
       For y = 1 To 3
          If ByteSize < z * (10 ^ y) Then
             sBytes = FormatNum(ByteSize / z, 3 - y)
             Exit Do
    End If: Next: Loop
    sUnits = Choose(x, "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB", "??", "??", "??", "??")
    FormatBytes = Space$(5 - Len(sBytes)) & sBytes & "  " & sUnits & " "
End Function

Private Function FormatNum(ByVal dNumber As Double, Optional ByVal lDecPlaces As Long = -1) As String
    Select Case lDecPlaces
        Case Is < 0
            FormatNum = CStr(CDec(dNumber))
        Case Is = 0
            FormatNum = Format$(CStr(CDec(dNumber)), "#0")
        Case Else
            FormatNum = Format$(CStr(CDec(dNumber)), "#0." & String$(lDecPlaces, "0"))
    End Select
End Function

Public Function AddressOfMethod(ByVal pAddressOf As Long) As Long
Attribute AddressOfMethod.VB_Description = "Returns a value that is the AddressOf the method."
   AddressOfMethod = pAddressOf
End Function

Public Function CheckUnzipAPI() As Boolean
Attribute CheckUnzipAPI.VB_Description = "Checks to see if Unzip32.dll is in the System32 directory and returns true if found."
Dim sFileName As String
' ... could do better.
    sFileName = GetWinSysDir
    If Len(sFileName) Then
        sFileName = sFileName & "\Unzip32.dll"
        If Dir$(sFileName, vbNormal) <> "" Then
            CheckUnzipAPI = True
        End If
    End If
    
End Function

Public Function CheckZipAPI() As Boolean
Attribute CheckZipAPI.VB_Description = "Checks to see if Zip32.dll is in the System32 directory and returns true if found."

Dim sFileName As String
' ... could do better.
    sFileName = GetWinSysDir
    If Len(sFileName) Then
        sFileName = sFileName & "\Zip32.dll"
        If Dir$(sFileName, vbNormal) <> "" Then
            CheckZipAPI = True
        End If
    End If
    
End Function

Public Function GetWindowsDirectory() As String
Attribute GetWindowsDirectory.VB_Description = "Returns the windows directory."

Dim sTmpBuffer As String
Dim lngRetLen As Long

    sTmpBuffer = String$(255, 0)
    lngRetLen = fGetWindowsDirectory(sTmpBuffer, 255)
    If lngRetLen > 0 Then
        sTmpBuffer = Left$(sTmpBuffer, lngRetLen)
        GetWindowsDirectory = sTmpBuffer
    End If
    
    sTmpBuffer = vbNullString
    lngRetLen = 0
    
End Function

Public Function GetWinSysDir() As String
Attribute GetWinSysDir.VB_Description = "Returns the windows system 32 directory."

Dim sTmpBuffer As String
Dim lngRetLen As Long

    sTmpBuffer = String$(255, 0)
    lngRetLen = fGetWinSysDir(sTmpBuffer, 255)
    If lngRetLen > 0 Then
        sTmpBuffer = Left$(sTmpBuffer, lngRetLen)
        GetWinSysDir = sTmpBuffer
    End If
    
    sTmpBuffer = vbNullString
    lngRetLen = 0
    
End Function

Public Function GetTempPath() As String
Attribute GetTempPath.VB_Description = "Attempts to return the Temp path on Windows OS."

Dim sTmpBuffer As String
Dim lngRetLen As Long

    sTmpBuffer = String$(255, 0)
    lngRetLen = fGetTempPath(255, sTmpBuffer)
    If lngRetLen > 0 Then
        sTmpBuffer = Left$(sTmpBuffer, lngRetLen)
        GetTempPath = sTmpBuffer
    End If
    
    sTmpBuffer = vbNullString
    lngRetLen = 0
    
End Function

Public Sub OpenWebPage(pAddress As String)
Attribute OpenWebPage.VB_Description = "Opens a Web Page or local htm/html document."
    fShellExecute 0, "Open", pAddress, "", "", 1
End Sub

Public Sub OpenFolder(pFolder As String)
    fShellExecute 0&, vbNullString, pFolder & "\", vbNullString, vbNullString, vbNormalFocus
End Sub

Public Sub RunProgram(pFile As String)
    fShellExecute 0&, vbNullString, pFile, vbNullString, vbNullString, vbNormalFocus
End Sub

' -------------------------------------------------------------------
' v6
Public Sub ParseConstantsItem(ByVal pConstDec As String, pConstInfo As ConstInfo)
Attribute ParseConstantsItem.VB_Description = "Attempts to parse a string into a ConstInfo structure / type."

' some examples
'Private Const LVM_FIRST As Long = &H1000
'Private Const ECM_FIRST = &H1500
'Public Const TIP_FILE = "Tips.txt"
'Public Const c_word_Property As String = "Property"

' note: constants are expected to have been written by CodeInfo
'       check examples above for general required format
'       coding: first scope draft.

Dim sFind As String
Dim lngFound As Long
Dim sTmp As String
Dim sTmpDec As String
Dim bOK As Boolean
Dim sWorkingDec As String

'Dim pConstInfo As ConstInfo

    pConstInfo.Declaration = pConstDec
    pConstInfo.Type = "[Variant]"
    
    If Len(pConstDec) Then
        sWorkingDec = pConstDec
        sTmpDec = UCase$(pConstDec)
        bOK = True
    End If

    If bOK Then
        ' -------------------------------------------------------------------
        ' ... looking for accessor.
        sFind = " "
        lngFound = InStr(1, sTmpDec, sFind)
        If lngFound > 0 Then
            sTmp = Left$(sTmpDec, lngFound - 1)
            Select Case sTmp
                Case "PRIVATE", "PUBLIC"
                    sTmpDec = Mid$(sTmpDec, lngFound + 1)
                    sWorkingDec = Mid$(sWorkingDec, lngFound + 1)
                    ' -------------------------------------------------------------------
                    pConstInfo.Scope = Left$(pConstDec, lngFound - 1)
                Case "CONST"
                    pConstInfo.Scope = "Public"
                Case Else
                    bOK = False
            End Select
        End If
    
    End If

    If bOK Then
        ' -------------------------------------------------------------------
        ' ... looking for word Const.
        If Left$(sTmpDec, 6) = "CONST " Then
            sTmpDec = Mid$(sTmpDec, 7)
            sWorkingDec = Mid$(sWorkingDec, 7)
        Else
            bOK = False
        End If
    
    End If

    If bOK Then
        ' -------------------------------------------------------------------
        ' ... looking for a name.
        sFind = " "
        lngFound = InStr(1, sTmpDec, sFind)
        If lngFound > 0 Then
'            pConstInfo.Name = Left$(sTmpDec, lngFound - 1)
            pConstInfo.Name = Left$(sWorkingDec, lngFound - 1)
            sTmpDec = Mid$(sTmpDec, lngFound + 1)
            sWorkingDec = Mid$(sWorkingDec, lngFound + 1)
        Else
            bOK = False
        End If
    End If
    
    If bOK Then
        
        If Left$(sTmpDec, 3) = "AS " Then
            sTmpDec = Mid$(sTmpDec, 4)
            sWorkingDec = Mid$(sWorkingDec, 4)
            lngFound = InStr(1, sTmpDec, sFind)
            If lngFound > 0 Then
                pConstInfo.Type = Left$(sWorkingDec, lngFound - 1) ' Left$(sTmpDec, lngFound - 1)
                sTmpDec = Mid$(sTmpDec, lngFound + 1)
                sWorkingDec = Mid$(sWorkingDec, lngFound + 1)
            End If
        End If
        
        If Left$(sTmpDec, 2) = "= " Then
            sTmpDec = Mid$(sTmpDec, 3)
            sWorkingDec = Mid$(sWorkingDec, 3)
'            pConstInfo.Value = modStrings.LeftOfComment(sTmpDec)
            pConstInfo.Value = modStrings.LeftOfComment(sWorkingDec)
        End If
        
        lngFound = InStrRevChar(pConstInfo.Declaration, "|")
        If lngFound > 0 Then
            pConstInfo.Declaration = LeftOfComment(pConstInfo.Declaration, "|", True)
        End If
    
    End If

End Sub

' -------------------------------------------------------------------


Public Sub ParseAPIInfoItem(pSArray As StringArray, pIndex As Long, ByRef pTheAPIInfo As APIInfo)
Attribute ParseAPIInfoItem.VB_Description = "Converts a string into an APIInfoType for processing API declarations read from a CodeInfo instance."
' ... shred an in-house api declaration into manageable Type parts.
Dim lngCount As Long
Dim sTmpAPIDec As String
Dim sTmpData As StringArray
Dim tAPIInfo As APIInfo
Dim lngFound As Long
Dim sFind As String
    
    If Not pSArray Is Nothing Then
        lngCount = pSArray.Count
        If lngCount > 0 Then
            If pIndex <= lngCount Then
                ' ... read the string array item returning it as a new string array item.
                Set sTmpData = pSArray.ItemAsStringArray(pIndex, "|")
                ' 1 = dec, 2 = type, 3 = access/scope, 4 = line no, 5 = editor line no, 6 = source name, 7 = source type
                sTmpAPIDec = sTmpData(1)
                
                If sTmpData.IndexExists(4) Then
                    tAPIInfo.LineNo = CLng(sTmpData.ItemAsNumberValue(4))
                End If
                If sTmpData.IndexExists(5) Then
                    tAPIInfo.EditorLineNo = CLng(sTmpData.ItemAsNumberValue(5))
                End If
                If sTmpData.IndexExists(6) Then
                    tAPIInfo.ParentName = sTmpData(6)
                End If
                If sTmpData.IndexExists(7) Then
                    tAPIInfo.ParentType = sTmpData(7)
                End If
                
                tAPIInfo.Type = CByte(sTmpData.ItemAsNumberValue(2))
                tAPIInfo.Scope = CByte(sTmpData.ItemAsNumberValue(3))
                
                ' v3/4 ... provide a declaration for copying.
                If tAPIInfo.Type = 1 Then
                    tAPIInfo.Declaration = "Declare Sub " & sTmpAPIDec
                ElseIf tAPIInfo.Type = 2 Then
                    tAPIInfo.Declaration = "Declare Function " & sTmpAPIDec
                End If
                sFind = " "
                lngFound = InStr(1, sTmpAPIDec, sFind)
                If lngFound > 0 Then
                
                    tAPIInfo.Name = Left$(sTmpAPIDec, lngFound - 1)
                                    
                    ' ... attempt to extract a Lib.
                    sFind = " Lib "
                    lngFound = InStr(lngFound, sTmpAPIDec, sFind)
                    If lngFound > 0 Then
                        tAPIInfo.Lib = Mid$(sTmpAPIDec, lngFound + 1 + Len(sFind))
                        lngFound = InStr(1, tAPIInfo.Lib, Chr$(34))
                        If lngFound > 0 Then
                            tAPIInfo.Lib = Left$(tAPIInfo.Lib, lngFound - 1)
                        End If
                        sFind = " Alias "
                        ' ... attempt to extract an Alias.
                        lngFound = InStr(lngFound, sTmpAPIDec, sFind)
                        If lngFound > 0 Then
                            tAPIInfo.Alias = Mid$(sTmpAPIDec, lngFound + 1 + Len(sFind))
                            lngFound = InStr(1, tAPIInfo.Alias, Chr$(34))
                            If lngFound > 0 Then
                                tAPIInfo.Alias = Left$(tAPIInfo.Alias, lngFound - 1)
                            End If
                        End If
                        sFind = "("
                        ' ... attempt to extract any parameters passed.
                        lngFound = InStr(lngFound + 1, sTmpAPIDec, sFind)
                        If lngFound > 0 Then
                            tAPIInfo.Parameters = Mid$(sTmpAPIDec, lngFound + Len(sFind))
                            sFind = ")"
                            lngFound = modStrings.InStrRevChar(tAPIInfo.Parameters, sFind) ' ... use instrrevchar when only one char to find.
                            If lngFound > 0 Then
                                tAPIInfo.Parameters = Left$(tAPIInfo.Parameters, lngFound - 1)
                                If tAPIInfo.Type = 2 Then
                                    sFind = ") As "
                                    lngFound = modStrings.InstrRev(sTmpAPIDec, sFind)
                                    If lngFound > 0 Then
                                        tAPIInfo.ReturnValue = Right$(sTmpAPIDec, Len(sTmpAPIDec) - lngFound - Len(sFind) + 1)
                                        tAPIInfo.ReturnValue = modStrings.LeftOfComment(tAPIInfo.ReturnValue)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If Not sTmpData Is Nothing Then
        Set sTmpData = Nothing
    End If
    
    pTheAPIInfo = tAPIInfo

End Sub

Public Sub ParseDataInfoItem(pSArray As StringArray, _
                             pIndex As Long, _
                             pDataInfo As DataInfo, _
                    Optional pAttributeDelimiter As String = "|")
Attribute ParseDataInfoItem.VB_Description = "Converts a Project Data Info String into a DataInfo item."

' ... Splits an item in a string array into a DataInfo type record; used when loading project explorer nodes.
' ... When a Project is parsed its info is stored in a row of delimited strings.
' ... Each row contains the following information (delimited by a pipe/bang symbol ' | '):
'       Name: Name of the project object.
'       Type: Type of Project Object
'       Exists: If a File, True when exists else False.
'       Extra Info: File Name or GUID (for now).
' ... When the Project info is read, line by line, each one is sent here
' ... to be converted into a DataInfo Type which then makes using the info more easy.

'   In :    pSArray                 ' ... the string array to read.
'   In :    pIndex                  ' ... the index of the item in the string array to parse.
'   Out:    pDataInfo               ' ... the DataInfo Type to be returned.
'   In :    pAttributeDelimiter     ' ... the delimiter used to, err, delimit the project info string.

' ... Example:
' ... Read namesd contents of VBP.
' ... pSSArray is a string array with the vbp info as described above.
'        If Not pSSArray Is Nothing Then
'            lngCount = pSSArray.Count
'            If lngCount > 0 Then
'                For lngLoop = 1 To lngCount
'
'                    ParseVBPInfoItem pSSArray, lngLoop, tDataInfo
'
'                    Print "Project Item " & lngLoop
'                    Print "Name: " & tDataInfo.Name
'                    Print "Type: " & tDataInfo.Type
'                    Print "Exists: " & tDataInfo.Exists
'                    Print "FileName / GUID: " & tDataInfo.ExtraInfo
'                    Print
'
'                Next lngLoop
'            End If
'        End If

Dim sTmpData As StringArray
Dim tDataInfo As DataInfo
Dim lngNumVal As Long
Dim sAttributeDelimiter As String
' ... v8
Dim lngFindIndex As Long
Dim sExtraInfo As String
'Dim lngIndex As Long
Dim sTmpInfo As String

    If Not pSArray Is Nothing Then
        If pSArray.Count > 0 Then
            If pIndex <= pSArray.Count Then
                sAttributeDelimiter = "|"
                If pAttributeDelimiter <> "" Then
                    sAttributeDelimiter = pAttributeDelimiter
                End If
                Set sTmpData = pSArray.ItemAsStringArray(pIndex, sAttributeDelimiter)
                With tDataInfo
                    .Name = sTmpData(1)
                    lngNumVal = sTmpData.ItemAsNumberValue(2)
                    If lngNumVal < 256 Then
                        .Type = CByte(lngNumVal)
                    End If
                    .Exists = sTmpData(3) = "1"
                    ' -------------------------------------------------------------------
                    ' ... v8, added #[Index] to end of data info string in VBProject
                    ' ... may need to extract this and trim the path string.
                    sExtraInfo = sTmpData(4)
                    lngFindIndex = InStr(1, sExtraInfo, "#")
                    If lngFindIndex > 0 Then
                        sTmpInfo = Mid$(sExtraInfo, lngFindIndex + 1)
                        sExtraInfo = Left$(sExtraInfo, lngFindIndex - 1)
                        ' -------------------------------------------------------------------
                        ' ... read the index from VBProject amendment.
                        .Index = CLng(Val(sTmpInfo))
                    End If
                    .ExtraInfo = sExtraInfo
                    ' -------------------------------------------------------------------
                End With
            End If
        End If
    End If
    ' -------------------------------------------------------------------
    pDataInfo = tDataInfo
    ' -------------------------------------------------------------------
    If Not sTmpData Is Nothing Then
        Set sTmpData = Nothing
    End If
    
    sTmpInfo = vbNullString
    sExtraInfo = vbNullString
    lngFindIndex = 0&
    
End Sub ' ... ParseDataInfoItem:

Public Sub ParseMemberInfoItem(pSArray As StringArray, pIndex As Long, pMemberInfo As MemberInfo)
Attribute ParseMemberInfoItem.VB_Description = "Converts a string into a MemberInfo for processing items read from a CodeInfo MembersStringArray Instance."

Dim sTmpData As StringArray
Dim tMemberInfo As MemberInfo
' -------------------------------------------------------------------
' Helper:   convert stringarray data to MemberInfo for adding stuff to the project explorer.
' Note:     the data in the strinarray item (index) is a delimited text string
'           representing info about an item found in the vbp.
' -------------------------------------------------------------------
'''' ...                    Name     :    Quick Member index     :    Type             :    Accessor             :    Value/Return Type
'''m_aMethods.AddItemString sName & ":" & CStr(m_MemberCount) & ":" & CStr(lngType) & ":" & CStr(lngAccessor) & ":" & sValueType
    
' ... v8, format
'''' ...                    Name     |    Quick Member index     |    Type             |    Accessor             |    Value/Return Type   |  Line Start
'''m_aMethods.AddItemString sName & "|" & CStr(m_MemberCount) & "|" & CStr(lngType) & "|" & CStr(lngAccessor) & "|" & sValueType & "|" & lngLineStart
    
    If Not pSArray Is Nothing Then
        If pSArray.Count > 0 Then
            If pIndex <= pSArray.Count Then
'                Set sTmpData = pSArray.ItemAsStringArray(pIndex, ":")
                Set sTmpData = pSArray.ItemAsStringArray(pIndex, "|") ' ... v8, update.
                With tMemberInfo
                    
                    .Name = sTmpData(1)
                    .Index = CLng(sTmpData.ItemAsNumberValue(2))
                    .Type = CLng(sTmpData.ItemAsNumberValue(3))
                    .Accessor = CLng(sTmpData.ItemAsNumberValue(4))
                    .ValueType = sTmpData(5)
                    ' -------------------------------------------------------------------
                    ' ... v8, added LineStart at pos 6 ...
                    .LineStart = CLng(sTmpData.ItemAsNumberValue(6))
                    
                    ' -------------------------------------------------------------------
                    ' ... v8
                    ' ... VBProject may add source object name and full file name and path.
                    If sTmpData.IndexExists(7) Then
                        .ParentName = sTmpData(7)
                    End If
                    
                    If sTmpData.IndexExists(8) Then
                        .ParentFileName = sTmpData(8)
                    End If
                    If sTmpData.IndexExists(9) Then
                        .MethodAttributes = sTmpData(9)
                    End If
                    If sTmpData.IndexExists(10) Then
                        .ParentTypeString = sTmpData(10)
                    End If
                    
                    ' -------------------------------------------------------------------
                    
                    .AccessorAsString = Choose(.Accessor, "Public", "Private", "Friend")
                    .TypeAsString = Choose(.Type, "Sub", "Function", "Property")
                    
                End With
            End If
        End If
    End If
    ' -------------------------------------------------------------------
    pMemberInfo = tMemberInfo
    ' -------------------------------------------------------------------
    If Not sTmpData Is Nothing Then
        Set sTmpData = Nothing
    End If

End Sub ' ... ParseMemberInfoItem:

Public Sub ParseVariableInfoItem(ByVal pVarDec As String, pVarInfo As VariableInfo)
Attribute ParseVariableInfoItem.VB_Description = "Parse a delimited variable declaration into a variable info structure."

' ... Parse a delimited variable declaration into a variable info structure.

' ... example output from codeinfo :: pParseClass
'    m_Value As Long|0|2|54|39
'    mInitialised As Boolean|0|2|30|15
'    moVBPInfo As VBPInfo|0|2|31|16
'    mAttributeDelimiter As String|0|2|32|17
'    m_StringArray() As String|0|2|176|159
'    m_Row As Long|0|2|178|161
'    m_Count As Long|0|2|180|163
'    m_DuplicatesAllowed As Boolean|0|2|181|164
'    m_Sortable As Boolean|0|2|182|165
'    m_bCompacted As Boolean|0|2|183|166
'    m_PreAllocationSize As Long|0|2|184|167
'    m_TheBufferCapacity As Long|0|2|188|171
'    m_TheBufferChunkSize As Long|0|2|189|172
'    m_TheString As String|0|2|190|173
'    m_TheTextLength As Long|0|2|191|174
'    Name As String|0|1|203|186
'    Tag As String|0|1|204|187
'    m_TheString As String|0|2|120|103


Dim saTmp As StringArray
Dim sTmpDec As String
Dim tPInfo As ParamInfo

    ' -------------------------------------------------------------------
    ' ... clean up incoming var info structure, in case of re-use.
    With pVarInfo
        
        .Accessor = 0
        .AccessorAsString = vbNullString
        .Declaration = vbNullString
        .EditorLineStart = 0
        .LineStart = 0
        .Name = vbNullString
        .Type = vbNullString
        .ObjectWithEvents = False
        
    End With
    
    If Len(Trim$(pVarDec)) = 0 Then Exit Sub
    
    Set saTmp = New StringArray
    saTmp.FromString pVarDec, "|", True
    
    sTmpDec = saTmp(1)
    
    If saTmp(3) = "2" Then
        sTmpDec = Mid$(sTmpDec, c_len_Private + 2)
    ElseIf saTmp(3) = "1" Then
        sTmpDec = Mid$(sTmpDec, c_len_Public + 2)
    End If
    
    If Left$(sTmpDec, Len("WithEvents ")) = "WithEvents " Then
        sTmpDec = Mid$(sTmpDec, Len("WithEvents ") + 1)
        pVarInfo.ObjectWithEvents = True
    End If
    
    ParseParamInfoItem sTmpDec, tPInfo
    
    With pVarInfo
                
        .Accessor = CLng(saTmp.ItemAsNumberValue(3))
        .AccessorAsString = Choose(.Accessor, "Public", "Private", "Friend")
        
        .Name = tPInfo.Name
        If tPInfo.IsArray Then
            .Name = .Name & "()" ' ... restore Array Signature
        Else
        
        End If
        .Type = tPInfo.Type
        
        .Declaration = saTmp(1) ' sTmpDec
        
        .LineStart = CLng(saTmp.ItemAsNumberValue(4))
        .EditorLineStart = CLng(saTmp.ItemAsNumberValue(5))
        
    End With
    
    sTmpDec = vbNullString
    

End Sub ' ... ParseVariableInfoItem:

Public Sub ParseParamInfoItem(ByVal pParamDec As String, pParamInfo As ParamInfo)
Attribute ParseParamInfoItem.VB_Description = "Receives a single parameter to a method/member and parses it into a ParamInfo Structure."

' ParamInfo Structure
'--------------------
'IsOptional As Boolean
'IsByRef As Boolean
'IsArray As Boolean
'IsParamArray As Boolean
'Name As String
'Type As String
'DefaultValue As String
'---------------------
' test data (note pParamInfo not in declaration but as a variable)
'Dim pParamInfo As ParamInfo
' -------------------------------------------------------------------
' ParseParamInfoItem "ByVal pParamDec As String"
' ParseParamInfoItem "ByVal pParamDec() As String"
' ParseParamInfoItem "pSArray As StringArray"
' ParseParamInfoItem "ParamArray pTheExtension() As Variant"
' ParseParamInfoItem "ParamArray pTheExtension()"
' ParseParamInfoItem "pSArray As StringArray, pIndex As Long, pMemberInfo As MemberInfo, Optional pErrMsg As String = vbNullString" ' ... crap data
' ParseParamInfoItem "Optional pErrMsg As String = vbNullString"
' ParseParamInfoItem "Optional pErrMsg As String = " & chr$(34) & "Code Broswer" & chr$(34)
' -------------------------------------------------------------------

Dim sParamDec As String
Dim sTmp As String
Dim lngAsFound As Long
Dim bDec() As Byte
Dim lngLoop As Long
Dim lngChar As Long
Dim lngPos As Long
Dim bNameTerminator As Boolean
    
    ' -------------------------------------------------------------------
    ' ... reset default values in case of loop reuse of the paraminfo structure in caller.
    
    With pParamInfo
        .DefaultValue = vbNullString
        .IsArray = False
        .IsByRef = False
        .IsOptional = False
        .IsParamArray = False
        .Name = vbNullString
        .Type = vbNullString
    End With
    
    ' -------------------------------------------------------------------
    
    If Len(pParamDec) > 0 Then
    

        sParamDec = pParamDec
        
        pParamInfo.Type = "Variant" ' ... default to variant for data type and change later as and if neccessary.
        
        ' ... Optional?
        If Left$(sParamDec, 9) = "Optional " Then
            pParamInfo.IsOptional = True
            sParamDec = Mid$(sParamDec, 10)
        ElseIf Left$(sParamDec, 11) = "ParamArray " Then
        ' ... or ParamArray?
            pParamInfo.IsParamArray = True
            sParamDec = Mid$(sParamDec, 12)
        End If
        
        ' ... ByVal or ByRef?
        pParamInfo.IsByRef = True
        
        sTmp = Left$(sParamDec, 6)
        Select Case sTmp
            Case "ByVal ", "ByRef "
                If sTmp = "ByVal " Then
                    pParamInfo.IsByRef = False
                End If
                sParamDec = Mid$(sParamDec, 7)
        End Select
        
        ' ... Name.
        bDec = sParamDec
        For lngLoop = 0 To UBound(bDec) Step 2
'            lngPos = lngPos + 1
            lngChar = bDec(lngLoop)
            Select Case lngChar
                Case 32, 40, 41 ' ... [space] , ) or (
                    bNameTerminator = True
                    Exit For    ' ... name terminator found.
            End Select
        Next lngLoop
        
        lngPos = lngLoop / 2 + 1
        
        If lngPos > 1 Then
            pParamInfo.Name = Left$(sParamDec, lngPos - 1)
            sParamDec = Mid$(sParamDec, lngPos)
        End If
        
        ' ... nothing required after name
        If Len(sParamDec) > 0 Then
                        
            If Left$(sParamDec, 2) = "()" Then
                pParamInfo.IsArray = True
                sParamDec = Mid$(sParamDec, 3)
            End If
                
            ' ... data type not required (defaulting to variant, see above)
            If Left$(sParamDec, 4) = " As " Then
                sParamDec = Mid$(sParamDec, 5)
                lngAsFound = 0
                If pParamInfo.IsOptional Then
                    lngAsFound = InStr(1, sParamDec, " = ")
                End If
                If lngAsFound = 0 Then
                    pParamInfo.Type = sParamDec
                Else
                    pParamInfo.Type = Left$(sParamDec, lngAsFound - 1)
                    pParamInfo.DefaultValue = Mid$(sParamDec, lngAsFound + Len(" = "))
                End If
            End If
                        
        End If
            
    End If

'    With pParamInfo
'        Debug.Print .Name, .Type, .IsOptional, .IsParamArray, .IsByRef, .IsArray, .DefaultValue
'    End With

End Sub ' ... ParseParamInfoItem:

Public Sub ParseVBPInfoItem(pSArray As StringArray, pIndex As Long, pDataInfo As DataInfo)
Attribute ParseVBPInfoItem.VB_Description = "Converts a string into a VBPInfoItem for processing items read from a VBPInfo instance."

Dim sTmpData As StringArray
Dim tDataInfo As DataInfo
' -------------------------------------------------------------------
' Helper:   convert stringarray data to DataInfo for adding stuff to the project explorer.
' Note:     the data in the strinarray item (index) is a delimited text string
'           representing info about an item found in the vbp.
' -------------------------------------------------------------------
    
    If Not pSArray Is Nothing Then
        If pSArray.Count > 0 Then
            If pIndex <= pSArray.Count Then
                Set sTmpData = pSArray.ItemAsStringArray(pIndex, "|")
                With tDataInfo
                    .Name = sTmpData(1)
                    .Type = CByte(sTmpData.ItemAsNumberValue(2))
                    .Exists = sTmpData(3) = "1"
                    .ExtraInfo = sTmpData(sTmpData.Count)
                End With
            End If
        End If
    End If
    ' -------------------------------------------------------------------
    pDataInfo = tDataInfo
    ' -------------------------------------------------------------------
    If Not sTmpData Is Nothing Then
        Set sTmpData = Nothing
    End If

End Sub ' ... ParseVBPInfoItem:

Public Sub ParseZipMemberInfoItem(pSArray As StringArray, _
                                  pIndex As Long, _
                                  pZipMemInfo As ZipMemberInfo, _
                         Optional pMemberDelimiter As String = "|")
Attribute ParseZipMemberInfoItem.VB_Description = "Converts a Zip Member Info String into a ZipMemberInfo item."
                         
Dim sTmpData As StringArray
Dim tZipMemInfo As ZipMemberInfo
Dim sMemberDelimiter As String
    
    If Not pSArray Is Nothing Then
        If pSArray.Count > 0 Then
            If pIndex <= pSArray.Count Then
                sMemberDelimiter = "|"
                If pMemberDelimiter <> "" Then
                    sMemberDelimiter = pMemberDelimiter
                End If
                Set sTmpData = pSArray.ItemAsStringArray(pIndex, sMemberDelimiter)
                With tZipMemInfo
                    .FileName = sTmpData(1)
                    .FilePath = sTmpData(2)
                    .FullPathAndName = sTmpData(3)
                    .FileDate = CVDate(sTmpData.ItemAsNumberValue(4))
                    .UnCompSize = CLng(sTmpData.ItemAsNumberValue(5))
                    .CompSize = CLng(sTmpData.ItemAsNumberValue(6))
                    .Index = CLng(sTmpData.ItemAsNumberValue(7))
                    .Encrypted = sTmpData.ItemAsNumberValue(8) = 1
                End With
            End If
        End If
    End If
    ' -------------------------------------------------------------------
    pZipMemInfo = tZipMemInfo
    ' -------------------------------------------------------------------
    If Not sTmpData Is Nothing Then
        Set sTmpData = Nothing
    End If
                         
                         
End Sub

Public Sub CheckHelp()
Attribute CheckHelp.VB_Description = "Returns True if Help File is found else False."

    ' ... v7, update
    ' ... look in app path for help, else look in app path \hlp
    ' ... app path \hlp originates from the project being copied.
    If Dir$(App.Path & "\SVBCB.chm") <> "" Then
        App.HelpFile = App.Path & "\SVBCB.chm"
    ElseIf Dir$(App.Path & "\hlp\SVBCB.chm") <> "" Then
        App.HelpFile = App.Path & "\hlp\SVBCB.chm"
    End If
    
End Sub

Public Sub ShowHelp()
Attribute ShowHelp.VB_Description = "Loads the Help Page ' Code Viewer.htm ' if found."
' ... no longer required as have chm now, v3/4.
' ... Helper: Load the Help document into Internet Explorer.
' ... Note:   Reliance upon hard coded value for name of help file.
Dim sFilePath As String
Dim sFilter As String

    On Error GoTo ErrHan:

'''    DeleteSetting app.EXEName,"Help","Path"
    Let sFilePath = GetSetting(App.EXEName, "Help", "Path", "")
    
    If Len(sFilePath) Then
        If Dir$(sFilePath, vbNormal) = "" Then
            Let sFilePath = vbNullString
        End If
    End If
    
    If Len(sFilePath) = 0 Then
        sFilePath = App.Path & "\Code Viewer.htm" ' ... name could be a constant.
        If Dir$(sFilePath, vbNormal) = "" Then
            sFilePath = vbNullString
        End If
    End If
    
    If Len(sFilePath) = 0 Then
        sFilter = modDialog.MakeDialogFilter("VB Code Browser Help", "Code Viewer", "htm")
        sFilePath = modDialog.GetOpenFileName("Code Viewer.htm", "Find Help Page", sFilter, , App.Path)
    End If
    
    If Len(Trim(sFilePath)) > 0 Then
        If Dir$(sFilePath) <> "" Then
            modGeneral.OpenWebPage sFilePath
            ' ... remove this to ignore writing to registry when help doc is in app path.
            Call SaveSetting(App.EXEName, "Help", "Path", sFilePath)
        End If
    Else
        Call MsgBox("Sorry, The Help File, 'Code Viewer.htm' could not be found.", vbInformation, "Help Unavailable")
    End If

Exit Sub
ErrHan:
    If Err.Number <> cDlgCancelErr Then
        Debug.Print "modGeneral.Error: " & Err.Number & "; " & Err.Description
    End If
End Sub


Public Function IsWordBreakChar(ByVal pKeyCode As Long) As Boolean
Attribute IsWordBreakChar.VB_Description = "Attempts to determine if a character represents a word break in vb."

    On Error GoTo ErrHan:

    IsWordBreakChar = pKeyCode < 48 Or pKeyCode > 57 And pKeyCode < 65 Or pKeyCode > 90 And pKeyCode < 96 Or pKeyCode > 122 And pKeyCode < 127

Exit Function

ErrHan:

    Debug.Print "modGeneral.IsWordBreakChar.Error: " & Err.Number & "; " & Err.Description

End Function

