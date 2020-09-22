VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class for Reading VB6 Code Files."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' what?
'  a class dedicated to reading code in plain text form
'  that provides access to the methods and members it contains.
' why?
'  re-use, extension...
'  to have a disposable class that can be left to parse code
'  as and when required.
' when?
'  not in IDE but want access to info in code files such as no. of members
'  member names.
' how?
'   ' dimension and create a CodeInfo object instance.
'   Dim oCInfo As CodeInfo
'       Set oCInfo = New CodeInfo
'   ' call its ReadCodeFile method with the full file name of a code file
'       oCInfo.ReadCodeFile sFileName
'   ' or just initialise it with some plain text...
'       oCInfo.ReadCodeString sCodeString
'
'   ' properties and members may then be accessed e.g.
'   ' print names of methods / members in order of appearance in code.
'
'   Dim lngCount As Long
'   Dim lngLoop As Long
'   Dim tQMI As QuickMemberInfo
'
'       lngCount = oCInfo.MemberCount               ' ... get the number of members/methods.
'       For lngLoop = 0 To lngCount - 1             ' ... loop thru' members.
'           tQMI = oCInfo.QuickMember(lngLoop)      ' ... read code member info.
'           Debug.Print tQMI.Name                   ' ... print the name of the method.
'       Next lngLoop
'
'   ' print a method's / member's lines.
'   Dim sCodeText As String
'   Dim lngMemberIndex As Long
'       lngMemberIndex = 0                          ' ... set to read first method.
'       sCodeText = oCInfo.GetMemberCodeLines(lngMemberIndex)
'       Debug.Print sCodeText
'
'   ' print a code file's name and description (if available).
'       Debug.Print oCInfo.Name & ": " & oCInfo.Description
'
'   ' release the object.
'       Set oCInfo = Nothing
'
' who?
'  d.c.

Option Explicit
' v8/9
Private mNodesCollection As Collection
Private miLastMemberIndex As Long
Private miMemHist() As Long
Private miMemHistCount As Long
Private msQuickReport As String

Private msTextRTF As String
Private msAPIs As String

Private msVars As String
Private mbReadDeclarations As Boolean
Private mbReadCNTRLStatements As Boolean
Private mlngCNTRLStatementCount As Long
Private mMenuStructure As String
Private mMenuMethods As String

Private mlngMemberDecLineCount As Long ' ... the number of lines given to member declarations.
Private mlngConstantDecLineCount As Long
Private mlngVarDecLineCount As Long
Private mlngAPIDecLineCount As Long
Private mlngTypeDecLineCount As Long
Private mlngEnumDecLineCount As Long
Private mlngEventDecLineCount As Long
' -------------------------------------------------------------------
' v6, fields to hold member info w/r find member name from char position
' ... in the text string.

Private mMemberIndexes() As Long
Private mMemberNames() As String
Private mMemberAttributes() As String

' -------------------------------------------------------------------

' ... new post release 4
' ... use modReader to read the code string and return lines array

Private mCodeText As String
Private mCodeLines() As Long
Private mTextLinesDelimter As String


' Note: Need to update this to include Header Attributes (if any).
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' ... new, post release 2.
Private mlngCommentedLineCount As Long
Private mlngEmptyLineCount As Long
Private mlngAttributeLineCount As Long
Private mFileName As String
' -------------------------------------------------------------------

' ... new, post release.
Private mType As String
Private mname As String
Private mDescription As String
' -------------------------------------------------------------------

Private maEvents As StringArray
Private maConstants As StringArray
Private maEnums As StringArray
Private maTypes As StringArray
Private maImplements As StringArray

Private maSubs As StringArray
Private maFunctions As StringArray
Private maProperties As StringArray
Private m_a As StringArray
Private m_aMethods As StringArray
Private m_aAPIs As StringArray
Private m_aImplements As StringArray

Private moCodeString As StringWorker
Private mInitialised As Boolean

Private m_HeaderStart As Long
Private m_HeaderEnd As Long

Private m_DecsStart As Long
Private m_DecsEnd As Long

Private m_MembersStart As Long
Private m_MembersEnd As Long

Private m_MemberCount As Long
Private mlngSubCount As Long
Private mlngFunctionCount As Long
Private mlngPropertyCount As Long

' -------------------------------------------------------------------

Private m_QuickMembers() As QuickMemberInfo

' -------------------------------------------------------------------

Private msEvents As String
Private tmpTypes As String      ' .. n/a mod level
Private msTypeName As String    ' .. n/a mod level
Private msConstants As String
Private msConstants9 As String
Private msTypes As String
Private msEnums As String

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "PSAPI" (ByVal hProcess As Long) As Long


 ' v5,  faffing around with the bug reading name sub, function & property
 '      applying same logic to public, private and friend.
 
Private Const cm_len_Sub As Long = 4
Private Const cm_len_Function As Long = 9
Private Const cm_len_Property As Long = 9
 
Private Const cm_word_Sub As String = "Sub "
Private Const cm_word_Function As String = "Function "
Private Const cm_word_Property As String = "Property "

Private Const cm_len_Public As Long = 7
Private Const cm_len_Private As Long = 8
Private Const cm_len_Friend As Long = 7

Private Const cm_word_Public As String = "Public "
Private Const cm_word_Private As String = "Private "
Private Const cm_word_Friend As String = "Friend "

Private mxReader As TextReader
' v8/9


Function QuickCodeReport() As String
' ... generate / read the quick report for the current code file.
Dim cReport As CodeReport
Dim stext As String
Dim sColourTable As String
    If mInitialised = False Then Exit Function
    ' -------------------------------------------------------------------
    If Len(msQuickReport) = 0 Then
        ' -------------------------------------------------------------------
        Set cReport = New CodeReport
        Set cReport.CodeReader = Me
        ' -------------------------------------------------------------------
        sColourTable = modEncode.AllBlackFontColours
        stext = cReport.ReportString
        stext = modEncode.BuildRTFString(stext, , , , sColourTable)
        msQuickReport = stext
        ' -------------------------------------------------------------------
        Set cReport = Nothing
        stext = vbNullString
        sColourTable = vbNullString
    End If
    ' -------------------------------------------------------------------
    QuickCodeReport = msQuickReport
End Function

Property Get ConstsStringv9() As String
    ConstsStringv9 = msConstants9
End Property

Property Get APIString() As String
    APIString = msAPIs
End Property

Property Get TheTextAsRTF() As String
' returns code module as rtf text, only compile rtf if not done already
    If mInitialised = False Then Exit Sub
    If Len(msTextRTF) = 0 Then
        If Len(mCodeText) > 0 Then
            msTextRTF = Me.GetDecsAndCode ' mCodeText
            msTextRTF = modEncode.BuildRTFString(msTextRTF, , , , , , True)
        End If
    End If
    TheTextAsRTF = msTextRTF
End Property


' ... v7/8

Public Property Get NonExecutingLineCount() As Long
Attribute NonExecutingLineCount.VB_Description = "No. of Lines that are either empty or commented in the source text."
    NonExecutingLineCount = mlngCommentedLineCount + mlngEmptyLineCount + mlngAttributeLineCount
End Property

Public Property Get DeclarativeLineCount() As Long
Attribute DeclarativeLineCount.VB_Description = "No. of Lines given to Method / Member Declarations."
    DeclarativeLineCount = mlngAPIDecLineCount + mlngConstantDecLineCount + mlngEnumDecLineCount + mlngMemberDecLineCount + mlngTypeDecLineCount + mlngVarDecLineCount
End Property

Public Property Get ControlStatementCount() As Long
Attribute ControlStatementCount.VB_Description = "No. of Lines given to Code Flow Control Statements."
    ControlStatementCount = mlngCNTRLStatementCount
End Property

Private Sub pGetCountOfCntrlStatements()

' ... pretty poor attempt to discover control statement lines.
' ... within the member code (ignoring header and declarations)

' ... multiple statements are completely ignored
' ... in cases not ending in a space there is much
' ... room for confusion e.g Else could indicate Else within an If block
' ... or it could just be the start of the name of a variable, method etc e.g. ElseDoIt (whatever!)

Dim i As Long
Dim sLine As String
Dim bLineIsCommented As Boolean
Dim bIncrement As Boolean
Dim bUpdateVarCount As Boolean

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... if this has already been since initialisation then don't do it again
    ' ... until the next initialisation.
    If mbReadCNTRLStatements Then Exit Sub
    
    If mInitialised = False Then
        Err.Raise vbObjectError + 1000, , "Parsing for Control Statements is only valid post general code parsing and the program thinks it has not yet done this."
    End If
    
    If m_a Is Nothing Then Exit Sub
    
    ' -------------------------------------------------------------------
    ' ... loop through the methods / members section of the source.
    For i = m_MembersStart To m_MembersEnd
        
        ' -------------------------------------------------------------------
        ' ... usual, grab line, trim it and check if it starts with a comment.
        sLine = LTrim$(m_a(i))
        bLineIsCommented = Left$(sLine, 1) = "'" Or Left$(sLine, 3) = "Rem"
        
        bIncrement = False
        
        
        If bLineIsCommented = False And Len(sLine) Then
            ' -------------------------------------------------------------------
            ' ... only bother reading non-commented, non-empty lines.
            
            If Left$(sLine, 9) = "On Error " Then
                bIncrement = True
                
            ElseIf Left$(sLine, 3) = "If " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 3) = "IIf(" Then
                bIncrement = True
                
            ElseIf Left$(sLine, 3) = "Do " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 4) = "Else" Then
                bIncrement = True
            
            ElseIf Left$(sLine, 4) = "Wend" Then
                bIncrement = True
                            
            ElseIf Left$(sLine, 4) = "For " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 5) = "Case " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 5) = "Next " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 5) = "Loop " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 5) = "Exit " Then
                bIncrement = True
                
            ElseIf Left$(sLine, 5) = "GoTo " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 5) = "With " Then
                bIncrement = True
                
            ElseIf Left$(sLine, 6) = "End If" Then
                bIncrement = True
            
            ElseIf Left$(sLine, 6) = "ElseIf" Then
                bIncrement = True
            
            ElseIf Left$(sLine, 6) = "ReDim " Then
                bIncrement = True
                
            ElseIf Left$(sLine, 6) = "Erase " Then
                bIncrement = True

            ElseIf Left$(sLine, 6) = "GoSub " Then
                bIncrement = True
                
            ElseIf Left$(sLine, 6) = "While " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 6) = "Resume" Then
                bIncrement = True
                
            ElseIf Left$(sLine, 8) = "End With" Then
                bIncrement = True
            
            ElseIf Left$(sLine, 8) = "Exit For" Then
                bIncrement = True
            
            ElseIf Left$(sLine, 9) = "For Each " Then
                bIncrement = True
            
            ElseIf Left$(sLine, 9) = "Do While " Then
                bIncrement = True
                
            ElseIf Left$(sLine, 9) = "End While" Then
                bIncrement = True
                        
            ElseIf Left$(sLine, 10) = "End Select" Then
                bIncrement = True

            ElseIf Left$(sLine, 12) = "Select Case " Then
                bIncrement = True
                
            Else
                
                bUpdateVarCount = False
                
                If Left$(sLine, 4) = "Dim " Then
                    bUpdateVarCount = True
                
                ElseIf Left$(sLine, 6) = "Const " Then
                    bUpdateVarCount = True
'                    mlngConstantDecLineCount = mlngConstantDecLineCount + 1
                    
                ElseIf Left$(sLine, 6) = "Static" Then
                    bUpdateVarCount = True
                
                End If
                
                If bUpdateVarCount Then
                    
                    mlngVarDecLineCount = mlngVarDecLineCount + 1
                
                End If
                
            End If
            
         
            If bIncrement Then
                mlngCNTRLStatementCount = mlngCNTRLStatementCount + 1
            End If
       
        End If
        
    Next i
    
    mbReadCNTRLStatements = True
    
ResumeError:

Exit Sub

ErrHan:

    Debug.Print "CodeInfo.pGetCountOfCntrlStatements.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Public Property Get EventDeclarationLineCount() As Long
Attribute EventDeclarationLineCount.VB_Description = "No. of Lines given over to Event Declarations."
    EventDeclarationLineCount = mlngEventDecLineCount
End Property

Public Property Get TypeDeclarationLineCount() As Long
Attribute TypeDeclarationLineCount.VB_Description = "No. of Lines given to Type Declarations."
    TypeDeclarationLineCount = mlngTypeDecLineCount
End Property

Public Property Get EnumDeclarationLineCount() As Long
Attribute EnumDeclarationLineCount.VB_Description = "No. of Lines given over to Enumerator Declarations."
    EnumDeclarationLineCount = mlngEnumDecLineCount
End Property

Public Property Get APIDeclarationLineCount() As Long
Attribute APIDeclarationLineCount.VB_Description = "No. of Lines used for API Declarations."
    APIDeclarationLineCount = mlngAPIDecLineCount
End Property

Public Property Get VarDeclarationLineCount() As Long
Attribute VarDeclarationLineCount.VB_Description = "No. of Lines given to Variable Declarations."
    VarDeclarationLineCount = mlngVarDecLineCount
End Property

Public Property Get ConstDeclarationLineCount() As Long
Attribute ConstDeclarationLineCount.VB_Description = "No. of Lines given to Constant Declarations."
    ConstDeclarationLineCount = mlngConstantDecLineCount
End Property

Public Property Get MemberDeclarationLineCount() As Long
Attribute MemberDeclarationLineCount.VB_Description = "Returns the number of lines given over to method / member declarations including End statements."

' ... Returns the number of lines given over to method / member declarations including End statements.

    MemberDeclarationLineCount = mlngMemberDecLineCount

End Property

Public Property Get VarsString() As String
Attribute VarsString.VB_Description = "Returns a vbCrLf delimited string of module level variables declared."
    
' ... Returns a vbCrLf delimited string of module level variables declared.

    VarsString = msVars
    
End Property
' -------------------------------------------------------------------
Public Function GetMemberFromLineNo(ByVal pLineNumber As Long, _
                           Optional ByRef pAttributes As String = vbNullString) As String
Attribute GetMemberFromLineNo.VB_Description = "Returns the Name of a Member from a Line Number in the Editor part of the code."

' ... Revision 6, updated in R7

' ... try and get the name of the member to which the
' ... line number belongs, include member attributes by way of pAttributes.

' -------------------------------------------------------------------

' ... mMemberNames and mMemberAttributes are two string arrays
' ... holding the names and attributes of the members found in the source file / string,
' ... they are generated when the source code is being parsed.
' ... mMemberIndexes is an array of longs that describe the line numbers
' ... on which successive members begin.

' -------------------------------------------------------------------


Dim x As Long
Dim y As Long
Dim a As Long
Dim b As Long

    On Error GoTo ErrHan:
    
    If mInitialised = False Then
        Err.Raise vbObjectError + 1000, , "CodeInfo not yet initialised, cannot give member name."
    End If
    
    If pLineNumber < 0 Then
        Err.Raise vbObjectError + 1000, , "Character Position specified is not valid, cannot give member name."
    End If

    If m_MemberCount < 1 Then
        Err.Raise vbObjectError + 1000, , "No Members found."
    End If
    
    ' -------------------------------------------------------------------
    ' ... clear pAttributes in case of reuse in calling method.
    pAttributes = vbNullString
    
    ' -------------------------------------------------------------------
    
    For x = 0 To m_MemberCount - 1
    
        a = mMemberIndexes(x)
        b = mMemberIndexes(x + 1)
        
        If pLineNumber >= a And pLineNumber < b Then
            pAttributes = mMemberAttributes(y)
            GetMemberFromLineNo = mMemberNames(y)
            Exit For
        ElseIf pLineNumber > b Then
            y = y + 1
        End If
    
    Next x
    
    ' -------------------------------------------------------------------
    ' ... if not caught in loop then must be last member.
    If Len(GetMemberFromLineNo) = 0 Then
        pAttributes = mMemberAttributes(m_MemberCount)
        GetMemberFromLineNo = mMemberNames(m_MemberCount)
    End If
    ' -------------------------------------------------------------------
    
ResumeError:
    a = 0&
    b = 0&
    x = 0&
    y = 0&
Exit Function

ErrHan:

    Debug.Print "CodeInfo.GetMemberFromLineNo.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

' ... Revision 3
Public Function Find(ByRef pFind As String, Optional ByRef pStart As Long = 1, Optional ByRef pCompareMethod As VbCompareMethod = vbBinaryCompare, Optional pWholeWordOnly As Boolean = False, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute Find.VB_Description = "Search the source text for a substring and return position or -1."
' ... trying to help extend the use of Project Search function.
    Find = -1
    
    If Not moCodeString Is Nothing Then
        Find = moCodeString.Find(pFind, pStart, pCompareMethod, pWholeWordOnly, pOK, pErrMsg)
    End If

End Function

Public Function FindAllMatches(ByVal pFind As String, pPositionArray() As Long, _
                      Optional ByVal pStart As Long = 0, _
                      Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
                      Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute FindAllMatches.VB_Description = "Returns the number of matches for a substring in the source string or -1 for fail."

' ... Returns the number of matches for a substring in the source string or -1 for fail.

    FindAllMatches = -1
    
    If Not moCodeString Is Nothing Then
        
        FindAllMatches = moCodeString.FindAllMatches(pFind, pPositionArray, pStart, pCompare, pErrMsg)
                
    End If
    
End Function


' ... Revision 4
Public Function GetAllMatchesEx(ByVal pFind As String, _
                       Optional ByVal pStart As Long = 0, _
                       Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
                       Optional ByVal pWholeWordOnly As Boolean = False, _
                       Optional ByRef pCount As Long = 0, _
                       Optional ByRef pUniqueLineCount As Long = 0, _
                       Optional ByRef pOK As Boolean = False, _
                       Optional ByRef pErrMsg As String = vbNullString) As String
    
    GetAllMatchesEx = mxReader.GetAllMatches(pFind, pStart, pCompare, pWholeWordOnly, pCount, pUniqueLineCount, pErrMsg)

End Function


' ... Revision 4
Public Function GetAllMatches(ByVal pFind As String, _
                     Optional ByVal pStart As Long = 0, _
                     Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
                     Optional ByVal pWholeWordOnly As Boolean = False, _
                     Optional ByRef pCount As Long = 0, _
                     Optional ByRef pOK As Boolean = False, _
                     Optional ByRef pErrMsg As String = vbNullString) As StringArray
Attribute GetAllMatches.VB_Description = "Return a StringArray of search results."

' ... Return a StringArray of search results.

' -------------------------------------------------------------------
' ... with a bit of luck a client can call this and get a stringarray
' ... of results, each line may then be parsed into a findresultinfo structure
' ... for easier processing and reading.
' ... perhaps a little over the top with the file name added each time.
' -------------------------------------------------------------------

Dim lngFoundPositions() As Long
Dim lngFileLines() As Long
Dim lngLines() As Long
Dim lngLens() As Long

Dim c As Long
Dim k As Long
Dim p As Long
Dim lIndex As Long
Dim lngLength As Long
Dim lngDelLen As Long           ' ... length of delimiter.
Dim lngLastLine As Long
Dim lngLineCount As Long
Dim x As Long
Dim y As Long
Dim a As Long
Dim b As Long

Dim sTmpLine As String
Dim sTmpMemberName As String
Dim sTmpAdd As String

Dim sFieldDel As String           ' ... output field delimiter.

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... a bit of validation.
    
    If moCodeString Is Nothing Then
        Err.Raise vbObjectError + 1000, , "No Source Code"
    End If
    
    If Len(pFind) < 1 Then
        Err.Raise vbObjectError + 1000, , "Search String length may not be zero"
    End If
    
    If pStart + Len(pFind) >= moCodeString.Length Then
        Err.Raise vbObjectError + 1000, , "Search Start Position > the source string"
    End If
    
    lngDelLen = Len(mTextLinesDelimter) ' ... needed for processing with Derive[...] methods below.
    
    If lngDelLen = 0 Then
        Err.Raise vbObjectError + 1000, , "No known Line Delimiter in search source"
    End If
    
    If pStart < 1 Then pStart = 1
    
    ' -------------------------------------------------------------------
    ' ... set up default return values.
    pErrMsg = vbNullString
    pOK = False
    ' -------------------------------------------------------------------
    
    sFieldDel = Chr$(0)
    
    ' -------------------------------------------------------------------
    ' ... initial find.
    pCount = moCodeString.FindAllMatches(pFind, lngFoundPositions, pStart, pCompare, pWholeWordOnly, pErrMsg)
    
    If Len(pErrMsg) > 0 Then
        Err.Raise vbObjectError + 1000, , pErrMsg
    End If
    
    ' -------------------------------------------------------------------
    Set GetAllMatches = New StringArray
    GetAllMatches.Sortable = False
    ' -------------------------------------------------------------------
    ' ... new properties for StringArray.
    GetAllMatches.Tag = mFileName   ' ... file name.
    GetAllMatches.Name = mname      ' ... module/class name.
    ' -------------------------------------------------------------------
    ' ... process  results for return, if any.
    ' -------------------------------------------------------------------
    If pCount > 0 Then
        
        ' ... .
        modReader.DeriveLinesArray mCodeText, mCodeLines, mTextLinesDelimter
        ' ... derive line numbers from search results; using Line Delimiter Length, File Lines and Char Pos Arrays, derive line nos. for substring finds.
        modReader.DeriveLineNumbers mCodeLines, lngFoundPositions, lngLines, lngDelLen
        ' ... derive line lengths from search results and line numbers.
        modReader.DeriveLineLengths mCodeLines, lngLines, lngLens, lngDelLen
    
        
        For k = 0 To UBound(lngLines)
            
            ' ... NOTE:
            ' ... at present, line numbering begins in the source file's header section not the start of its declarative section.
            ' ... extract lines where the substring is found.
            c = c + 1
            
            lIndex = lngLines(k)
'            lIndex = lIndex - GetHeaderLineCount
            
            If lIndex < 1 And k > 0 Then
                lIndex = UBound(mCodeLines) - 1
            Else
                p = lngFoundPositions(k) - mCodeLines(lIndex)
            End If
            
            If p < 1 Then p = 1
            
            If lIndex <> lngLastLine Or lngLastLine = 0 Then
            
                lngLineCount = lngLineCount + 1
                
                lngLastLine = lIndex
                lngLength = lngLens(k)
                
                ' -------------------------------------------------------------------
                ' ... extract current line from file text.
                ' ... there's a bug calculating line lengths (or something)
                ' ... when substring is the first thing in the file e.g. VERSION
                ' ... we read an extra character in the line.
                
                sTmpLine = Mid$(mCodeText, mCodeLines(lIndex), lngLength)
                
                If p = 1 Then
                    ' ... hack bug.
                    If Right$(sTmpLine, 1) = vbCr Then
                        sTmpLine = Left$(sTmpLine, Len(sTmpLine) - 1)
                    End If
                End If
                ' -------------------------------------------------------------------
                
                ' -------------------------------------------------------------------
                ' ... fall back on module level arrays of member names and member indexes.
'                sTmpMemberName = mMemberNames(UBound(mMemberNames)) ' ... default to last because loop never gets there.
                lIndex = lIndex - GetHeaderLineCount
                sTmpMemberName = Me.GetMemberFromLineNo(lIndex) ' - Me.GetHeaderLineCount)
                ' -------------------------------------------------------------------
'                ' ... try finding the name of the method or member.
'                For x = y To UBound(mMemberIndexes) - 1
'
'                    a = mMemberIndexes(x)
'                    b = mMemberIndexes(x + 1) - 1
'
'                    If lIndex >= a And lIndex <= b Then
'                        sTmpMemberName = mMemberNames(y)
'                        Exit For
'                    ElseIf lIndex > b Then
'                        y = y + 1
'                    End If
'
'                Next x
                
                If Len(sTmpLine) Then
                    
                    ' -------------------------------------------------------------------
                    ' ... build output line
                    ' ... file name, module name, member name, line containing, index, line number, line pos., text pos.
                    ' ... using Chr$(0) as field delimiter for now.
                    ' ... have added Name and Tag to string array so could try using these
                    ' ... for the file name and the module name instead of repeating them for each
                    ' ... line of data.
                    sTmpAdd = mFileName & sFieldDel & mname & sFieldDel & sTmpMemberName & sFieldDel & sTmpLine & sFieldDel
                    sTmpAdd = sTmpAdd & c & sFieldDel & lIndex & sFieldDel & p & sFieldDel & lngFoundPositions(k)
                    ' ... add output line to stringarray
                    GetAllMatches.AddItemString sTmpAdd
                    
                    ' ... release resources.
                    sTmpAdd = vbNullString
                    sTmpLine = vbNullString
                    sTmpMemberName = vbNullString
                    
                End If
                
            End If
                                                    
        Next k
    
    End If
    
    pOK = True
    
ResumeError:

    ' -------------------------------------------------------------------
    ' ... clean up.
    Erase lngFileLines
    Erase lngFoundPositions
    Erase lngLines
    Erase lngLens
    
Exit Function

ErrHan:
    
    pOK = False
    pErrMsg = Err.Description
    Debug.Print "CodeInfo.GetAllMatches.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

' -------------------------------------------------------------------
' ... Revision 2.
Public Property Get ActualLineCount() As Long
Attribute ActualLineCount.VB_Description = "Returns no. of text lines proceeding header that are not commented."
Dim lngLineCount As Long
    
' ... Note: currently ignoring header lines in calculation.

    lngLineCount = GetLineCount
        
    If lngLineCount > 0 Then ' mlngMemberDecLineCount
'        If mbReadDeclarations = False Then
'            pGetCountOfCntrlStatements
'        End If
        ' ... lngLineCount = lngLineCount - mlngCommentedLineCount - mlngEmptyLineCount ' ... v7/8 update as next line shows.
        lngLineCount = lngLineCount - mlngCommentedLineCount - mlngEmptyLineCount - mlngMemberDecLineCount - mlngConstantDecLineCount - mlngVarDecLineCount - mlngAPIDecLineCount - mlngEventDecLineCount - mlngTypeDecLineCount - mlngEnumDecLineCount - mlngCNTRLStatementCount + mlngAttributeLineCount
        If lngLineCount < 0 Then lngLineCount = 0
    End If
    
    ActualLineCount = lngLineCount
    
End Property

Public Property Get EmptyLineCount() As Long
Attribute EmptyLineCount.VB_Description = "Returns the no. of empty lines post header."
    EmptyLineCount = mlngEmptyLineCount
End Property

Public Property Get CountCommentLines() As Long
Attribute CountCommentLines.VB_Description = "Returns the number of lines beiginning with a comment e.g. ' or Rem."
    CountCommentLines = mlngCommentedLineCount
End Property

Public Property Get CountAttributeLines() As Long
Attribute CountAttributeLines.VB_Description = "No. of Lines given to hidden Attributes in Source Member Code."
    CountAttributeLines = mlngAttributeLineCount
End Property

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a Boolean describing whether the class has been initialised and therefore useable as is."
    Initialised = mInitialised
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "The name of the source file currently loaded, if any."
    FileName = mFileName
End Property

' -------------------------------------------------------------------
' ... Revision 1.
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name of the code file / class as written in its header (if any)."
    Name = mname
End Property

Public Property Get Description() As String
Attribute Description.VB_Description = "Returns the Description of the code file / class as found in its header."
    Description = mDescription
End Property

Public Property Get MemberType() As String
Attribute MemberType.VB_Description = "Returns a string describing the type of code file e.g. Form, Modue, Class & User Control."
    MemberType = mType
End Property
' -------------------------------------------------------------------

Private Sub ClearMemory()
' -------------------------------------------------------------------
' Helper:   Clears up application memory use.
' Note:     Trimmed down version, minimal functionality.
' -------------------------------------------------------------------
Dim lngProcHandle As Long
Dim lngRet As Long
    On Error GoTo LogMemory_Err
    lngProcHandle = GetCurrentProcess()
    lngRet = EmptyWorkingSet(lngProcHandle)
LogMemory_Exit:
Exit Sub
LogMemory_Err:
End Sub

Public Function StructuresStringArray() As StringArray
Attribute StructuresStringArray.VB_Description = "Attempts to return the Types / Structures found in the code file's Declarations section as a StringArray."
    If Not maTypes Is Nothing Then
        Set StructuresStringArray = maTypes
    Else
        Set StructuresStringArray = New StringArray
    End If
End Function

Public Function EnumsStringArray() As StringArray
Attribute EnumsStringArray.VB_Description = "Attempts to return the Enumerators found in the code file's Declarations section as a StringArray."
    If Not maEnums Is Nothing Then
        Set EnumsStringArray = maEnums
    Else
        Set EnumsStringArray = New StringArray
    End If
End Function

Public Function EventsStringArray() As StringArray
Attribute EventsStringArray.VB_Description = "Attempts to return the Events found in the code file's Declarations section as a StringArray."
    If Not maEvents Is Nothing Then
        Set EventsStringArray = maEvents
    Else
        Set EventsStringArray = New StringArray
    End If
End Function

Public Function ConstantsStringArray() As StringArray
Attribute ConstantsStringArray.VB_Description = "Attempts to return the Constants found in the code file's Declarations section as a StringArray."
    If Not maConstants Is Nothing Then
        Set ConstantsStringArray = maConstants
    Else
        Set ConstantsStringArray = New StringArray
    End If
End Function

Public Function APIStringArray() As StringArray
Attribute APIStringArray.VB_Description = "Attempts to return the APIs found in the code file's Declarations section as a StringArray."
    If Not m_aAPIs Is Nothing Then
        Set APIStringArray = m_aAPIs
    Else
        Set APIStringArray = New StringArray
    End If
End Function

Private Sub Class_Terminate()
    
    pRelease

End Sub

Public Property Get ConstsString() As String
Attribute ConstsString.VB_Description = "Returns a vbCrLF delimited String describing Constants found."
    ConstsString = msConstants
End Property

Public Property Get CountOfFunctions() As Long
Attribute CountOfFunctions.VB_Description = "Returns a count of the Functions found."
    CountOfFunctions = mlngFunctionCount
End Property

Public Property Get CountOfProperties() As Long
Attribute CountOfProperties.VB_Description = "Returns a count of the Properties found (Get and Let/Set are counted separately)."
    CountOfProperties = mlngPropertyCount
End Property

Public Property Get CountOfSubs() As Long
Attribute CountOfSubs.VB_Description = "Returns a count of the Subs found."
    CountOfSubs = mlngSubCount
End Property

Public Function Declarations() As String
Attribute Declarations.VB_Description = "Attempts to parse the Declarations section of the code for Events, Constants, Types, Enums... This needs to have been called by the client prior to reading Declaration related data."

' ... This method attempts to read the declarative section of the code file
' ... and record Events, Constants, Types and Enums (hoping to add API and Implements).
' ... Each item found is added to a group string on a new line.
' ... Types and Enumerators along with their members are written as a single line
' ... with 2 delimiters the first (:) provides the name to its left and is written once
' ... and the second (;) is used to identify each member e.g. MyType:First As String;Second As Long;Third As Date etc.
' ...                                                    &    MyEnum:First;Second;Third etc.
' ... End Type / Enum are not included.

' ... The intention, hope, is that the grouped string data may then easily be read, sorted, presented etc.

' v7
' ... Notes:    Hoping to expand the information available to include
' ...           line start and editor line start as with API.
' ...           Also, have finally realised I'm not capturing variables
' ...           and Objects (is this case, I mean Object to be anything dimmed WithEvents).

Dim i As Long
Dim j As Long ' v6, extended line loop.
Dim sLine As String
Dim bLineIsCommented As Boolean
Dim lngAccessor As Long
Dim sTmp As String
Dim sEnd As String
Dim bInType As Boolean
Dim bInEnum As Boolean
Dim sTmpEnums As String
Dim sLeft As String
Dim sRight As String
Dim lngAsFound As Long

' -------------------------------------------------------------------
' ... v7, ... line number and editor line number.
Dim lngLineNo As Long
Dim lngEditorLineNo As Long
Dim bRemovedAccessor As Boolean
Dim strAccessor As String

' -------------------------------------------------------------------
' Note: may only be called following initial parsing in pParseClass, else nothing to go on.
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... v7, avoid appending to existing data; a call once method.
    If mbReadDeclarations = True Then Exit Function
    
    If mInitialised = False Then
        Err.Raise vbObjectError + 1000, , "Parsing Declarations section is only valid post general code parsing and the program thinks it has not yet done this."
    End If
    
    If Not m_a Is Nothing Then
    
        If m_DecsStart < 0 And m_DecsEnd < 0 Or m_DecsStart > m_DecsEnd Then
            Exit Function
        
        End If
        
        Set m_aImplements = New StringArray
        
        ' -------------------------------------------------------------------
        ' ... loop through the lines of the declaration section of the source.
        For i = m_DecsStart To m_DecsEnd
        
            sLine = LTrim$(m_a(i))                      ' ... capture and trim the line.
            bLineIsCommented = Left$(sLine, 1) = "'" Or Left$(sLine, 3) = "Rem"     ' ... check if line is already commented.
            
            If bLineIsCommented = False And Len(sLine) Then
                
                lngAccessor = 1                             ' ... default to public accessor.
                bRemovedAccessor = False
                strAccessor = vbNullString
                
                If Left$(sLine, cm_len_Public) = cm_word_Public Then
                    sLine = Mid$(sLine, c_len_Public + 2)
                    bRemovedAccessor = True
                    strAccessor = cm_word_Public
                ElseIf Left$(sLine, cm_len_Private) = cm_word_Private Then
                    sLine = Mid$(sLine, c_len_Private + 2)
                    lngAccessor = 2
                    bRemovedAccessor = True
                    strAccessor = cm_word_Private
                End If
                
                ' ... trim line behind comment for safer line extension testing.
                sTmp = modStrings.LeftOfComment(sLine)
                ' -------------------------------------------------------------------

                If Left$(sLine, 6) = "Event " Then
                    
                    sTmp = Trim$(sTmp)
                    
                    If Len(sTmp) Then
                        mlngEventDecLineCount = mlngEventDecLineCount + 1
                        ' -------------------------------------------------------------------
                        ' ... v7, rework.
                        ' ... check for line extension.
                        If Right$(sTmp, 2) = " _" Then
                            If i < m_DecsEnd Then
                                For j = i + 1 To m_DecsEnd
                                    
                                    sLine = LeftOfComment(m_a(j))
                                    sTmp = Left$(sTmp, Len(sTmp) - 1)
                                    
                                    If Right$(sTmp, 3) = " & " Then
                                        ' ... string concatenation.
                                        sTmp = sTmp & Mid$(sLine, InStr(1, sLine, Chr$(34)))
                                        ' ... eg "test " & _ (newline) & "data" will output >> "test " & "data"
                                    Else
                                        ' ... line extension.
                                        sTmp = sTmp & LTrim$(sLine)
                                    End If

                                    If Right$(sTmp, 1) <> "_" Then
                                        i = j
                                        Exit For
                                    Else
                                        mlngEventDecLineCount = mlngEventDecLineCount + 1
                                    End If
                                    
                                Next j
                                
                            End If
                        End If
                        
                        lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart

                        
                        If Len(msEvents) Then msEvents = msEvents & vbCrLf
                        msEvents = msEvents & sTmp
                        
                        maEvents.AddItemString sTmp
                    
                    End If
                    
                ElseIf Left$(sLine, 11) = "Implements " Then
                    ' -------------------------------------------------------------------
                    ' ... v7 note: Implements is a bit flaky here.
                    sTmp = modStrings.Replace(sTmp, "Implements", vbNullString)
                    sTmp = Trim$(sTmp)
                    If Len(sTmp) Then
                        lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart
                        m_aImplements.AddItemString sTmp
                    End If
                    
                ElseIf Left$(sLine, 6) = "Const " Then
                    
                    sTmp = Trim$(sTmp)
                    
                    If Len(sTmp) Then
                    
                        lngAsFound = InStr(1, sTmp, " As ")
                        
                        If lngAsFound > 0 Then
                            ' ... handle the case where there are many spaces between the name and the As [Type] bit.
                            sRight = Mid$(sTmp, lngAsFound)
                            sLeft = RTrim$(Left$(sTmp, lngAsFound - 1))
                            sTmp = sLeft & sRight
                        End If
                    
                        ' ... v7, rework.
                        ' ... check for line extension.
                        ' -------------------------------------------------------------------
                        mlngConstantDecLineCount = mlngConstantDecLineCount + 1
                        ' -------------------------------------------------------------------
                        If Right$(sTmp, 2) = " _" Then
                        
                            If i < m_DecsEnd Then
                            
                                For j = i + 1 To m_DecsEnd
                                    
                                    sLine = LeftOfComment(m_a(j))
                                    sTmp = Left$(sTmp, Len(sTmp) - 1)
                                    
                                    If Right$(sTmp, 3) = " & " Then
                                        ' ... string concatenation.
                                        sTmp = sTmp & Mid$(sLine, InStr(1, sLine, Chr$(34)))
                                        ' ... eg "test " & _ (newline) & "data" will output >> "test " & "data"
                                    Else
                                        ' ... line extension.
                                        sTmp = sTmp & LTrim$(sLine)
                                    End If
                                    
                                    If Right$(sTmp, 1) <> "_" Then
                                        i = j
                                        Exit For
                                    Else
                                        mlngConstantDecLineCount = mlngConstantDecLineCount + 1
                                    End If
                                    
                                Next j
                                
                            End If
                            
                        End If
                        
                        lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart
                        
                        sTmp = Choose(lngAccessor, c_word_Public, c_word_Private, c_word_Friend) & Space$(1) & sTmp
                        
                        If Len(msConstants) Then msConstants = msConstants & vbCrLf
                        msConstants = msConstants & sTmp
                        
                        If Len(msConstants9) > 0 Then msConstants9 = msConstants9 & vbCrLf
                        msConstants9 = msConstants9 & sTmp & "|" & lngAccessor & "|" & lngLineNo & "|" & lngEditorLineNo
                        
                        maConstants.AddItemString sTmp
                        
                    End If
                    
'                ElseIf Left$(sLine, 11) = "WithEvents " Then
'                    Debug.Print "WithEvents", sLine
                    
                ElseIf Left$(sLine, 5) = "Type " Or bInType Then
                    
                    If bInType = False Then
                        sEnd = "End Type"
                        bInType = True
                        ' ... get name of type, next word.
                        msTypeName = Mid$(sTmp, 6)
                        tmpTypes = vbNullString
                    Else
                        ' -------------------------------------------------------------------
                        ' ... v7/8
                        mlngTypeDecLineCount = mlngTypeDecLineCount + 1
                        ' -------------------------------------------------------------------
                        If Left$(sTmp, Len(sEnd)) = sEnd Then
                            
                            lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart
                            
                            If Len(msTypes) Then msTypes = msTypes & vbCrLf
                            msTypes = msTypes & msTypeName & ":" & tmpTypes
                            maTypes.AddItemString msTypeName & ":" & tmpTypes
                            bInType = False
                        Else
                            If Len(Trim$(sTmp)) Then
                                lngAsFound = InStr(1, sTmp, " As ")
                                If lngAsFound > 0 Then
                                    ' ... handle the case where there are many spaces between the name and the As [Type] bit.
                                    sRight = Mid$(sTmp, lngAsFound)
                                    sLeft = RTrim$(Left$(sTmp, lngAsFound - 1))
                                    sTmp = sLeft & sRight
                                End If
                                If Len(tmpTypes) Then tmpTypes = tmpTypes & ";"
                                tmpTypes = tmpTypes & sTmp
                            End If
                        End If
                    End If
                    
                ElseIf Left$(sLine, 5) = "Enum " Or bInEnum Then
                    
                    If bInEnum = False Then
                        sEnd = "End Enum"
                        bInEnum = True
                        msTypeName = Mid$(sTmp, 6)
                        sTmpEnums = vbNullString
                    Else
                        ' -------------------------------------------------------------------
                        mlngEnumDecLineCount = mlngEnumDecLineCount + 1
                        ' -------------------------------------------------------------------
                        If Left$(sTmp, Len(sEnd)) = sEnd Then
                            
                            lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart
                            
                            If Len(msEnums) Then msEnums = msEnums & vbCrLf
                            msEnums = msEnums & msTypeName & ":" & sTmpEnums
                            maEnums.AddItemString msTypeName & ":" & sTmpEnums
                            bInEnum = False
                        Else
                            If Len(Trim$(sTmp)) Then
                                If Len(sTmpEnums) Then sTmpEnums = sTmpEnums & ";"
                                sTmpEnums = sTmpEnums & sTmp
                            End If
                        End If
                    End If
                
                Else
                    ' -------------------------------------------------------------------
                    ' ... Module level vars and objects (in progress)
                    If bRemovedAccessor Then
                        ' ... looks like we have a variable or an api, ignore apis, only want vars/objects.
                        If Left$(sLine, 8) <> "Declare " Then
                            If Left$(sLine, 11) = "WithEvents " Then
'                                Debug.Print "WithEvents", sLine
                                sLine = Mid$(sLine, 12)
                                strAccessor = strAccessor & "WithEvents "
                            End If
                            mlngVarDecLineCount = mlngVarDecLineCount + 1
                            lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart
                            sLine = modStrings.LeftOfComment(sLine)
                            sLine = strAccessor & sLine & "|0|" & CStr(lngAccessor) & "|" & CStr(lngLineNo) & "|" & CStr(lngEditorLineNo)
                            If Len(msVars) Then sLine = vbCrLf & sLine
                            msVars = msVars & sLine
                        End If
                        
                    Else
                        If Left$(sLine, 4) = "Dim " Then
                            mlngVarDecLineCount = mlngVarDecLineCount + 1
                            lngLineNo = i + 1: lngEditorLineNo = lngLineNo - m_DecsStart
                            sLine = modStrings.LeftOfComment(sLine)
                            sLine = Mid$(sLine, 5)
                            sLine = sLine & "|0|" & CStr(lngAccessor) & "|" & CStr(lngLineNo) & "|" & CStr(lngEditorLineNo)
                            If Len(msVars) Then sLine = vbCrLf & sLine
                            msVars = msVars & sLine
                        End If
'                    ElseIf Left$(sLine, 11) <> "WithEvents " Then
'                        Debug.Print "WithEvents", sLine
                        
'                        Debug.Print sLine
                    End If
                    
                    
                    
                End If
                
            End If
            
        Next i
            
    Else
        Err.Raise vbObjectError + 1000, , "There is no raw data on which to work."
    End If
    
    If Not m_aImplements Is Nothing Then
        If m_aImplements.Count > 0 Then
            m_aImplements.Sort
        End If
    End If
    
    If Not maTypes Is Nothing Then
        If maTypes.Count > 0 Then maTypes.Sort
    End If
    
    If Not maEnums Is Nothing Then
        If maEnums.Count > 0 Then maEnums.Sort
    End If
    
    pParseForAPIs
    ' -------------------------------------------------------------------
    ' ... v7/8, line count update.
    pGetCountOfCntrlStatements
    
    ' -------------------------------------------------------------------
    ' ... v7/8 Menu, only required for forms and user controls.
    ' ... another overhead to reading the declarations section.
    mMenuStructure = ParseVBMenu(GetHeader, mTextLinesDelimter, mMenuMethods)
    
    mbReadDeclarations = True

ResumeError:
    ClearMemory
Exit Function

ErrHan:

    Debug.Print "Code.pParseDeclarations.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function ' ... Declarations: String

Public Property Get MenuStructure() As String
Attribute MenuStructure.VB_Description = "Returns a vbCrLf Line Delimited string where each line describes various attributes of each menu item encountered."
    MenuStructure = mMenuStructure
End Property

Public Property Get MenuMethods() As String
Attribute MenuMethods.VB_Description = "Returns a string of the menu click event names so we can test if a procedure name belongs to a menu or not."
    MenuMethods = mMenuMethods
End Property

Public Property Get HasMenu() As Boolean
Attribute HasMenu.VB_Description = "Returns a Boolean describing whether the source code has a VB Menu written in its Header."
    HasMenu = CBool(Len(mMenuStructure) > 0)
End Property

Public Property Get EnumsString() As String
Attribute EnumsString.VB_Description = "A vbCrLf delimited string detailing enumerators and their members."
    EnumsString = msEnums
End Property

Public Property Get EventsString() As String
Attribute EventsString.VB_Description = "Returns a vbCrLf delimited string describing the Events found in the Class along with their parameters."
    EventsString = msEvents
End Property

Public Function ExtractParamsString(ByRef pDeclaration As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As String
Attribute ExtractParamsString.VB_Description = "Returns the Parameters part of a method's / member's declaration without the parenthesis."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' ... Returns the Parameters part of a Declaration String without the parenthesis.

' ... NOTE: Needs a re-think, could end up with bolox.

Dim lngFind As Long
Dim ltmp As Long
Dim sParamString As String

    On Error GoTo ErrHan:
    
    If Len(pDeclaration) Then
        ' -------------------------------------------------------------------
        ' ... get the parameters string for digest.
        ' ... find first open bracket, find last close bracket that is not
        ' ... included in an array signature, rip the text inbetween.
        lngFind = InStr(1, pDeclaration, "(")
        If lngFind > 0 Then
            For ltmp = Len(pDeclaration) To 1 Step -1
                If Mid$(pDeclaration, ltmp, 1) = ")" Then
                    If Mid$(pDeclaration, ltmp - 1, 1) <> "(" Then  ' ... e.g. not ().
                        sParamString = Mid$(pDeclaration, lngFind + 1, ltmp - lngFind - 1)
                        Exit For
                    End If
                End If
            Next ltmp
        End If
    Else
        Err.Raise vbObjectError + 1001, , "No Parameters to return because the Declaration String was empty."
    End If

    Let pOK = True

    Let ExtractParamsString = sParamString
    
    
    ' ... v8, new method
    ' ... step through the chars
Dim bBytes() As Byte
Dim i As Long
Dim iChar As Long
Dim iOpen As Long
Dim iClosed As Long
Dim bInQuotes As Boolean
Dim iAdd As Long
Dim iBegin As Long
Dim iEnd As Long

    If Len(pDeclaration) Then
        bBytes = pDeclaration
        For i = 0 To UBound(bBytes) Step 2
            iChar = bBytes(i)
            iAdd = 0
            Select Case iChar
                Case 34: bInQuotes = Not bInQuotes
                Case 40, 41
                    If bInQuotes = False Then
                        iAdd = 1
                        If iChar = 40 Then
                            iAdd = -1
                        End If
                    End If
            End Select
            If iBegin = 0 And iAdd = 1 Then iBegin = i
            If iEnd = 0 And iAdd = 0 Then iEnd = i
        Next i
    End If

Exit Function
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "Code.ExtractParamsString", Err.Number, Err.Description

End Function ' ... ExtractParamsString: String.


Public Sub GetDecParams(pParamString As String, pParamsArray() As String, Optional pCount As Long = 0)
Attribute GetDecParams.VB_Description = "Returns a string array (via pParamsArray) of the parameters found in the input parameters string."

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
    
End Sub ' ... GetDecParams:

Public Function GetDeclarations() As String
Attribute GetDeclarations.VB_Description = "Attempts to return the Declarations section of the code."

' ... Attempts to return the Declarations section of the code.

Dim sResult As String

    On Error GoTo ErrHan:
    
    If m_DecsStart < 0 And m_DecsEnd < 0 Or m_DecsStart > m_DecsEnd Then
        Exit Function
    
    End If
        
    sResult = pGetTextLines(m_DecsStart, m_DecsEnd)
    
    GetDeclarations = sResult

ResumeError:
    
    sResult = vbNullString
    
Exit Function

ErrHan:

    Debug.Print "Code.GetDeclarations.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function ' ... GetDeclarations: String

Public Function GetDeclarationsLineCount() As Long
Attribute GetDeclarationsLineCount.VB_Description = "Returns the number of lines between the end of the header and the first method's declaration."
    GetDeclarationsLineCount = m_DecsEnd - m_DecsStart
End Function

Public Function GetDecsAndCode() As String
Attribute GetDecsAndCode.VB_Description = "Returns the entire code and declarations from the code file / class e.g. everything but the header (if any)."

' ... Returns the entire code and declarations from the code file / class e.g. everything but the header (if any).

Dim sResult As String
Dim lngLoop As Long
Dim lngStart As Long

    On Error GoTo ErrHan:
        
    If m_DecsStart > 0 And m_MembersEnd > m_DecsStart Then
    
        ' ... v7,   try just extracting the full text with mid$
        ' ...       and forgo the loop in pGetTextLines.
        
        lngStart = mCodeLines(m_DecsStart - 1)
        
        If lngStart Then
            
            ' ... adding vbCrLf for backward compatibility
            ' ... encoding needs updating, if last word is
            ' ... keyword and no new line after then
            ' ... keyword is coloured whether in comment or not.
            
            sResult = Mid$(mCodeText, lngStart) & vbCrLf
        Else
            sResult = pGetTextLines(m_DecsStart, m_MembersEnd)
        End If
    
    End If
    
    GetDecsAndCode = sResult
    
ResumeError:
    
    sResult = vbNullString
    
Exit Function
ErrHan:
    
    Debug.Print "Code.GetDeclarations.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function ' ... GetDecsAndCode: String

Public Function GetHeader() As String
Attribute GetHeader.VB_Description = "Attempts to return the Header section of the code e.g. the hidden properties and any Designer related stuff."

' ... Attempts to return the Header section of the code e.g. the hidden properties and any Designer related stuff.

Dim sResult As String

    On Error GoTo ErrHan:
    
    If m_HeaderStart < 0 And m_HeaderEnd < 0 Or m_HeaderStart > m_HeaderEnd Then
        Exit Function
    
    End If
    
    sResult = pGetTextLines(m_HeaderStart, m_HeaderEnd)
    
    GetHeader = sResult

ResumeError:
    
    sResult = vbNullString
    
Exit Function

ErrHan:

    Debug.Print "Code.GetHeader.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function ' ... GetHeader: String

Public Function GetHeaderLineCount() As Long
' v8
    If m_HeaderStart < 0 And m_HeaderEnd < 0 Or m_HeaderStart > m_HeaderEnd Then Exit Function
    GetHeaderLineCount = m_HeaderEnd
End Function

Public Function GetLineCount() As Long
Attribute GetLineCount.VB_Description = "Attempts to return a Count of Lines for the Code section."

' ... Attempts to return a Count of Lines for the Code section.

'    If m_MembersEnd > m_MembersStart Then
'        ' -------------------------------------------------------------------
'        ' ... rough indicator that we have code.
'        GetLineCount = m_MembersEnd - m_DecsStart
'    End If

' ... Attempts to return a Count of Lines for the Code and Declarative sections.
' ... v8, just amending the description and initial test.

    If m_MembersEnd > m_DecsStart Then
        ' -------------------------------------------------------------------
        ' ... rough indicator that we have code.
        GetLineCount = m_MembersEnd - m_DecsStart
    End If

End Function ' ... GetLineCount: Long

Function MemberHistoryCurrent() As Long
    MemberHistoryCurrent = -1
    If miMemHistCount > 0 Then
        If miLastMemberIndex <= miMemHistCount Then
            MemberHistoryCurrent = miMemHist(miLastMemberIndex)
        End If
    End If
End Function

Function MemberHistoryPrevious() As Long
    MemberHistoryPrevious = -1
    If miMemHistCount > 0 Then
        If miLastMemberIndex - 1 < 0 Then
            miLastMemberIndex = miMemHistCount - 1
        Else
            miLastMemberIndex = miLastMemberIndex - 1
        End If
        MemberHistoryPrevious = miMemHist(miLastMemberIndex)
    End If
End Function

Function MemberHistoryNext() As Long
    MemberHistoryNext = -1
    If miMemHistCount > 0 Then
        If miLastMemberIndex + 1 > miMemHistCount - 1 Then
            miLastMemberIndex = 0
        Else
            miLastMemberIndex = miLastMemberIndex + 1
        End If
        MemberHistoryNext = miMemHist(miLastMemberIndex)
    End If
End Function

Private Sub pAddMemberHistory(pMemberIndex As Long)
' save indexes of members read by GetMemberCodeLines
' to provide a member history
    If pMemberIndex = miLastMemberIndex Then Exit Sub
    ReDim Preserve miMemHist(miMemHistCount)
    miMemHist(miMemHistCount) = pMemberIndex
    miLastMemberIndex = miMemHistCount
    miMemHistCount = miMemHistCount + 1
        
End Sub

Public Function GetMemberCodeLines(pIndex As Long, Optional pWithAddHistory As Boolean = True) As String
Attribute GetMemberCodeLines.VB_Description = "Attempts to return the text of a method from the code section (Index is the  sequential index of the method in the code actual)."

' ... Attempts to return the text of a method from the code section (Index is the  sequential index of the method in the code actual).

Dim x As QuickMemberInfo
Dim lEnd As Long
Dim lStart As Long
Dim sResult As String

    If pIndex <= m_MemberCount Then
        x = QuickMember(pIndex)
        ' v8/9, capture member index for history navigation -----------------
        If pWithAddHistory Then pAddMemberHistory pIndex
        If Len(x.Name) Then
            Let lStart = x.LineStart
            If pIndex = 0 Then
                lStart = lStart - 1
            End If
            lEnd = x.LineEnd
            sResult = pGetTextLines(lStart, lEnd)
        End If
    End If
    ' -------------------------------------------------------------------
    GetMemberCodeLines = sResult
    ' -------------------------------------------------------------------
    sResult = vbNullString
    
End Function ' ... GetMemberCodeLines: String

Public Function GetMembers() As String
Attribute GetMembers.VB_Description = "Attempts to return the text of the code section to include all methods."

' ... Attempts to return the text of the code section to include all methods.

Dim sResult As String

    On Error GoTo ErrHan:
        
    If m_MembersStart < 0 And m_MembersEnd < 0 Or m_MembersStart > m_MembersEnd Then
        Exit Function
    
    End If
    
    sResult = pGetTextLines(m_MembersStart - 1, m_MembersEnd)
    
    GetMembers = sResult

ResumeError:
    
    sResult = vbNullString
    
Exit Function

ErrHan:

    Debug.Print "Code.GetMembers.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Function ' ... GetMembers: String

Public Function ImplementsStringArray() As StringArray
Attribute ImplementsStringArray.VB_Description = "Attempts to return the Implements Statements found in the code file's Declarations section as a StringArray."

' ... Attempts to return the Implements Statements found in the code file's Declarations section as a StringArray.

    If Not m_aImplements Is Nothing Then
        Set ImplementsStringArray = m_aImplements
    Else
        Set ImplementsStringArray = New StringArray
    End If

End Function ' ... ImplementsStringArray: StringArray

Public Property Get MemberCount() As Long
Attribute MemberCount.VB_Description = "Returns a count of the members found."

' ... Returns a count of the members found.
' ... members inclused are subs, functions and properties.

    MemberCount = m_MemberCount

End Property

Public Function MembersStringArray() As StringArray
Attribute MembersStringArray.VB_Description = "Returns a StringArray of all the methods and properties found in the code file / class."

' ... Returns a StringArray of all the methods and properties found in the code file / class.

    Set MembersStringArray = m_aMethods

End Function


Public Sub ReadCodeString(pTheFileText As String, Optional pSourceName As String = "TextString")
Attribute ReadCodeString.VB_Description = "Receives a string that represents some code and then tries to parse it."

' ... Receives a string that represents some code and then tries to parse it.

Dim bOK As Boolean
Dim sErrMsg As String
Dim lng13Found As Long
Dim lng10Found As Long
Dim sDel As String

    ' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    pInit
    ' -------------------------------------------------------------------
    If moCodeString Is Nothing Then
        Err.Raise vbObjectError + 1000, , "The program could not create the code reader."
    End If
    
    moCodeString = pTheFileText
    mxReader.ReadText pTheFileText
    mxReader.Tag = pSourceName
    
    moCodeString.Compact ' ... trim the string worker in case it needs it.
    
    mCodeText = moCodeString.TheString
    
    ' ... check if lines delimited by CRLF or just LF.
    lng10Found = moCodeString.Find(Chr$(10), 1, vbBinaryCompare)
    lng13Found = moCodeString.Find(Chr$(13), 1, vbBinaryCompare)
    
    sDel = vbCrLf ' ... default code file delimiter.
    
    If lng13Found + 1 <> lng10Found Then
        Let bOK = lng10Found > 0
        If bOK = True Then
            sDel = vbLf
        Else
            bOK = moCodeString.Length > 0
            If bOK = False Then
                Err.Raise vbObjectError + 1000, , "The program could not split the code's text: no line delimiter found."
            End If
        End If
    End If
    
    bOK = m_a.FromString(moCodeString, sDel, , , , sErrMsg)
    
    ' -------------------------------------------------------------------
    ' ... v8
    mTextLinesDelimter = sDel
    ' -------------------------------------------------------------------
    
    ' -------------------------------------------------------------------
    ' ... check file was opened ok.
    If bOK = True Then
        ' -------------------------------------------------------------------
        ' ... should have file text contents so can try parsing.

        pParseClass 4 ' ... 4 for raw text, no header, just decs and/or code
    Else
        ' -------------------------------------------------------------------
        ' ... failed to open file so end game.
        Err.Raise vbObjectError, , "The file to parse was not opened:" & vbNewLine & sErrMsg
    End If
    
ResumeError:
Exit Sub
ErrHan:
    mInitialised = False
    Debug.Print "CodeInfo.ReadCodeString.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
        
End Sub ' ... ReadCodeString:

Function FileSize() As Double
    If mxReader Is Nothing Then Exit Function
    FileSize = mxReader.TextByteSize
End Function

Public Sub ReadCodeFile(pTheFileName As String, Optional pSourceName As String = vbNullString)
Attribute ReadCodeFile.VB_Description = "Receives the name of a file, tries to open it as a code file and then tries to parse the contents."

' ... Receives the name of a file, tries to open it as a code file and then tries to parse the contents.

Dim bOK As Boolean
Dim sErrMsg As String
Dim lng13Found As Long
Dim lng10Found As Long
Dim sDel As String

    ' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    pInit
    ' -------------------------------------------------------------------
    If moCodeString Is Nothing Then
        Err.Raise vbObjectError + 1000, , "The program could not create the code reader."
    End If
    
    moCodeString.FromFile pTheFileName, , , bOK, sErrMsg
    ' -------------------------------------------------------------------
    mxReader.ReadFile pTheFileName
    ' -------------------------------------------------------------------
    ' ... v8, hoping to improve searching but
    ' ... may end up impacting original load.
    mCodeText = moCodeString.TheString
    ' -------------------------------------------------------------------
    
    If bOK = True Then
        
        mFileName = pTheFileName
        
        ' ... check if lines delimited by CRLF or just LF.
        lng10Found = moCodeString.Find(Chr$(10), 1, vbBinaryCompare)
        lng13Found = moCodeString.Find(Chr$(13), 1, vbBinaryCompare)
        
        sDel = vbCrLf ' ... default code file delimiter.
        
        If lng13Found + 1 <> lng10Found Then
            Let bOK = lng10Found > 0
            If bOK = True Then
                sDel = vbLf
            Else
                bOK = moCodeString.Length > 0
                If bOK = False Then
                    Err.Raise vbObjectError + 1000, , "The program could not split the file: no line delimiter found."
                End If
            End If
        End If
        
        
        If bOK = True Then
            
            bOK = m_a.FromString(moCodeString, sDel, , , , sErrMsg)
            
            ' -------------------------------------------------------------------
            ' ... v8, hoping to improve search performance...
            modReader.DeriveLinesArray mCodeText, mCodeLines, sDel
            ' -------------------------------------------------------------------
            
        End If
        
    End If
    
    ' -------------------------------------------------------------------
    ' ... v8
    mTextLinesDelimter = sDel
    ' -------------------------------------------------------------------
SkippedV5:

    ' -------------------------------------------------------------------
    ' ... check file was opened ok.
    If bOK = True Then
        ' -------------------------------------------------------------------
        ' ... should have file text contents so can try parsing.
        pParseCode
        If Len(pSourceName) > 0 Then
            mxReader.Tag = pSourceName
        Else
            mxReader.Tag = mname
        End If
    Else
        ' -------------------------------------------------------------------
        ' ... failed to open file so end game.
        Err.Raise vbObjectError, , "The file to parse was not opened:" & vbNewLine & sErrMsg
    End If
    
ResumeError:

Exit Sub
ErrHan:
    
    mInitialised = False
    Debug.Print "CodeInfo.ReadCodeFile.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
        
End Sub ' ... ReadCodeFile:

Private Function pGetTextLines(ByVal pStart As Long, ByVal pEnd As Long) As String
Attribute pGetTextLines.VB_Description = "Private method: Helper to return a requested range of strings from the main String Array holding the entire text file contents."

' ... tries to return a text from the source string given line start and line end.

' ... NOTE: Am hoping to improve this by replacing the loop with a Mid$
' ...       e.g. find the char positions for start and end lines and
' ...       extract the bit inbetween and want to do this to reduce processing time for larger start / end ranges.

Dim sTmp As String
Dim i As Long
Dim lngLineCount As Long
Dim sX As SBuilder ' StringWorker
    
    If Not m_a Is Nothing Then
        Set sX = New SBuilder ' StringWorker
        For i = pStart To pEnd
            sTmp = m_a(i)
            sX.AppendAsLine sTmp, , True 'lngLineCount > 0
            lngLineCount = lngLineCount + 1
        Next i
    End If
    ' -------------------------------------------------------------------
    pGetTextLines = sX
    ' -------------------------------------------------------------------
    sTmp = vbNullString
    If Not sX Is Nothing Then Set sX = Nothing
    
End Function ' ... pGetTextLines: String

Private Sub pInit()
    
' ... Reset everything to defaults for re-use as empty shell.

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    pRelease
    ' -------------------------------------------------------------------
    Set maEvents = New StringArray:     maEvents.Sortable = True
    Set maConstants = New StringArray:  maConstants.Sortable = True
    Set maEnums = New StringArray:      maEnums.Sortable = True
    Set maTypes = New StringArray:      maTypes.Sortable = True
    Set m_aAPIs = New StringArray:      m_aAPIs.Sortable = True ' v5/6
    Set maImplements = New StringArray: maImplements.Sortable = True
    ' -------------------------------------------------------------------
    Set maSubs = New StringArray:       maSubs.Sortable = True
    Set maFunctions = New StringArray:  maFunctions.Sortable = True
    Set maProperties = New StringArray: maProperties.Sortable = True
    ' -------------------------------------------------------------------
    Set m_aMethods = New StringArray:   m_aMethods.Sortable = True ' v5/6
    ' -------------------------------------------------------------------
    Set m_a = New StringArray:          m_a.Sortable = True
    ' -------------------------------------------------------------------
    Set moCodeString = New StringWorker
    ' -------------------------------------------------------------------
    mInitialised = True
    ' -------------------------------------------------------------------
    ' ... v7, avoid re-reading declarative and api sections.
    mbReadDeclarations = False
    mbReadCNTRLStatements = False
    Set mxReader = New TextReader
    Set mNodesCollection = New Collection
    miLastMemberIndex = -1
    miMemHistCount = 0
    
Exit Sub
ErrHan:

    Debug.Print "CodeInfo.pInit.Error: " & Err.Number & "; " & Err.Description
    Err.Clear: Resume Next

End Sub ' ... pInit:

Private Sub pParseClass(Optional pType As Long = 0)
Attribute pParseClass.VB_Description = "(laughingly) Attempts to read text and translate it into useful code related information."

' ... Main method used to read the source content and parse it into subs, functions and properties.

Dim i As Long
Dim j As Long
Dim sItem As String
Dim bBegin As Boolean
Dim bEnd As Boolean
Dim bAttributesDone As Boolean
Dim bHasAttributes As Boolean
Dim bLineIsCommented As Boolean
Dim sFindEnd As String
Dim lngType As Long
Dim sAttribute As String
Dim sAttributes As String
Dim sName As String
Dim lngWordLength As Long
Dim lngFindBracketL As Long
Dim sDeclaration As String
Dim lngLineStart As Long
Dim lngAccessor As Long
Dim lngEditorLineStart As Long
Dim lngFindAs As Long
Dim sValueType As String
Dim lngFoundSpace As Long

Dim sLeft As String
Dim sRight As String
Dim sAccessor As String

Const cAttribute As String = "Attribute"
Const cAttribLen As Long = 9

    ' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    
'    mlngEmptyLineCount = 0
'    mlngCommentedLineCount = 0
'    mlngAttributeLineCount = 0
    m_HeaderStart = -1
    m_HeaderEnd = -1:       m_DecsStart = -1
    m_DecsEnd = -1:         m_MembersStart = -1
    m_MembersEnd = -1:
    
    ' -------------------------------------------------------------------
    ' ... if no array data then nothing doing.
    If m_a Is Nothing Then
        Exit Sub
    End If
    ' -------------------------------------------------------------------
    
    ' v6, set up member indexes and names.
    ReDim mMemberIndexes(1199)
    ReDim mMemberNames(1199)
    ReDim mMemberAttributes(1199)
    
    mMemberIndexes(0) = 0
    mMemberNames(0) = "Declarations"
    mMemberAttributes(0) = "Declarative Section of Source File"
    ' -------------------------------------------------------------------
'    m_MemberCount = 0
'    Erase m_QuickMembers
    ' -------------------------------------------------------------------
    ' ... find where the code actually starts (beyond the class header info).
    m_HeaderStart = 1
    
    If pType <> 4 Then
        
        If pType = 2 Then
            bBegin = True
            bEnd = True
        End If
        
        For i = 1 To m_a.Count  ' ... start at the 2nd line in the file, first line brought us here.
            sItem = m_a(i)
            ' -------------------------------------------------------------------
            ' ... begin / end.
            If bEnd = False Then
                If bBegin = False Then
                    Select Case pType
                        Case 0:                                 ' ... Class.
                            If Left$(sItem, 5) = "BEGIN" Then
                                bBegin = True
                            End If
                        Case 1, 3:                                ' ... Form/UserControl.
                            If Left$(sItem, 5) = "Begin" Then ' VB.Form" Then
                                bBegin = True
                            End If
                    End Select
                Else
                    Select Case pType
                        Case 0:                                 ' ... Class.
                            If Left$(sItem, 3) = "END" Then
                                bEnd = True
                            End If
                        Case 1, 3:                                ' ... Form/UserControl.
                            If Left$(sItem, 3) = "End" Then
                                bEnd = True
                            End If
                    End Select
                End If
            Else
                ' ... attributes.
                If Left$(sItem, cAttribLen) = cAttribute Then
                    ' ... an attribute, so carry on looping.
                    bHasAttributes = True
                    modStrings.SplitStringPair sItem, "=", sLeft, sRight, True, True
                    modStrings.RemoveQuotes sRight
                    sLeft = Mid$(sLeft, 14)
                    Select Case sLeft
                        Case "Name":        mname = sRight
                        Case "Description": mDescription = sRight
                        ' ... other file (e.g. class/form...) properties.
                    End Select
                Else
                    bAttributesDone = True
                    Exit For
                End If
            End If
        Next i
    Else
        If pType = 2 Then
            i = 2
        ElseIf pType = 4 Then
            i = 1 ' ... raw text as code.
        End If
    End If
    ' -------------------------------------------------------------------
    lngEditorLineStart = i
    m_HeaderEnd = i - 1
    ' -------------------------------------------------------------------
    m_DecsStart = i
    m_DecsEnd = m_a.Count   ' ... initialise a value for the end of the declarations (updated as required later).
    ' -------------------------------------------------------------------
    ' ... i should now equal the first code row following the class header.
    ' -------------------------------------------------------------------
    
    ' ... get all the subs, functions and properties from the code.
    ' -------------------------------------------------------------------
    If bBegin And bEnd And bHasAttributes And bAttributesDone Or pType = 2 Or pType = 4 Then
        ' -------------------------------------------------------------------
        
        Do While i < m_a.Count
    
            ' ... k, so now to find where the methods and properties start.
            For i = i To m_a.Count 'For i = i + 1 To m_a.Count
                
                bAttributesDone = False
                
                sItem = LTrim$(m_a(i))  ' ... trim the line.
                bLineIsCommented = Left$(sItem, 1) = "'" Or Left$(sItem, 3) = "Rem"  ' ... check if line is already commented.
                                                                                     ' ... v2, added Rem to comment test.
                If bLineIsCommented Then mlngCommentedLineCount = mlngCommentedLineCount + 1
                If Len(sItem) = 0 Then mlngEmptyLineCount = mlngEmptyLineCount + 1
                
                If bLineIsCommented = False And Len(sItem) Then
                    ' -------------------------------------------------------------------
                    ' ... strip accessor if exists.
                    If Left$(sItem, cm_len_Public) = cm_word_Public Then
                        sAccessor = c_word_Public ' v2
                        sItem = Mid$(sItem, c_len_Public + 2)
                        lngAccessor = 1
                    ElseIf Left$(sItem, cm_len_Private) = cm_word_Private Then
                        sAccessor = c_word_Private ' v2
                        sItem = Mid$(sItem, c_len_Private + 2)
                        lngAccessor = 2
                    ElseIf Left$(sItem, cm_len_Friend) = cm_word_Friend Then
                        sAccessor = c_word_Friend ' v2
                        sItem = Mid$(sItem, c_len_Friend + 2)
                        lngAccessor = 3
                    End If
                    ' -------------------------------------------------------------------
                    ' ... check for method type signature,
                    ' ... if found, generate end method signature to find next
                    ' ... and capture method name.
                    If Left$(sItem, cm_len_Sub) = cm_word_Sub Then ' v4 fix, make sure is single word and not part of word.
                        lngType = 1: mlngSubCount = mlngSubCount + 1
                        lngWordLength = c_len_Sub + 2
                    ElseIf Left$(sItem, cm_len_Function) = cm_word_Function Then ' v4 fix, make sure is single word and not part of word.
                        lngWordLength = c_len_Function + 2
                        lngType = 2: mlngFunctionCount = mlngFunctionCount + 1
                    ElseIf Left$(sItem, cm_len_Property) = cm_word_Property Then ' v4 fix, make sure is single word and not part of word.
                        lngWordLength = c_len_Property + 2 + 4
                        lngType = 3: mlngPropertyCount = mlngPropertyCount + 1
                    End If
                    
                    If lngType > 0 Then
                        ' -------------------------------------------------------------------
                        ' ... we appear to have a member (cautious cos i know i got a bug).
                        If lngAccessor = 0 Then lngAccessor = 1 ' ... default to public accessor if no accessor found.
                        ' -------------------------------------------------------------------
                        If Len(sAccessor) Then sAccessor = sAccessor & Space$(1) ' ... v2
                        ' -------------------------------------------------------------------
                        ' ... set where the declarations end.
                        If m_DecsEnd = -1 Or m_DecsEnd = m_a.Count Then
                            m_DecsEnd = i - 1
                            m_MembersStart = i
                            m_MembersEnd = m_a.Count  ' ... has to be set somewhere, as good a place as any for now.
                        End If
                        ' -------------------------------------------------------------------
                        lngLineStart = i
                        ' -------------------------------------------------------------------
                        sDeclaration = sItem
                        ' -------------------------------------------------------------------
                        mlngMemberDecLineCount = mlngMemberDecLineCount + 1
                        If Right$(sDeclaration, 1) = "_" Then
                            ' -------------------------------------------------------------------
                            ' ... concatenate the declaration's lines into a single line.
                            sDeclaration = Left$(sDeclaration, Len(sDeclaration) - 1)
                            For j = i + 1 To m_a.Count
                                ' -------------------------------------------------------------------
                                ' ... left trim spaces from next line.
                                sDeclaration = sDeclaration & LTrim$(m_a(j))
                                If Right$(sDeclaration, 1) <> "_" Then
                                    i = j
                                    Exit For
                                Else
                                    ' -------------------------------------------------------------------
                                    ' ... trim line extender.
                                    sDeclaration = Left$(sDeclaration, Len(sDeclaration) - 1)
                                    mlngMemberDecLineCount = mlngMemberDecLineCount + 1
                                End If
                            Next j
                        End If
                        
                        If lngType > 1 Then
                            lngFindAs = modStrings.InstrRev(sDeclaration, " As ", Len(sDeclaration))
                            If lngFindAs > 0 Then
                                sValueType = Mid$(sDeclaration, lngFindAs + 4)
                                If lngType = 3 Then ' ... if prop and s/let will have closing bracket at end.
                                    If Right$(sValueType, 1) = ")" Then
                                        If Right$(sValueType, 2) <> "()" Then
                                            sValueType = Left$(sValueType, Len(sValueType) - 1)
                                        End If
                                        If Mid$(sDeclaration, c_len_Property + 2, 1) <> "G" Then
                                            sValueType = ""
                                        End If
                                    End If
                                ElseIf lngType = 2 Then ' ... function, v4.
                                    ' ... make sure there is a return value.
                                    ' ... if no value then last char should be )
                                    If Right$(sValueType, 1) = ")" Then
                                        sValueType = "[Variant assumed]" 'vbNullString
                                    End If
                                End If
                            End If
                        End If
                        ' -------------------------------------------------------------------
                        ' ... extract the member name.
                        ' ... do this from the declaration string rather than the item string
                        ' ... because the member declaration may span several lines (check function fart below).
                        sItem = Mid$(sDeclaration, lngWordLength)              ' ... skip member type signature.
                        
                        lngFindBracketL = InStr(1, sItem, "(")            ' ... find open bracket as member name end signature.
                        
                        If lngFindBracketL > 0 Then
                            sItem = Left$(sItem, lngFindBracketL - 1)   ' ... strip down to name only.
                        End If
                        ' -------------------------------------------------------------------
                        sName = Trim$(sItem) ' ... backup the name.
                        ' -------------------------------------------------------------------
                        ' ... restore the accessor word.
                        sDeclaration = sAccessor & LTrim$(sDeclaration)
                        ' ... complete the end method signature to find.
                        sFindEnd = c_word_End & Space$(1) & Choose(lngType, c_word_Sub, c_word_Function, c_word_Property)
                        
                        ' -------------------------------------------------------------------
                        ' ... find the end of the member.
                        For j = i + 1 To m_a.Count
                            
                            sItem = LTrim$(m_a(j))  ' ... trim the line.
                            bLineIsCommented = Left$(sItem, 1) = "'" Or Left$(sItem, 3) = "Rem"    ' ... check if line is already commented.
                            
                            If bLineIsCommented Then mlngCommentedLineCount = mlngCommentedLineCount + 1
                            If Len(sItem) = 0 Then mlngEmptyLineCount = mlngEmptyLineCount + 1
                            
                            If bLineIsCommented = False And Len(sItem) Then
                                ' -------------------------------------------------------------------
                                ' ... check for line numbering, added this after reading Qwerti60's SetPixel (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73412&lngWId=1)
                                ' ... project where End Sub can be preceded by a line number.
                                ' ... I've written some other code to handle specifically finding End Sub, Function and Property
                                ' ... to get line numbers, it's byte oriented and may be found in mod\modMembers.bas as GetMember_5
                                ' -------------------------------------------------------------------
                                If Val(sItem) > 0 Then
                                    lngFoundSpace = InStr(1, sItem, " ")
                                    If lngFoundSpace > 0 Then
                                        sItem = Mid$(sItem, lngFoundSpace + 1)
                                    End If
                                End If

                                ' -------------------------------------------------------------------
                                ' ... check for attributes.
                                If bAttributesDone = False Then
                                    If Left$(sItem, cAttribLen) = cAttribute Then
                                        mlngAttributeLineCount = mlngAttributeLineCount + 1
                                        If Len(sAttributes) Then sAttributes = sAttributes & "|" '":"
                                        sAttribute = Mid$(sItem, InStr(1, sItem, ".") + 1)
                                        ' -------------------------------------------------------------------
                                        ' ... concatenate attributes with a colon delimiter.
                                        sAttributes = sAttributes & sAttribute
                                    Else
                                        bAttributesDone = True
                                    End If
                                End If
                                
                                ' ... check for end method signature.
                                If Left$(sItem, Len(sFindEnd)) = sFindEnd Then
                                    ' -------------------------------------------------------------------
                                    ' ... v7/8, number of lines given over to member declarations
                                    ' ... including End Sub, Function, Property.
                                    mlngMemberDecLineCount = mlngMemberDecLineCount + 1
                                    ' ... hack, bug with first member's start line in editor
                                    If m_MemberCount = 0 Then
                                        lngLineStart = lngLineStart + 1
                                    End If
                                    ' -------------------------------------------------------------------
                                    ' ... add a new quick member to the quick member array.
                                    ' ... it may be that a stringarray is more appropriate at this stage
                                    ' ... and then parse that into a QuickMember as required.
                                    ' -------------------------------------------------------------------
                                    ReDim Preserve m_QuickMembers(m_MemberCount)
                                    ' -------------------------------------------------------------------
                                    m_QuickMembers(m_MemberCount).Accessor = lngAccessor
                                    m_QuickMembers(m_MemberCount).Attribute = sAttributes
                                    m_QuickMembers(m_MemberCount).Declaration = sDeclaration
                                    m_QuickMembers(m_MemberCount).EditorLineStart = lngLineStart - lngEditorLineStart '+ IIf(m_MemberCount = 0, 1, 0)
                                    m_QuickMembers(m_MemberCount).Index = m_MemberCount
                                    m_QuickMembers(m_MemberCount).LineCount = j - lngLineStart + 1 + IIf(m_MemberCount = 0, 1, 0)
                                    m_QuickMembers(m_MemberCount).LineEnd = j '+ IIf(m_MemberCount = 0, 1, 0)
                                    m_QuickMembers(m_MemberCount).LineStart = lngLineStart
                                    m_QuickMembers(m_MemberCount).Name = sName
                                    m_QuickMembers(m_MemberCount).Type = lngType
                                    m_QuickMembers(m_MemberCount).ValueType = sValueType
                                    
                                    ' -------------------------------------------------------------------
'''                                    ' ...                    Name     :    Quick Member index     :    Type             :    Accessor             :    Value/Return Type
'''                                    m_aMethods.AddItemString sName & ":" & CStr(m_MemberCount) & ":" & CStr(lngType) & ":" & CStr(lngAccessor) & ":" & sValueType
                                    ' ... v8, changed delimiter from : to | so could add file paths later.
                                    ' ...                    Name     |    Quick Member index     |    Type             |    Accessor             |    Value/Return Type   |  Line Start
                                    m_aMethods.AddItemString sName & "|" & CStr(m_MemberCount) & "|" & CStr(lngType) & "|" & CStr(lngAccessor) & "|" & sValueType & "|" & lngLineStart

                                    ' -------------------------------------------------------------------
                                    i = j ' ... update i for next go.
                                    ' -------------------------------------------------------------------
                                    m_MemberCount = m_MemberCount + 1
                                    ' -------------------------------------------------------------------
                                    ' v6
                                    mMemberIndexes(m_MemberCount) = lngLineStart - m_DecsStart - 1 '  m_MembersStart
                                    mMemberNames(m_MemberCount) = Choose(lngType, "Sub", "Function", "Property") & ": " & sName & IIf(Len(sValueType) > 0, ": " & sValueType, "")
                                    mMemberAttributes(m_MemberCount) = sAttributes
                                    
                                    Exit For
                                    
                                End If
                                
                            End If
                            
                        Next j
                        
                    End If
                
                End If
                
                sValueType = vbNullString
                sAttributes = vbNullString
                sAttribute = vbNullString
                sAccessor = vbNullString ' v2.
                lngType = 0
                lngAccessor = 0
                
            Next i
        
        Loop
        
    Else
        ' ... invalid header read.
        Exit Sub
    End If
    
    If m_aMethods.Count > 0 Then
        m_aMethods.Sort
    End If

ResumeError:
    ' -------------------------------------------------------------------
    ' v6, resize member indexes and names.
    ReDim Preserve mMemberIndexes(m_MemberCount)
    ReDim Preserve mMemberNames(m_MemberCount)
    ' ... clean up
    sDeclaration = vbNullString
    sFindEnd = vbNullString
    sName = vbNullString
    sValueType = vbNullString
    sAttributes = vbNullString
    sAttribute = vbNullString
    sAccessor = vbNullString
    
    lngType = 0&
    lngAccessor = 0&
    lngLineStart = 0&
    lngEditorLineStart = 0&
    lngFindAs = 0&
    lngFoundSpace = 0&
    
    ' -------------------------------------------------------------------
    ClearMemory
    ' -------------------------------------------------------------------
Exit Sub

ErrHan:

    Debug.Print "Code.pParseClass.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pParseClass:

Private Sub pParseCode()
Attribute pParseCode.VB_Description = "Private method: Intermediary helper between public ParseCode and private pParseClass."

' ... precursor to pParseClass ...

Dim sLine As String
    
    If m_a Is Nothing Then
        Exit Sub
    End If

    If m_a.Count > 0 Then
        sLine = m_a(1)
        If Left$(sLine, 17) = "VERSION 1.0 CLASS" Then
            mType = "Class" ' ... Revision 1.
            pParseClass 0   ' ... class.
        ElseIf Left$(sLine, 12) = "VERSION 5.00" Then
            ' -------------------------------------------------------------------
            ' ... Revision 1.
            If Not moCodeString Is Nothing Then
                If moCodeString.Find("Begin VB.UserControl") > 0 Then
                    mType = "User Control"
                Else
                    mType = "Form"
                End If
            End If
            ' -------------------------------------------------------------------
            pParseClass 1   ' ... form/usercontrol.
        ElseIf Left$(sLine, 19) = "Attribute VB_Name =" Then
            mType = "Module"    ' ... Revision 1.
            pParseClass 2       ' ... module.
        Else
            MsgBox "Unknown File Type", vbInformation, "Parse Code File"
        End If
    End If
    
End Sub ' ... pParseCode:

Private Sub pParseForAPIs()

' ... This method attempts to read the declarative section of the code file
' ... and record APIs declared.

' ... v7,   update
' ...       have added extra info to the output describing each api
'           line number and editor line number.
' ...       this will impact methods needing to parse this output.

Dim i As Long
Dim sLine As String
Dim bLineIsCommented As Boolean
Dim lngAccessor As Long
Dim sDeclaration As String
Dim j As Long
Dim sItem As String
Dim lngType As Long
Dim lngWordLength As Long
Dim bHaveDeclare As Boolean
Dim sType As String
Dim sAccessor As String

' -------------------------------------------------------------------
' ... v7, ... line number and editor line number.
Dim lngLineNo As Long
Dim lngEditorLineNo As Long

    On Error GoTo ErrHan:
    
    If Not m_a Is Nothing Then
    
        If m_DecsStart < 0 And m_DecsEnd < 0 Or m_DecsStart > m_DecsEnd Then
            Exit Sub
        
        End If
        
        For i = m_DecsStart To m_DecsEnd
            sLine = LTrim$(m_a(i))                      ' ... capture and trim the line.
            bLineIsCommented = Left$(sLine, 1) = "'" Or Left$(sLine, 3) = "Rem"    ' ... check if line is already commented.
            If bLineIsCommented = False And Len(sLine) Then
                lngAccessor = 1                             ' ... default to public accessor.
                lngType = 0                                 ' ... default type of nothing, later, 1 = sub, 2 = function.
                sAccessor = cm_word_Public
                If Left$(sLine, cm_len_Public) = cm_word_Public Then
                    sLine = Mid$(sLine, c_len_Public + 2) ' + 8)
                ElseIf Left$(sLine, cm_len_Private) = cm_word_Private Then
                    sLine = Mid$(sLine, c_len_Private + 2) ' + 8)
                    lngAccessor = 2
                    sAccessor = cm_word_Private
                End If
                
                bHaveDeclare = False
                If Left$(sLine, 8) = "Declare " Then
                    bHaveDeclare = True
                    sLine = Mid$(sLine, 9)
                End If
'                 -------------------------------------------------------------------
'                 ... check for method type signature,
'                 ... if found, generate end method signature to find next
'                 ... and capture method name.
'                sType = cm_word_Sub
                If Left$(sLine, cm_len_Sub) = cm_word_Sub Then
                    lngType = 1
                    lngWordLength = c_len_Sub + 2
                ElseIf Left$(sLine, cm_len_Function) = cm_word_Function Then
                    lngWordLength = c_len_Function + 2
                    lngType = 2
'                    sType = cm_word_Function
                End If
                    
                If bHaveDeclare = True Then
                    ' ... looks like it could be an api dec.
                    sDeclaration = sLine
                    mlngAPIDecLineCount = mlngAPIDecLineCount + 1
                    ' -------------------------------------------------------------------
                    ' ... caution: if the underscore present is it in a comment?
                    ' ...          if so, things will go awry.
                    If Right$(sDeclaration, 2) = " _" Then
                        ' -------------------------------------------------------------------
                        ' ... concatenate the declaration's lines into a single line.
                        sDeclaration = Left$(sDeclaration, Len(sDeclaration) - 1)
                        For j = i + 1 To m_a.Count
                            ' -------------------------------------------------------------------
                            ' ... left trim spaces from next line.
                            sDeclaration = sDeclaration & LTrim$(m_a(j))
                            If Right$(sDeclaration, 1) <> "_" Then
                                i = j
                                Exit For
                            Else
                                ' -------------------------------------------------------------------
                                ' ... trim line extender.
                                sDeclaration = Left$(sDeclaration, Len(sDeclaration) - 1) '1)
                                mlngAPIDecLineCount = mlngAPIDecLineCount + 1
                            End If
                        Next j
                    End If
                    
                    ' ... name of the sub/function.
                    sItem = Mid$(sDeclaration, lngWordLength)              ' ... skip member type signature.
                    ' -------------------------------------------------------------------
                    sItem = modStrings.LeftOfComment(sItem, "'", True) ' v6, added, remove any comments at end of declaration.
                    ' -------------------------------------------------------------------
'                    sItem = Trim$(sItem) & "|" & CStr(lngType) & "|" & CStr(lngAccessor)
                    ' ... v7, adding line number from top of source and editor line number from start of declarations section.
                    lngLineNo = i + 1
                    lngEditorLineNo = lngLineNo - m_DecsStart
                    sItem = Trim$(sItem) & "|" & CStr(lngType) & "|" & CStr(lngAccessor) & "|" & CStr(lngLineNo) & "|" & CStr(lngEditorLineNo)
'                    Debug.Print sItem
                    
                    m_aAPIs.AddItemString sItem & "|" & mname & "|" & mType
                    If Len(msAPIs) > 0 Then msAPIs = msAPIs & vbNewLine
                    msAPIs = msAPIs & sAccessor & "Declare " & sDeclaration
                    ' -------------------------------------------------------------------
                
                End If
                
            End If
        Next i
        
    End If

    If Not m_aAPIs Is Nothing Then
        If m_aAPIs.Count > 0 Then
            m_aAPIs.Sort
        End If
    End If

ResumeError:

Exit Sub

ErrHan:

    Debug.Print "Code.pParseForAPIs.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub ' ... pParseForAPIs:

Private Sub pRelease()
    
' ... release all resources from memory ahead of re-use and terminate instance.

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    m_MemberCount = 0
    mInitialised = False
    ' -------------------------------------------------------------------
    ' ... note: could reduce object release to single line for less lines.
    If Not maEvents Is Nothing Then
        Set maEvents = Nothing
    End If
    If Not maConstants Is Nothing Then
        Set maConstants = Nothing
    End If
    If Not maEnums Is Nothing Then
        Set maEnums = Nothing
    End If
    If Not maTypes Is Nothing Then
        Set maTypes = Nothing
    End If
    If Not m_aAPIs Is Nothing Then
        Set m_aAPIs = Nothing
    End If
    If Not maImplements Is Nothing Then
        Set maImplements = Nothing
    End If
    If Not maSubs Is Nothing Then
        Set maSubs = Nothing
    End If
    If Not maFunctions Is Nothing Then
        Set maFunctions = Nothing
    End If
    If Not maProperties Is Nothing Then
        Set maProperties = Nothing
    End If
    If Not m_a Is Nothing Then
        Set m_a = Nothing
    End If
    If Not m_aMethods Is Nothing Then
        m_aMethods.Clear
        Set m_aMethods = Nothing
    End If
    If Not m_aImplements Is Nothing Then
        Set m_aImplements = Nothing
    End If
    If Not moCodeString Is Nothing Then
        Set moCodeString = Nothing
    End If
    ' -------------------------------------------------------------------
    ' ... Revision 1.
    mType = vbNullString
    mname = vbNullString
    mDescription = vbNullString
    ' -------------------------------------------------------------------
    ' ... Revision 2.
    mFileName = vbNullString
    ' -------------------------------------------------------------------
    ' ... Revision 3.
    mCodeText = vbNullString
    mTextLinesDelimter = vbNullString
    Erase mCodeLines
    ' -------------------------------------------------------------------
    ' ... Revision 6.
    Erase mMemberIndexes
    Erase mMemberNames
    Erase mMemberAttributes
    ' -------------------------------------------------------------------
    ' ... Revision 7.
    mTextLinesDelimter = vbNullString
    ' -------------------------------------------------------------------
    ' ... Revision 8
    Erase m_QuickMembers
    
    msEvents = vbNullString
    tmpTypes = vbNullString
    msTypeName = vbNullString
    msConstants = vbNullString
    msTypes = vbNullString
    msEnums = vbNullString
    msVars = vbNullString
    mMenuStructure = vbNullString
    mMenuMethods = vbNullString
    
    mlngMemberDecLineCount = 0&
    mlngConstantDecLineCount = 0&
    mlngVarDecLineCount = 0&
    mlngAPIDecLineCount = 0&
    mlngTypeDecLineCount = 0&
    mlngEnumDecLineCount = 0&
    mlngEventDecLineCount = 0&
    mlngCNTRLStatementCount = 0&
    mlngAttributeLineCount = 0&
    mlngEmptyLineCount = 0&
    mlngCommentedLineCount = 0&
    
    mlngSubCount = 0&
    mlngFunctionCount = 0&
    mlngPropertyCount = 0&
    
    msTextRTF = vbNullString
    msAPIs = vbNullString
    msConstants9 = vbNullString
    Set mxReader = Nothing
    Set mNodesCollection = Nothing
    miLastMemberIndex = 0
    miMemHistCount = 0
    Erase miMemHist
    msQuickReport = vbNullString
    
Exit Sub
ErrHan:
    
    Debug.Print "CodeInfo.pRelease.Error: " & Err.Number & "; " & Err.Description
    Err.Clear: Resume Next

End Sub ' ... pRelease:

Friend Property Get QuickMember(pIndex As Long) As QuickMemberInfo
Attribute QuickMember.VB_Description = "Attempts to return a quick member info structure at the given index from the main quick member array."

' ... Attempts to return a quick member info structure at the given index from the main quick member array.

Dim x As QuickMemberInfo
    
    If m_MemberCount > 0 Then
        If pIndex <= m_MemberCount Then
            With m_QuickMembers(pIndex)
                x.Accessor = .Accessor
                x.Attribute = .Attribute
                x.Declaration = .Declaration
                x.EditorLineStart = .EditorLineStart
                x.Index = .Index
                x.LineCount = .LineCount
                x.LineEnd = .LineEnd
                x.LineStart = .LineStart
                x.Name = .Name
                x.Type = .Type
                x.ValueType = .ValueType
            End With
        End If
    End If
    
    QuickMember = x
    
End Property ' ... QuickMember: QuickMemberInfo

Public Property Get TypesString() As String
Attribute TypesString.VB_Description = "Returns a vbCrLf delimited string containing all Type declarations found in Declarations section."

' ... Returns a vbCrLf delimited string containing all Type declarations found in Declarations section.

    TypesString = msTypes

End Property ' ... TypesString: String
