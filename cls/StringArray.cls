VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class to manage an Array of Strings, no API/TLB or VB6 only VBA functions."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' ------------------------------------------------------------
' Name:         StringArray (3)
' Purpose:      A Class to manage an Array of Strings.
' Author:       D.C.
' Date:         Friday 25 June 2010
' ------------------------------------------------------------
' Version:      005
' Date:         22 May 2011
' Note:         Added public property name.
'               Added public property tag
' -------------------------------------------------------------------
' Version:      004
' Date:         16 May 2011
' Note:         fixed bug in Method::Exists (didn't actually return the exists test result).
'               have made changes to make this class compatible with vb5;
'               essentially, all properties and functions returning arrays have
'               had their return type commented so they return a Variant instead,
'               may also have commented entire methods (FromArray? ToArray?).
'               ToFile method has had its escape char param default value changed to help
'               because it caused an error when being read on windows(chinese) bug noted by Tony Yang
' -------------------------------------------------------------------
' Version:      003
' Date          07 March 2011
' Note;         Ammended pEscapeCHar param default value from ¬ to : in methods
'               ToFile and ToString on advice from TonyYang w/r Windows(Chinese).
' -------------------------------------------------------------------
' Version:      002 (see items marked v2 & v3 within source).
' Date:         29 August 2010
' Note:         found a bug in the file writing (was testing for illegal characters in filename and including the path delimiters in the test).
'               added pCheckFileName from last StringBuilder release and altered pWriteToFile to use this for the same test.
'               also found the call to pWriteToFile within ToFile incorrect (missing a parameter).
'               ItemAsNumberValue didn't actually bother returning the numeric value.
'               pQuickSortOnStringArray: amended code in attempt to sort alphabetically with upper and lower case strings.
'               pReadTextFile: amended Open File to use Binary to sidestep bad reads when unicode chars exist in file's text.
'               Remove: had added Compact to list of instructions but then used m_Count for loop which had now been set to zero.
'
'               These issues were found when leaning on this class in another program.
'               (Public) Descriptive Attributes have also been re-written as they were previously lost doing a bulk copy of code
'               from the IDE (which rather ungraciously ignores such things).
'
' Date:         28 Oct 2010.
'               Removed vb6 only vba methods by implementing 2 new methods:
'               Function pReplace: String and Sub.pReplaceChars
'               and dropped use of FormatDateTime when sorting by date.
'
' -------------------------------------------------------------------

Option Explicit

' Credits:
'
'           Thank You (in no particular order)...
'           Squirm and Optikon:     QuickSort algorithms on XtremeVBTalk: http://www.xtremevbtalk.com/showthread.php?t=78889&highlight=Quick+Sort
'           G. van den Hoven:       StringArrayToString (Method: ToString). Planet Source Code: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=13008&lngWId=1
'           RDe:                    A Complete set of VB6 functions not available in VB5
'                                   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=71812&lngWId=1

' Notes:
'           This is my bodge to make a class i can use to manage a bunch of VB6 strings.
'           It's got Sorting, Output Formatting and File Reading / Writing.
'           It's got a Type Ahead / Auto-Complete helper method and can be used to store
'           delimited strings which can then be returned as a new instance of self.
'           It can be instructed to permit only unique string values and it can deny sorting
'           to keep items in their original order.
'
'           It doesn't have Keyed values.
'           It may have a few bugs and be a bit incomplete, it's a work in progress.
'           Maybe better as an outline of interfaces that could be used
'           as a template to something similar.
'
'           You might ask why?  Please don't, it was just some fun that
'           seems to have turned out ok.  Collection coding would have been
'           far easier and the Dictionary Object could be used instead.
'           According to my documentation Collections can perform faster than
'           Arrays post 1000 items.
'
'           I used my Member Composer program to generate the interfaces,
'           likewise it created a load of help topic files but I haven't compiled them
'           as I have made some small changes to the public methods and their parameters
'           and so some of the topic files need updating, I've left them unaltered in the
'           hlp folder just in case.
'           -------------------------------------------------------------------
'           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73200&lngWId=1
'           -------------------------------------------------------------------
'
' -------------------------------------------------------------------
' Note: to Code Browser release 2:
'       Considered reducing functionality to minimum used within the browser thing
'       in order to reduce size...  I didn't because I may want to do more stuff later
'       on with this class in the context of the code browser.
'       Within the context of VB5, this class ain't much good at the moment
'       on account of VB5 not allowing arrays as return values to functions
'       but I done made a start that way.
' -------------------------------------------------------------------
'
' ... Properties.
'
'   ... Public.
'
' P   Count: Long.                      ' ... Returns the number of items currently stored.
' P   DuplicatesAllowed: Boolean.       ' ... When False forces unique test on items before adding them.
' P   Item: String.                     ' ... Sets / Returns the Item at the given Index (default property).
' P   PreAllocationSize: Long.          ' ... Sets / Returns the buffer Pre-Allocation Size for ReDim Preserve Statements.
' P   Sortable: Boolean.                ' ... Write-Once Property to allow Sorting (True) or not (False).
' P   ToArray: Boolean.                 ' ... Provides access to a copy of the internal array of strings currently stored.
'
'
' ... Methods.
'
'   ... Public.
'
' F   AddItemString: Boolean.           ' ... Add a String value to the internal array.
' F   Clear: Boolean.                   ' ... Reset the internal array with no elements/data.
' S   Compact.                          ' ... Shrinks the internal buffer to the number of items currently stored..
' F   EditItemString: Boolean.          ' ... Changes the String Value of an Item to a new value if queried value exists.
' F   Exists: Boolean.                  ' ... As pExists but with option to compare text or binary values.
' F   FindClosestItem: Long.            ' ... Returns the position of the first item partially matching the search string.
' F   FromArray: Boolean.               ' ... Copies an Array String() to the buffer.
' F   FromFile: Boolean.                ' ... Reads the contents of a text file and converts it to the instance data.
' F   FromString: Boolean.              ' ... Converts a String into a StringArray given an Optional Delimiter (default = comma).
' F   IndexExists: Boolean.             ' ... Describes whether a given index exists within the one based array (True) or Not (False).
' F   Item: String.                     ' ... Checks valid array and index to return item in array.
' F   ItemAsDateValue: Date             ' ... Returns a CVDate conversion of a string by index from the string() buffer.
' F   ItemAsNumberValue: Double.        ' ... Returns a Val conversion of a string by index from the string() buffer.
' F   ItemAsStringArray: StringArray.   ' ... Returns a string converted into a StringArray via a Delimiter that defaults to a colon ( : ) if omitted.
' F   Remove: Boolean.                  ' ... Deletes an Item from the internal buffer (optionally deletes all occurrances or first only).
' F   Sort: Boolean.                    ' ... Sorts the items in the buffer as text using a Selection Sort algorithm.
' F   SortAsDate: Boolean.              ' ... Attempts to sort buffer by date value.
' F   SortAsNumber: Boolean.            ' ... Attempts to sort buffer by numeric value.
' F   ToFile: Boolean.                  ' ... Writes the contents of the buffer to disk.
' F   ToString: String.                 ' ... Returns a concatenation of the buffer items with various options.
'
'   ... Private.
'
' F   pAutoCompleteText: Long.          ' ... Searches the array for an item (partial or exact), if found, index returned along with full string (for Partial matching)..
' F   pBinarySearch: Long.              ' ... Searches a one-dimensional string array for the search string and returns its index if found or -1.
' ... v2
' F   pCheckFileName: Boolean           ' ... Checks a file name for invalid characters.
' ...
' F   pExists: Boolean.                 ' ... Checks to see if a value exists in the current data.
' F   pKnownItem: String.               ' ... Returns a string from the buffer with no testing, assumes use when index is known to exist.
' S   pQuickSortOnLongArray:            ' ... .
' S   pQuickSortOnStringArray:          ' ... .
' ... v2
' F   pReplace: String                  ' ... replaces one string within another.
' S   pReplaceChars:                    ' ... method used by pReplace.
' ...
' S   pSortNumber:                      ' ... .
' F   pTestArray: Boolean.              ' ... Returns a Boolean indicating that the internal array is OK to use (True) or Not (False).
'
' F   pFileExists: Boolean.             ' ... Tests for file existence, returns True if found else False.
' F   pReadTextFile: String.            ' ... The contents of a text file as a string.
' F   pWriteTextFile: Boolean.          ' ... Write the contents of a string to a file, optionally in Append mode.
'
'
' -------------------------------------------------------------------
' ... Local Variables.
Private m_StringArray() As String       ' ... private field for property ArrayString in vb6.

Private m_Row As Long                   ' ... private field for looping through array.

Private m_Count As Long                 ' ... private field for property Count.
Private m_DuplicatesAllowed As Boolean  ' ... private field for property DuplicatesAllowed.
Private m_Sortable As Boolean           ' ... private field for property Sortable.
Private m_bCompacted As Boolean
Private m_PreAllocationSize As Long ' ... private field for property PreAllocationSize.

' -------------------------------------------------------------------
' ... StringBuilder for AddItemString when no duplicates allowed.
Private m_TheBufferCapacity As Long
Private m_TheBufferChunkSize As Long
Private m_TheString As String
Private m_TheTextLength As Long

' -------------------------------------------------------------------
' ... Constants.
Private Const c_def_DuplicatesAllowed As Boolean = True ' ... default value for property DuplicatesAllowed.
Private Const c_def_Sortable As Boolean = False         ' ... default value for property Sortable.
Private Const c_def_PreAllocationSize As Long = 1024 ' ... default value for property PreAllocationSize.

Private Const c_def_InvalidFileNameChars As String = "`!$%^&*()-=+[]{}'#@~;,.<>/?\| " ' ... needs " as well, see pCheckFileName.

' -------------------------------------------------------------------
' ...Properties
Public Name As String
Public Tag As String
' -------------------------------------------------------------------

' -------------------------------------------------------------------

' Function:        AddItemString
' Returns:         Boolean.
' Description:     Add a String value to the internal array.

Public Function AddItemString(ByVal pItemString As String, Optional pExistsTest As Byte = 0, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute AddItemString.VB_Description = "Add a String value to the internal array."

'... Parameters.
'    V__ pItemString: String         ' ... The String to Add to the Array.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim lTheStringLength As Long
Dim sBufferItem As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    ' -------------------------------------------------------------------
    
    If m_DuplicatesAllowed = False Then
    
        ' Note:
        '       If unique only then test item doesn't exist.
        '       In order of performance, fastest first,
        '       three ways of testing
'        If pTestArray = False Then
'            m_Count = 0
'            ReDim m_StringArray(m_Count)
'        End If
        
        If pExistsTest = 0 Then
        
            ' ... Try using (StringBuilder type) buffer and instr.
            Let sBufferItem = Chr(1) & pItemString & Chr(2)
'            Let sBufferItem = pItemString ' V3.
            
            If InStr(1, m_TheString, sBufferItem, vbBinaryCompare) > 0 Then
            
                Let pErrMsg = "Item Exists within Unique Collection."
                Exit Function
                
            End If
            
            Let lTheStringLength = Len(sBufferItem)
            
            If lTheStringLength > 2 Then
                
                ' -------------------------------------------------------------------
                If (m_TheTextLength + lTheStringLength) > m_TheBufferCapacity Then
                    
                    If (lTheStringLength > m_TheBufferChunkSize) Then
                        
                        Let m_TheString = m_TheString & Space$(lTheStringLength)
                    
                    Else
                        
                        Let m_TheString = m_TheString & Space$(m_TheBufferChunkSize)
                    
                    End If
                
                End If
                
                ' -------------------------------------------------------------------
                If m_TheTextLength = 0 Then
                    
                    Mid$(m_TheString, 1, lTheStringLength) = sBufferItem
                
                Else
                    
                    Mid$(m_TheString, m_TheTextLength + 1, lTheStringLength) = sBufferItem
                
                End If
                
                ' -------------------------------------------------------------------
                Let m_TheTextLength = m_TheTextLength + lTheStringLength
                
                Let m_TheBufferCapacity = Len(m_TheString)
            
            End If
        
        'ElseIf pExistsTest = 1 Then
        ElseIf pExistsTest = 1 Or pExistsTest = 2 Then  ' ... added secondary test for 2 as avoiding use
                                                        ' ... of vba6.join function for compatibility with vb5.
        
            ' ... Or, try finding with looping.
            
            ' ... If we knew we were sorted a Binary
            ' ... Search algorithm could be useful but then we'd
            ' ... have to insert the item in its rightful place.
            ' ... Could use Kernal32.RtlMoveMemory(AKA [Bruce McKinney's] CopyMemory) API
            ' ... to help shifting a block of data to this end but prefer to be API free.
            
            If pExists(pItemString) = True Then
    
                Let pErrMsg = "Item Exists within Unique Collection."
    
                Exit Function
    
            End If
        
            
'        ElseIf pExistsTest = 2 Then ' ... see immediately above for comment reason.
'
'            ' ... Try finding without looping.
'
'            ' ... Unforunately this ended up taking the most time
'            ' ... but requires only one line of code so it kind of
'            ' ... looked better to start with.
'
'            If m_Count > 0 Then
'
'                If InStr(Join(m_StringArray), pItemString) > 0 Then
'                    Let pErrMsg = "Item Exists within Unique Collection."
'                    Exit Function
'
'                End If
'
'            End If
        
        
        End If
        
    End If
    
    ' -------------------------------------------------------------------
    If m_Count Mod m_PreAllocationSize = 0 Then
        
        ' -------------------------------------------------------------------
        ' ... allocate more room in the buffer.
        ' -------------------------------------------------------------------
        ' ... haven't observed much performance gain below
        ' ... 100,000 items; this approach required Compact to be
        ' ... introduced; ReDim Preserve m_StringArray(m_Count)
        ' ... seemed perfectly ok and simplified the code.
        ' ... 1,000,000 items saw a performance gain but made my
        ' ... computer's fan whirl into action, too much for my liking
        ' ... so I stopped testing beyond 500,000 items for fear I
        ' ... might burn out my processor (running an Advent 1.6 GHz dual core).
        ' ... Would be interested to know what faster processors could achieve.
        ' -------------------------------------------------------------------
        ReDim Preserve m_StringArray(m_Count + m_PreAllocationSize)
        
        Let m_bCompacted = False
        
    End If
    
    Let m_StringArray(m_Count) = pItemString
    
    Let m_Count = m_Count + 1
    
    Let bOK = True
    
    ' -------------------------------------------------------------------
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let AddItemString = bOK
    
    ' -------------------------------------------------------------------

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    
    Debug.Print "StringArray.AddItemString", Err.Number, Err.Description
    Resume ErrResume:
Resume

End Function ' ... AddItemString: Boolean.

Private Sub Class_Initialize()
    
    Let m_bCompacted = False
    Let m_TheBufferChunkSize = 8192
    
    Let m_PreAllocationSize = c_def_PreAllocationSize    ' ... 1024?
    Let m_DuplicatesAllowed = c_def_DuplicatesAllowed    ' ... TRUE.
    Let m_Sortable = c_def_Sortable                      ' ... FALSE.
    
End Sub

' Function:        Clear
' Returns:         Boolean.
' Description:     Reset the internal array with no elements/data.

Public Function Clear(Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute Clear.VB_Description = "Reset the internal array with no elements/data."

' -------------------------------------------------------------------

Dim bOK As Boolean      ' ... Returns success or failure of this method.
Dim sErrMsg As String   ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... AddItemString, StringBuilder fields.
    Let m_TheString = vbNullString
    Let m_TheTextLength = 0
    Let m_TheBufferCapacity = 0
    
    ' -------------------------------------------------------------------
    Let m_bCompacted = False
    
    Let m_Count = 0
    
    Erase m_StringArray
    
    ' -------------------------------------------------------------------
    
    Let bOK = True
    
    ' -------------------------------------------------------------------
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let Clear = bOK
    
    ' -------------------------------------------------------------------

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.Clear", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... Clear: Boolean.

' Sub:             Compact
' Description:     Shrinks the internal buffer to the number of items currently stored.

Public Sub Compact()
Attribute Compact.VB_Description = "Shrinks the internal buffer to the number of items currently stored."

' Note:     This method became neccessary to the code after I
'           added configurable buffer pre-allocation.

    On Error GoTo ErrHan:
    
    If m_Count >= 1 Then

        If m_bCompacted = False Then
            ReDim Preserve m_StringArray(m_Count - 1)
            Let m_bCompacted = True
        End If

    End If

ErrResume:


Exit Sub
ErrHan:

    Debug.Print "StringArray.Compact", Err.Number, Err.Description
    Resume ErrResume:


End Sub ' ... Compact.

' Property Get:    Count
' Type:            Long.
' Description:     Returns the number of items currently stored.

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items currently stored."

' Note:     following configurable buffer pre-allocation
'           this property returns the only true way to
'           read the number of items added and UBound(m_StringArray)
'           needs Compact to be run first to be useable elsewhere in code.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    Let Count = m_Count

Exit Property
ErrHan:

    Debug.Print "StringArray.Count", Err.Number, Err.Description

End Property ' ... Count: Long.

' Property Get:    DuplicatesAllowed
' Type:            Boolean.
' Description:     When False forces unique test on items before adding them.

Public Property Get DuplicatesAllowed() As Boolean
Attribute DuplicatesAllowed.VB_Description = "When False forces unique test on items before adding them."

    On Error GoTo ErrHan:

    Let DuplicatesAllowed = m_DuplicatesAllowed

Exit Property
ErrHan:

    Debug.Print "StringArray.DuplicatesAllowed", Err.Number, Err.Description

End Property ' ... DuplicatesAllowed: Boolean.

' Property Let:    DuplicatesAllowed

Public Property Let DuplicatesAllowed(ByVal pNewValue As Boolean)

    On Error GoTo ErrHan:

    Let m_DuplicatesAllowed = pNewValue
'   Call PropertyChanged("DuplicatesAllowed")

Exit Property
ErrHan:

    Debug.Print "StringArray.DuplicatesAllowed", Err.Number, Err.Description

End Property ' ... DuplicatesAllowed: Boolean.

' Function:        EditItemString
' Returns:         Boolean.
' Description:     Changes the String Value of an Item to a new value if queried value exists.

Public Function EditItemString(ByVal pItemString As String, ByVal pNewValue As String, Optional ByVal pEditAll As Boolean = False, Optional ByVal pEditCount As Long = 0, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute EditItemString.VB_Description = "Changes the String Value of an Item to a new value if queried value exists"

'... Parameters.
'    V__ pItemString: String         ' ... The item to find.
'    V__ pNewValue: String           ' ... The replacement value.
'    VO_ pEditAll: Boolean           ' ... Instruction to edit all search values when true else just first.
'    VO_ pEditCount: Long            ' ... Returns the number of items edited.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim lngEditCount As Long
Dim bArrayOK As Boolean

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
        
    Let bArrayOK = pTestArray
    
    ' -------------------------------------------------------------------
    
    If bArrayOK = True Then
        
        Call Compact
        
        For m_Row = 0 To UBound(m_StringArray)
            
            If m_StringArray(m_Row) = pItemString Then
            
                Let lngEditCount = lngEditCount + 1
            
                Let m_StringArray(m_Row) = pNewValue
                
                Let bOK = True
    
                If pEditAll = False Then
                    
                    Exit For
                
                End If
                
            End If
            
        Next m_Row
        
    Else
    
        Let sErrMsg = "There are no items available to edit."
        
    End If
    
    ' -------------------------------------------------------------------

ErrResume:

    Let pEditCount = lngEditCount
    Let pErrMsg = sErrMsg
    Let EditItemString = bOK
    
    ' -------------------------------------------------------------------

Exit Function
ErrHan:

'''    Let lngEditCount = 0 ' ... appropriate???
    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.EditItemString", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... EditItemString: Boolean.

' Function:        Exists
' Returns:         Boolean.
' Description:     As Exists but with option to compare text or binary values.

Public Function Exists(ByVal pItemString As String, Optional ByVal pBinarySearch As Boolean = True, Optional ByVal pStartPos As Long = 1, Optional ByRef pItemIndex As Long = 0, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute Exists.VB_Description = "As pExists but with option to compare text or binary values.\r\n"

'... Parameters.
'    V__ pItemString: String         ' ... The item to find.
'    VO_ pBinarySearch: Boolean      ' ... Option to search and compare by binary (True) or text (False) value.
'    RO_ pItemIndex: Long            ' ... Returns the one based index of the item if found else 0.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim lngItemIndex As Long
Dim lngStartPos As Long
Dim bIndexed As Boolean
Dim strItem As String
Dim strSearch As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    ' -------------------------------------------------------------------

    If pBinarySearch = True Then
        
        Let bOK = pExists(pItemString, pStartPos, pItemIndex, sErrMsg)
    
    Else
                
        ' -------------------------------------------------------------------
        
        Let bIndexed = IndexExists(pStartPos, sErrMsg)
        
        ' -------------------------------------------------------------------
        
        If bIndexed = True Then
                        
            ' -------------------------------------------------------------------
            
            Let lngStartPos = pStartPos - 1
            
            Let strSearch = pItemString
            
            ' -------------------------------------------------------------------
            
            For m_Row = lngStartPos To m_Count - 1 ' UBound(m_StringArray)
                
                ' -------------------------------------------------------------------
                Let strItem = m_StringArray(m_Row)
                
                ' -------------------------------------------------------------------
                Let bOK = StrComp(strItem, strSearch, vbTextCompare) = 0
                
                ' -------------------------------------------------------------------
                If bOK = True Then
                    
                    Let lngItemIndex = m_Row + 1
                    
                    Exit For
                
                End If
            
            Next m_Row
            
            ' -------------------------------------------------------------------
            
        Else
            
            Let sErrMsg = "There are no items to test for existence."
            
        End If
        
    End If

    ' -------------------------------------------------------------------

ErrResume:

    Let Exists = bOK
    Let pItemIndex = lngItemIndex
    Let pErrMsg = sErrMsg
    
    ' -------------------------------------------------------------------

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.Exists", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... Exists: Boolean.

' Function:        FindClosestItem
' Returns:         Long.
' Description:     Returns the position of the first item partially matching the search string.

Public Function FindClosestItem(ByVal pTheSearchText As String, Optional ByRef pTheFoundText As String = vbNullString, Optional ByVal pArrayIsSorted As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute FindClosestItem.VB_Description = "Returns the position of the first item partially matching the search string."

'... Parameters.
'    V__ pTheSearchText: String      ' ... The String to find.
'    RO_ pTheFoundText: String       ' ... The whole text of the matching item if found.
'    VO_ pArrayIsSorted: Boolean     ' ... If you know the array is sorted asc. pass this as true and a quicker binary search will be executed, else definitely pass false or omit.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim lngReturn As Long                ' ... a return value to this function.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim sTheItem As String
Dim sTheSearch As String
Dim sTheFoundText As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If pArrayIsSorted = False Then
    
        Let lngReturn = -1
        
        Let sTheSearch = UCase$(pTheSearchText) & "*"
        
        If pTestArray = True Then
        
            For m_Row = 1 To m_Count
                        
                Let sTheItem = UCase$(m_StringArray(m_Row - 1))
                
                If sTheItem Like sTheSearch Then
                    
                    Let lngReturn = m_Row - 1
                    
                    Let sTheFoundText = m_StringArray(m_Row - 1)
                    
                    Exit For
                    
                End If
            
            Next m_Row
        
        End If
    
    Else
        
        ' ... use the sorted array search to return index and full text of item.
        Let lngReturn = pAutoCompleteText(pTheSearchText, sTheFoundText, True, True)
                
    End If
    
ErrResume:

    Let pTheFoundText = sTheFoundText
    Let pErrMsg = sErrMsg
    Let FindClosestItem = lngReturn


Exit Function
ErrHan:

    Let lngReturn = -1
    Let sErrMsg = Err.Description
    
    Debug.Print "StringArray.FindClosestItem", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... FindClosestItem: Long.
'
'' Function:        FromArray
'' Returns:         Boolean.
'' Description:     Copies an Array String() to the buffer.
'
'Public Function FromArray(ByRef pTheStringArray() As String, Optional ByRef pErrMsg As String = vbNullString) As Boolean
'
''... Parameters.
''    R_A pTheStringArray: String     ' ... The array string() to copy.
''    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.
'
'' -------------------------------------------------------------------
'
'' Note:     Warning, this ignores Duplicates testing and clears the
''           Duplicates StringBuilder so existence testing would then need
''           to use either the Loop or the Instr(Join()) version in AddItemString.
'
'' -------------------------------------------------------------------
'
'' BIGGER NOTE:
''           This function is bolox in vb5, just comment it out.
'' -------------------------------------------------------------------
'
'Dim bOK As Boolean      ' ... Returns success or failure of this method.
'Dim sErrMsg As String   ' ... Returns an error message trapped / generated here-in.
'
'' -------------------------------------------------------------------
'
'    On Error GoTo ErrHan:
'
'    Call Clear
'
'    Let m_StringArray = pTheStringArray
'
'    Let bOK = pTestArray
'
'    If bOK = True Then
'
'        Let m_Count = UBound(m_StringArray) + 1
'        ReDim Preserve m_StringArray(m_Count)
'
'    End If
'
'ErrResume:
'
'    Let pErrMsg = sErrMsg
'    Let FromArray = bOK
'
'Exit Function
'ErrHan:
'
'    Let sErrMsg = Err.Description
'    Let bOK = False
'    Debug.Print "StringArray.FromArray", Err.Number, Err.Description
'    Resume ErrResume:
'
'
'End Function ' ... FromArray: Boolean.

' Function:        FromFile
' Returns:         Boolean.
' Description:     Reads the contents of a text file and converts it to the instance data.

Public Function FromFile(ByVal pTheFileName As String, Optional ByVal pTheDelimiter As String = ",", Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute FromFile.VB_Description = "Reads the contents of a text file and converts it to the instance data."

'... Parameters.
'    V__ pTheFileName: String        ' ... The Full Path Name and Location of the FIle to Read.
'    VO_ pTheDelimiter: String       ' ... The Delimiter to use to shred the data content.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim strFileText As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    
    Let strFileText = pReadTextFile(pTheFileName, bOK, sErrMsg)
    
    ' -------------------------------------------------------------------
    
    If bOK = True Then
        
        ' ... trim items, no empty items, clear existing items first.
        Let bOK = FromString(strFileText, pTheDelimiter, True, False, True, sErrMsg)
        
    Else
    
        ' ... update sErrMSg explaining file not read.
        Let sErrMsg = "FromFile: Did not Read the File: " & pTheFileName & vbCrLf & sErrMsg
    
    End If
    
ErrResume:
    
    ' -------------------------------------------------------------------
    
    Let pErrMsg = sErrMsg
    Let FromFile = bOK


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.FromFile", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... FromFile: Boolean.

' Function:        FromString
' Returns:         Boolean.
' Description:     Converts a String into a StringArray given an Optional Delimiter (default = comma).

Public Function FromString(ByVal pTheString As String, Optional ByVal pTheDelimiter As String = ",", Optional ByVal pTrimItems As Boolean = False, Optional ByVal pAllowEmptyItems As Boolean = True, Optional ByVal pClearExistingData As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute FromString.VB_Description = "Converts a String into a StringArray given an Optional Delimiter (default = comma)."

'... Parameters.
'    V__ pTheString: String          ' ... The String to Convert into the internal Array.
'    VO_ pTheDelimiter: String       ' ... The Delimiter to use to shred the data content.
'    VO_ pTrimItems: Boolean         ' ... Optional instruction to Trim Spaces from Items to be added to array (deault = False).
'    VO_ pAllowEmptyItems: Boolean   ' ... Opional instruction to allow empty items to be added to the array (default = True).
'    VO_ pClearExistingData: Boolean ' ... Optional instruction to clear existing array first before adding anything (default = False).
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim xCount As Long
Dim sTheStringData() As String
Dim sTheCurrentString As String

' -------------------------------------------------------------------

'   ??? ... does prefixing the vba functions help any
'       ... beyond intellisense ???

    On Error GoTo ErrHan:
        
    Let m_bCompacted = False
    
    If pClearExistingData = True Then
    
        Call Clear
    
    End If

    If Len(Trim$(pTheString)) > 0 Then
        
        ' -------------------------------------------------------------------
        If Len(pTheDelimiter) > 0 Then
            
            ' -------------------------------------------------------------------
            ' ... split only accepts one char so if splitting on
            ' ... vbCrLf the replace vbCrLf with just vbCr.
'            If pTheDelimiter = vbCrLf Then
'                Let pTheString = pReplace(pTheString, vbCrLf, vbCr)
'                Let pTheDelimiter = vbCr
'            End If
            
            ' -------------------------------------------------------------------
'            Let sTheStringData = VBA.Split(pTheString, pTheDelimiter)
            pSplitB04 pTheString, sTheStringData, pTheDelimiter
            
'            If pClearExistingData = True Then
                
'                ' ... if over-writing existing data then
'                ' ... take short-cut with split derived array.
'                ' ... Note: FromArray does not check for
'                ' ... uniqueness so things could easily go wrong
'                ' ... after this if DuplicatesAllowed = False
'                ' ... when new items are added because the string builder buffer
'                ' ... will not have been created. see FromArray.
'
'                Let bOK = FromArray(sTheStringData, sErrMsg) ' V3 comment
'
'                If bOK = True Then
'                    m_TheString = pTheString
'                    m_TheBufferCapacity = Len(m_TheString)
'                    m_TheTextLength = m_TheBufferCapacity
'                    GoTo ResumeExit:
'
'                End If
                
'            End If
            
            ' -------------------------------------------------------------------
            Let xCount = UBound(sTheStringData)
            
            ' -------------------------------------------------------------------
            For m_Row = 0 To xCount
                
                ' -------------------------------------------------------------------
                Let sTheCurrentString = sTheStringData(m_Row)
                
                ' -------------------------------------------------------------------
                If pTrimItems = True Then
                
                    Let sTheCurrentString = Trim$(sTheCurrentString)
                
                End If
                
                ' -------------------------------------------------------------------
                If Len(sTheCurrentString) > 0 Then
                    
                    ' -------------------------------------------------------------------
                    Call AddItemString(sTheCurrentString)
                    
                Else
                    
                    ' -------------------------------------------------------------------
                    If pAllowEmptyItems = True Then
                    
                        Call AddItemString(sTheCurrentString)
                    
                    End If
                    
                End If
                
                ' -------------------------------------------------------------------
            
            Next m_Row
            
            ' -------------------------------------------------------------------
            Let bOK = True
        
        Else
            
            ' -------------------------------------------------------------------
            Let sErrMsg = "A Delimiter is required to parse the string."
        
        End If
    
    Else
        
        ' -------------------------------------------------------------------
        Let sErrMsg = "There is no string data to parse."
    
    End If


ResumeExit:

ErrResume:
    
    ' -------------------------------------------------------------------
    Let pErrMsg = sErrMsg
    Let FromString = bOK


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.FromString", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... FromString: Boolean.

' Function:        IndexExists
' Returns:         Boolean.
' Description:     Describes whether a given index exists within the one based array (True) or Not (False).

Public Function IndexExists(ByVal pItemIndex As Long, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute IndexExists.VB_Description = "Describes whether a given index exists within the one based array (True) or Not (False)."

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index to query.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
' Note:     Depends on m_Count being correct.
'           following buffer pre-allocation.
'           Otherwise might be tempted to just
'           reference the item directly and catch an
'           out of bounds error to return false if
'           no such index exists.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If pTestArray = True Then
        
        If pItemIndex > 0 Then
                        
'''            Call Compact ' added this blindly after buffer pre-allocation implementation
                            ' and it messed everything up, left it commented as reminder.
            
'''            If UBound(m_StringArray) >= pItemIndex - 1 Then
            If m_Count - 1 >= pItemIndex - 1 Then
            
                Let bOK = True
            
            End If
        
        End If
                
        If bOK = False Then
            
            Let sErrMsg = "The index provided is not within the valid range [ 1 to " & CStr(m_Count) & " ] ."
            
        End If
                
    Else
    
        Let sErrMsg = "There are no items to test for index existence."
        
    End If

ErrResume:

    Let pErrMsg = sErrMsg
    Let IndexExists = bOK


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.IndexExists", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... IndexExists: Boolean.

' Property Get:    Item
' Type:            String.
' Description:     Sets / Returns the Item at the given Index (default property).

Public Property Get Item(ByVal pItemIndex As Long) As String
Attribute Item.VB_Description = "Checks valid array and index to return item in array."
Attribute Item.VB_UserMemId = 0

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index of the item to return (index is one based).

' -------------------------------------------------------------------

Dim bOK As Boolean
Dim sErrMsg As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    Let bOK = IndexExists(pItemIndex, sErrMsg)

    If bOK = True Then

        Let Item = m_StringArray(pItemIndex - 1)

    End If

    If bOK = False Then

        Let sErrMsg = "StringArray.Item: Unable to return Item Text: " & vbCrLf & sErrMsg
        Debug.Print sErrMsg
        
    End If

Exit Property
ErrHan:

    Debug.Print "StringArray.Item", Err.Number, Err.Description

End Property ' ... Item: String.

' Property Let:    Item

Public Property Let Item(ByVal pItemIndex As Long, ByVal pNewValue As String)

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index of the item to return (index is one based).
'    V__ pNewValue: String           ' ... the value to this property.
' -------------------------------------------------------------------

Dim bOK As Boolean
Dim sErrMsg As String

' -------------------------------------------------------------------
' Note:     No unique test when DuplicatesAllowed = False or True,
'           so value edited could mess things up later.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let bOK = IndexExists(pItemIndex, sErrMsg)

    If bOK = True Then

        Let m_StringArray(pItemIndex - 1) = pNewValue

    End If

    If bOK = False Then

        Let sErrMsg = "StringArray.Item: Unable to return Item Text: " & vbCrLf & sErrMsg
        Debug.Print sErrMsg
        
    End If

Exit Property
ErrHan:

    Debug.Print "StringArray.Item", Err.Number, Err.Description

End Property ' ... Item: String.

' Function:        ItemAsDateValue
' Returns:         Date.
' Description:     Returns a date conversion of a string by index from the string() buffer.

Public Function ItemAsDateValue(ByVal pItemIndex As Long, Optional ByRef pItemFound As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Date
Attribute ItemAsDateValue.VB_Description = "Returns a CVDate conversion of a string by index from the string() buffer."

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index to query.
'    RO_ pItemFound: Boolean         ' ... Returns a Boolean describing whether the item was found and therefore whether a return of zero represents the value or not.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim datReturn As Date                ' ... a return value to this function.
Dim bOK As Boolean
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim strItem As String

' -------------------------------------------------------------------

' Note:     Using VB's Date functions that are locale aware so different
'           date formats may easily be disregarded or return the wrong date.
'           Too much extra work to truely confirm date string is date value.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    Let datReturn = -1
    
    Let pItemFound = False
    
    Let strItem = pItem(pItemIndex, bOK, sErrMsg)

    If bOK = True Then
        
        Let pItemFound = True
        
        If IsDate(strItem) Then
            
            Let datReturn = CVDate(strItem)
        
        Else
            
            ' ... explain that the value was not converted to a date because it is not recognised as being so.
            Let sErrMsg = "The Value [" & strItem & "] at position [ " & pItemIndex & " ] was not recognised as being a Date."
        
        End If
    
    End If
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let ItemAsDateValue = datReturn


Exit Function
ErrHan:

    Let datReturn = -1
    Let sErrMsg = Err.Description
    Debug.Print "StringArray.ItemAsDateValue", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... ItemAsDateValue: Double.

' Function:        ItemAsNumberValue
' Returns:         Double.
' Description:     Returns a Val conversion of a string by index from the string() buffer.

Public Function ItemAsNumberValue(ByVal pItemIndex As Long, Optional ByRef pItemFound As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Double
Attribute ItemAsNumberValue.VB_Description = "Returns a Val conversion of a string by index from the string() buffer."

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index to query.
'    RO_ pItemFound: Boolean         ' ... Returns a Boolean describing whether the item was found and therefore whether a return of zero represents the value or not.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim dblReturn As Double              ' ... a return value to this function.
Dim bOK As Boolean
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim strItem As String

' -------------------------------------------------------------------
'   Note:       As ItemAsDateValue this method could return the wrong value
'               or conclude an invalid number if number read is in a different
'               format to locale settings e.g. £25.50 on UK locale would be ok
'               but $25.50 would not be OK and vice-versa.
'               The same goes for general numbers where the decimal and thousandths
'               spearators do not concur with locale settings.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let pItemFound = False
    
    Let strItem = pItem(pItemIndex, bOK, sErrMsg)

    If bOK = True Then
        
        Let pItemFound = True
        
        If IsNumeric(strItem) Then
            ' ... BUG.      v3, had Let strItem = CDbl(strItem) and updated.
            Let dblReturn = CDbl(strItem) ' ... CDbl rather than Val to help with Locale aware Numeric/Currency values.
        
        Else
            
            ' ... explain that the value was not converted to a number because it is not recognised as being so.
            Let sErrMsg = "The Value [" & strItem & "] at position [ " & pItemIndex & " ] was not recognised as being Numeric."
        
        End If
    
    End If
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let ItemAsNumberValue = dblReturn


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Debug.Print "StringArray.ItemAsNumberValue", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... ItemAsNumberValue: Double.

' Function:        ItemAsStringArray
' Returns:         StringArray.
' Description:     Returns a string converted into a StringArray via a Delimiter that defaults to a colon ( : ) if omitted.

Public Function ItemAsStringArray(ByVal pItemIndex As Long, Optional ByVal pTheDelimiter As String = ":", Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As StringArray
Attribute ItemAsStringArray.VB_Description = "Returns a string converted into a StringArray via a Delimiter that defaults to a colon ( : ) if omitted."

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index to query.
'    VO_ pTheDelimiter: String       ' ... The Delimiter to use to shred the data content.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim objReturn As StringArray         ' ... a return value to this function.
Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim sTmpItem As String
Dim bIndexed As Boolean

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... create a new string array regardless, even if it will have no data,
    ' ... personal preference rather than return nothing.
    
    Set objReturn = New StringArray
    
    ' -------------------------------------------------------------------
    
    Let bIndexed = IndexExists(pItemIndex, sErrMsg)
        
    ' -------------------------------------------------------------------
    
    If bIndexed = True Then
    
        ' ... use pKnownItem as have already tested index above.
        Let sTmpItem = pKnownItem(pItemIndex)
        
        ' -------------------------------------------------------------------
        ' ... re-use FromString on the Item.
        Let bOK = objReturn.FromString(sTmpItem, pTheDelimiter, , , , sErrMsg)
    
    Else
    
        Let sErrMsg = "There are no items available to convert to a StringArray."
        
    End If
    
ErrResume:
    
    ' -------------------------------------------------------------------
    
    Let pErrMsg = sErrMsg
    Let pOK = bOK
    Set ItemAsStringArray = objReturn

    On Error Resume Next
    
    Set objReturn = Nothing

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.ItemAsStringArray", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... ItemAsStringArray: StringArray.

'**************************************
' Name: StringArrayToString
' Description:This code converts a stringarray into a string by first converting it into a byte array.
' I wrote this code after finding out that the new version of VB (vb7, codename .NET) does noet support varPtr, objPtr e.g. pointers.<BR><BR>
' Before i did this using the API call RtlMoveMemory
' By: G. van den Hoven
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
' Side Effects:None
' This code is copyrighted and has limited warranties.
' Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.13008/lngWId.1/qx/vb/scripts/ShowCode.htm for details.
'
' use http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=13008&lngWId=1
'**************************************

'Public Function StringArrayToString(pbIn() As String) As String
'    Dim bTemp() As Byte
'    Dim lSize As Long
'    Dim lNow As Long
'    Dim lTotal As Long
'    Dim lNowArray As Long
'    Dim tTemp As String
'    Dim lNow2 As Long
'    Dim lTotal2 As Long
'
'    '-----
'    ' Calculate size of inputarray
'    '-----
'    lSize = 0
'    lTotal = UBound(pbIn)
'
'
'    For lNow = 0 To lTotal
'        lSize = lSize + Len(pbIn(lNow))
'    Next
'
'    '-----
'    ' Create byte array which is big
'    ' enough to hold all the bytes
'    '-----
'    ReDim bTemp(0 To lSize)
'
'    '-----
'    ' Convert the string array to a byte arr
'    '     ay
'    '-----
'    lNow = 0
'    lNowArray = 0
'
'
'    While lNow <> lSize
'        tTemp = pbIn(lNowArray)
'        lTotal2 = Len(tTemp)
'        '-----
'        ' Loop through the temp string
'        ' and place the byte character
'        ' in the correct position
'        ' Mid$(...) is faster then Mid(...)
'        '-----
'
'
'        For lNow2 = 0 To lTotal2 - 1
'            bTemp(lNow + lNow2) = Asc(Mid$(tTemp, lNow2 + 1, 1))
'        Next
'        lNow = lNow + lTotal2
'        lNowArray = lNowArray + 1
'    Wend
'
'    '-----
'    ' Convert byte array to string
'    '-----
'    tTemp = StrConv(bTemp(), vbUnicode)
'    StringArrayToString = tTemp
'End Function



' -------------------------------------------------------------------
' Function:         pAutoCompleteText
' Returns:          Long.
' Description:      Searches the array for an item (partial or exact), if found, index returned along with full string (for Partial matching).
' -------------------------------------------------------------------

Private Function pAutoCompleteText(ByVal pTheSearchText As String, ByRef pTheFoundText As String, Optional ByVal pPatialMatch As Boolean = True, Optional ByVal pGuardPatternCharacters As Boolean = False) As Long

' Parameters:
'   ... V__ pTheSearchText: String.                     ' ... .
'   ... RO_ pTheFoundText: String.                      ' ... .
'   ,,, VO_ pPatialMatch: Boolean.                      ' ... .
'   ,,, VO_ pGuardPatternCharacters: Boolean.           ' ... .
' -------------------------------------------------------------------

Dim strText As String
Dim intLength As Integer
Dim lngIndex As Long
Dim sCurrentText As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let pAutoCompleteText = -1  ' ... default return.
    
    ' -------------------------------------------------------------------
    Let strText = pTheSearchText
    
    Let sCurrentText = strText
    
    Let intLength = Len(strText)
    
    ' -------------------------------------------------------------------
    
    ' ... Don't search if the length is zero.
    If intLength > 0 Then
        
        If pGuardPatternCharacters = True Then
            
            ' -------------------------------------------------------------------
            ' ... if allowing all characters and comparing then need to
            ' ... safe guard against pattern characters used with vb's Like function.
            
            Let strText = pReplace(strText, "#", "[#]")
            Let strText = pReplace(strText, "*", "[*]")
            Let strText = pReplace(strText, "?", "[?]")
        
        End If
        
        Let lngIndex = pBinarySearch(strText, pPatialMatch, pGuardPatternCharacters)
        
        If lngIndex >= 0 Then
            Let pTheFoundText = m_StringArray(lngIndex)
        End If
        
        Let pAutoCompleteText = lngIndex
        
    End If

ResumeError:

Exit Function

ErrHan:

    Debug.Print "StringArray.pAutoCompleteText.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:


End Function ' ... pAutoCompleteText: Long

' Function:         pBinarySearch
' Returns:          Long.
' Description:      Searches a one-dimensional string array for the search string and returns its index if found or -1.

Private Function pBinarySearch(ByVal pSearchString As String, Optional ByVal pPartialMatch As Boolean, Optional ByVal pGuardPatternCharacters As Boolean = True) As Long

' Parameters:
'   ... V__ pSearchString: String.              ' ... The string to find.
'   ... VO_ pPartialMatch: Boolean.             ' ... True means partial match else exact match.
'   ... VO_ pGuardPatternCharacters: Boolean.   ' ... True means guard against pattern matching chars (#, *, ?) in search string.
' -------------------------------------------------------------------

Dim lngLow As Long
Dim lngUpper As Long
Dim lngPos As Long

Dim sSearch As String
Dim sPartial As String

' -------------------------------------------------------------------
' Note:     Thanks to Dan Fox, Pure Visual Basic, for the code base here
'           I added the pattern character guard bits to help with using the LIKE function
'           and reduced this to a one dimensional string array rather than the multi-dimmension
'           variant array binary search.
'           IMPORTANT.  Only use this on a sorted string array (asc) else you may not find
'           what you know is there.
'           ALSO, this will not work on numbers as strings if they are not formatted correctly
'           e.g. instead of 10, 100 use 0000010, 0000100 to get correct sorting.
'           This is class is dedicated to an array of strings rather than numbers.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let sSearch = CStr(pSearchString)
    
    If pGuardPatternCharacters = True Then
        
        Let sSearch = pReplace(sSearch, "[#]", "#")
        Let sSearch = pReplace(sSearch, "[*]", "*")
        Let sSearch = pReplace(sSearch, "[?]", "?")
    
    End If
        
    Let sPartial = sSearch & "*"

    Let lngLow = 0 ' LBound(m_StringArray)
    Let lngUpper = m_Count - 1 ' UBound(m_StringArray)

    Do While True
        
        ' ... Divide the array to search.
        Let lngPos = (lngLow + lngUpper) / 2
        
        If pPartialMatch Then
            
            If m_StringArray(lngPos) Like sPartial Then
                Let pBinarySearch = lngPos
                Exit Function
            End If
        
        Else
            
            If m_StringArray(lngPos) = sSearch Then
                Let pBinarySearch = lngPos
                Exit Function
            End If
            
        End If
            
            ' ... Check to see if its the last value to be checked.
            If lngUpper = lngLow + 1 Then
                Let lngLow = lngUpper
            Else
                ' ... If we get to the lowest position then it's not there.
                If lngPos = lngLow Then
                    Let pBinarySearch = -1
                    Exit Function
                Else
                    ' ... Determine whether to look in the upper or
                    ' ... lower half of the array.
                    If m_StringArray(lngPos) > sSearch Then
                        Let lngUpper = lngPos
                    Else
                        Let lngLow = lngPos
                    End If
                    
                End If
                
            End If
            
    Loop

Exit Function

ErrHan:

    Let pBinarySearch = -1
    Debug.Print "StringArray.pBinarySearch.Error: " & Err.Number & "; " & Err.Description

End Function ' ... pBinarySearch: Long

' Function:        pCheckFileName
' Returns:         Boolean.
' Description:     Checks a file name for invalid characters.

Private Function pCheckFileName(ByVal pTheFileName As String, Optional ByRef pErrMsg As String = vbNullString) As Boolean

'... Parameters.
'    R__ pTheFileName: String        ' ... The Full Path Name and Location of the FIle to check.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------
Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sFilePath As String
Dim sFileName As String
Dim sFileExt As String
Dim sTheChar As String
Dim sInvalidChars As String
Dim lngLoop As Long
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... Thanks to Fransesco Balena, VB2TheMax
    ' ... Split the full file name into path, name and extension.
    ' -------------------------------------------------------------------
    sFileName = pTheFileName
    ' -------------------------------------------------------------------
    For lngLoop = Len(sFileName) To 1 Step -1
        If Mid$(sFileName, lngLoop, 1) = "." Then
            sFileExt = Mid$(sFileName, lngLoop + 1)
            sFileName = Left$(sFileName, lngLoop - 1)
        ElseIf InStr(":\", Mid$(sFileName, lngLoop, 1)) Then
            sFilePath = Left$(sFileName, lngLoop)
            If Right$(sFilePath, 1) = "\" Then sFilePath = Left$(sFilePath, lngLoop - 1)
            sFileName = Mid$(sFileName, lngLoop + 1)
            Exit For
        End If
    Next lngLoop
    ' -------------------------------------------------------------------
        
    ' -------------------------------------------------------------------
    ' ... having extracted the file name make sure it is free of invalid characters.
    Let sInvalidChars = c_def_InvalidFileNameChars & Chr$(34) ' "`!$%^&*()-=+[]{}'#@~;,.<>/?\| " & Chr$(34)
    ' -------------------------------------------------------------------
    For lngLoop = 1 To Len(sFileName)
        Let sTheChar = Mid$(sFileName, lngLoop, 1)
        If InStr(sInvalidChars, sTheChar) Then
            ' -------------------------------------------------------------------
            ' ... easier to throw an error than anything else if invalid char. found.
            Err.Raise vbObjectError + 1001, , "The File Name [ " & sFileName & " ] has the following invalid character [ " & sTheChar & " ]."
            ' ... an alternative is to generate error msg and exit function here
            ' ... but that would be wasting the error handler.
        End If
    Next lngLoop
    ' -------------------------------------------------------------------
    ' ... could also check valid directory and prompt to create.
    
    ' -------------------------------------------------------------------
    
    Let bOK = True

ErrResume:

    Let pCheckFileName = bOK


Exit Function
ErrHan:

    Let pErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringWorker.pCheckFileName", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pCheckFileName: Boolean.

' Function:        pExists
' Returns:         Boolean.
' Description:     Checks to see if a value exists in the current data.

Private Function pExists(ByVal pItemString As String, Optional ByVal pStartPos As Long = 1, Optional ByRef pItemIndex As Long = 0, Optional ByRef pErrMsg As String = vbNullString) As Boolean

'... Parameters.
'    V__ pItemString: String         ' ... The string to search for.
'    RO_ pItemIndex: Long            ' ... returns the one based index of the item if found else 0.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim lngItemIndex As Long
Dim lngStartPos As Long
Dim bIndexed As Boolean

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    
    Let bIndexed = IndexExists(pStartPos, sErrMsg)
    
    ' -------------------------------------------------------------------
    
    If bIndexed = True Then
        
        ' -------------------------------------------------------------------
        
        Let lngStartPos = pStartPos - 1
       
        ' -------------------------------------------------------------------
        
        For m_Row = lngStartPos To m_Count - 1 ' UBound(m_StringArray)
        
            If m_StringArray(m_Row) = pItemString Then
                
                Let lngItemIndex = m_Row + 1
                
                Let bOK = True
                
                Exit For
            
            End If
        
        Next m_Row
        
        ' -------------------------------------------------------------------
        
    Else
    
        Let sErrMsg = "There are no items to test for existence."
    
    End If
    
    ' -------------------------------------------------------------------

ErrResume:

    Let pItemIndex = lngItemIndex
    Let pErrMsg = sErrMsg
    Let pExists = bOK
    
    ' -------------------------------------------------------------------

Exit Function
ErrHan:
    
    Let lngItemIndex = 0
    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.pExists", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pExists: Boolean.

' Function:        pFileExists
' Returns:         Boolean.
' Description:     Tests for file existence, returns True if found else False.

Private Function pFileExists(ByVal pTheFileName As String, Optional ByRef pErrMsg As String = vbNullString) As Boolean

'... Parameters.
'    V__ pTheFileName: String        ' ... The Full Path Name and Location of the File to find.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let bOK = (Dir$(pTheFileName) <> "")
    
    If bOK = False Then
        
        Let sErrMsg = "The File [ " & pTheFileName & " ] could not be found."
    
    End If
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let pFileExists = bOK

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.pFileExists", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pFileExists: Boolean.

' Function:        pItem
' Returns:         String.
' Description:     Checks valid array and index to return item in array.

Private Function pItem(ByVal pItemIndex As Long, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As String

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index of the item to return (index is one based).
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim strReturn As String              ' ... a return value to this function.
Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    Let bOK = IndexExists(pItemIndex, sErrMsg)

    If bOK = True Then

        Let strReturn = m_StringArray(pItemIndex - 1)

    End If

    If bOK = False Then

        Let sErrMsg = "Unable to return Item Text: " & vbCrLf & sErrMsg

    End If

ErrResume:

    Let pErrMsg = sErrMsg
    Let pOK = bOK
    Let pItem = strReturn


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.pItem", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pItem: String.

' Function:         pKnownItem
' Returns:          String.
' Description:      Returns a string from the buffer with no testing, assumes use when index is known to exist.

Private Function pKnownItem(ByVal pItemIndex As Long) As String

' Parameters:
'    V__ pItemIndex: Long            ' ... The index of the item to return (index is one based).

' -------------------------------------------------------------------

' Notes:    only use this method when the index has been previously tested or is known to exist.

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let pKnownItem = m_StringArray(pItemIndex - 1)

Exit Function

ErrHan:

    Debug.Print "StringArray.pKnownItem.Error: " & Err.Number & "; " & Err.Description

End Function

' Note: QuickSort algorithms c/o Squirm and Optikon on XtremeVBTalk.
' http://www.xtremevbtalk.com/showthread.php?t=78889&highlight=Quick+Sort

Private Sub pQuickSortOnLongArray(ByRef lngArray() As Long, ByVal iLeftEnd As Long, ByVal iRightEnd As Long)

' -------------------------------------------------------------------
' Thank You Squirm.
' -------------------------------------------------------------------

Dim iLeftCur As Long
Dim iRightCur As Long
Dim iPivot As Long
Dim iTemp As Long

' -------------------------------------------------------------------
' Note:     Recursive Method.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If iLeftEnd >= iRightEnd Then Exit Sub
    
    iLeftCur = iLeftEnd
    iRightCur = iRightEnd + 1
    iPivot = lngArray(iLeftEnd)
    
    Do
        'Arrange values so that < pivot are on the left and > pivot are on the right
        
        'Find >= value on left side
        Do
            iLeftCur = iLeftCur + 1
        Loop While lngArray(iLeftCur) < iPivot

        'Find <= value on right side
        Do
            iRightCur = iRightCur - 1
        Loop While lngArray(iRightCur) > iPivot

        'No more swapping to do
        If iLeftCur >= iRightCur Then Exit Do
        
        'Swap
        iTemp = lngArray(iLeftCur)
        lngArray(iLeftCur) = lngArray(iRightCur)
        lngArray(iRightCur) = iTemp
    
    Loop
        
    'Call quicksort recursively on left and right subarrays
    lngArray(iLeftEnd) = lngArray(iRightCur)
    lngArray(iRightCur) = iPivot
    
    pQuickSortOnLongArray lngArray, iLeftEnd, iRightCur - 1
    pQuickSortOnLongArray lngArray, iRightCur + 1, iRightEnd


Exit Sub

ErrHan:

    Debug.Print "StringArray.pQuickSortOnLongArray.Error: " & Err.Number & "; " & Err.Description
    
End Sub

Private Sub pQuickSortOnStringArray(ByRef MyArray() As String, ByVal lngLBound As Long, ByVal lngUbound As Long)

' -------------------------------------------------------------------
'   Thank you Optikon.
' -------------------------------------------------------------------

Dim lngPivot As String
Dim k As Long
Dim p As Long
Dim lngTemp As String

' -------------------------------------------------------------------
' Note:     Recursive Method.
' -------------------------------------------------------------------
    
    On Error GoTo ErrHan:
    
    If lngLBound >= lngUbound Then
        Exit Sub
    End If
    
    k = lngLBound + 1
    
    'this if statement is not needed, but it speeds up the code slightly
    If lngUbound = k Then
'        If MyArray(lngLBound) > MyArray(lngUBound) Then
        ' ... putting in this test buggers potential performance
        ' ... but will have to do for now.
        If LCase$(MyArray(lngLBound)) > LCase$(MyArray(lngUbound)) Then ' ... v3
            'swap MyArray(lngLBound) and MyArray(lngUBound)
            lngTemp = MyArray(lngLBound)
            MyArray(lngLBound) = MyArray(lngUbound)
            MyArray(lngUbound) = lngTemp
        End If
        Exit Sub
    End If
    
    '*** Uncomment the following 4 lines and it will make it perform equally well on (almost) sorted data
    'p = lngLBound + ((lngUBound - lngLBound) \ 2)
    'lngTemp = MyArray(lngLBound)
    'MyArray(lngLBound) = MyArray(p)
    'MyArray(p) = lngTemp
    
'    lngPivot = MyArray(lngLBound)
    lngPivot = LCase$(MyArray(lngLBound))
    p = lngUbound
    
'    Do Until (MyArray(k) > lngPivot) Or (k >= lngUBound)
    Do Until (LCase$(MyArray(k)) > lngPivot) Or (k >= lngUbound)
        k = k + 1
    Loop
    
'    Do Until MyArray(p) <= lngPivot
    Do Until LCase$(MyArray(p)) <= lngPivot
        p = p - 1
    Loop
    
    Do While k < p
        'swap MyArray(k) and MyArray(p)
        lngTemp = MyArray(k)
        MyArray(k) = MyArray(p)
        MyArray(p) = lngTemp
        
        Do
            k = k + 1
'        Loop Until MyArray(k) > lngPivot
        Loop Until LCase$(MyArray(k)) > lngPivot
        
        Do
            p = p - 1
'        Loop Until MyArray(p) <= lngPivot
        Loop Until LCase$(MyArray(p)) <= lngPivot
    Loop
    
    'swap MyArray(p) and MyArray(lngLBound)
    lngTemp = MyArray(p)
    MyArray(p) = MyArray(lngLBound)
    MyArray(lngLBound) = lngTemp
    
    pQuickSortOnStringArray MyArray, lngLBound, p - 1
    pQuickSortOnStringArray MyArray, p + 1, lngUbound

Exit Sub

ErrHan:

    Debug.Print "StringArray.pQuickSortOnStringArray.Error: " & Err.Number & "; " & Err.Description

End Sub

' Function:        pReadTextFile
' Returns:         String.
' Description:     The contents of a text file as a string.

Private Function pReadTextFile(ByVal pTheFileName As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As String

'... Parameters.
'    V__ pTheFileName: String        ' ... The Full Path Name and Location of the FIle to Read.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim strReturn As String              ' ... a return value to this function.
Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
Dim bFileExists As Boolean

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Let bFileExists = pFileExists(pTheFileName, sErrMsg)
    
    If bFileExists = True Then
        
        ' -------------------------------------------------------------------
        ' ... get the next free file number.
        iFileNumber = FreeFile()
        ' -------------------------------------------------------------------
        Open pTheFileName For Binary Access Read As #iFileNumber
        ' -------------------------------------------------------------------
        ' ... if execution flow got here, the file has been open without error.
        bFileIsOpen = True
        
        strReturn = Space$(LOF(iFileNumber))
        
        ' -------------------------------------------------------------------
        ' ... read the entire contents in one single operation.
        Get #iFileNumber, , strReturn
        ' -------------------------------------------------------------------
        ' ... Print Statement from pWriteTextFile adds vbCRLF to the text written.
        If Len(strReturn) > 2 Then
            If Right$(strReturn, 2) = vbCrLf Then
                strReturn = Left$(strReturn, Len(strReturn) - 2)
            End If
        End If
        ' -------------------------------------------------------------------
        pReadTextFile = strReturn    ' ... return the string from the text file.
        bOK = True                   ' ... return success.
        ' -------------------------------------------------------------------

    End If
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let pOK = bOK
    Let pReadTextFile = strReturn

    On Error Resume Next    ' ... don't fall over because we couldn't release the file handle.

    If bFileIsOpen = True Then
        If Err.Number <> 0 Then Err.Clear
        Close #iFileNumber
        If Err.Number <> 0 Then
            Debug.Print "StringArray.pReadTextFile (Close File)", Err.Number, Err.Description
            Err.Clear
        End If
    End If

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.pReadTextFile", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pReadTextFile: String.

' Property Get:    PreAllocationSize
' Type:            Long.
' Description:     Sets / Returns the buffer Pre-Allocation Size for ReDim Preserve Statements.

Public Property Get PreAllocationSize() As Long
Attribute PreAllocationSize.VB_Description = "Sets / Returns the buffer Pre-Allocation Size for ReDim Preserve Statements."

    On Error GoTo ErrHan:

    Let PreAllocationSize = m_PreAllocationSize

Exit Property
ErrHan:

    Debug.Print "StringArray.PreAllocationSize", Err.Number, Err.Description

End Property ' ... PreAllocationSize: Long.

' Property Let:    PreAllocationSize

Public Property Let PreAllocationSize(ByVal pNewValue As Long)

    On Error GoTo ErrHan:
    
    If pNewValue > 0 Then
    
        Let m_PreAllocationSize = pNewValue
'        Call PropertyChanged("PreAllocationSize")
        
    End If

Exit Property
ErrHan:

    Debug.Print "StringArray.PreAllocationSize", Err.Number, Err.Description

End Property ' ... PreAllocationSize: Long.

Private Function pReplace(pExpression As String, pFind As String, ppReplace As String, Optional pStart As Long = 1, Optional pCount As Long = -1, Optional pCompare As VbCompareMethod = vbBinaryCompare) As String

Dim sTmp As String
    
    sTmp = pExpression
    
    If pCompare = vbTextCompare Then
        sTmp = LCase$(sTmp)
        pFind = LCase$(pFind)
    End If
    
    pReplaceChars sTmp, pFind, ppReplace, pStart, pCount
    
    pReplace = sTmp
    
    sTmp = vbNullString
    
End Function

' Sub:             pReplace
' Description:     Replace a string within another string with a different string :).

Private Sub pReplaceChars(pTheString As String, ByRef pSearchFor As String, ByRef pReplaceWith As String, Optional pStart As Long = 1, Optional ByRef pCountReplaced As Long = -1, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)

'... Parameters.
'    R__ pSearchFor: String          ' ... The character/s to search for within the user text.
'    R__ pReplaceWith: String        ' ... The replacement string value.
'    RO_ pStart: Long                ' ... The position from which to begin the replacing.
'    RO_ pCountReplaced: Long        ' ... Returns the number ofsearch string replacements.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.


'   ... With Thanks.
'   ... submitted 18-Dec-2000 by Jost Schwider to VBSpeed.
'   ... http://www.xbeat.net/vbspeed/c_Replace.htm#Replace09

Dim TextLen As Long
Dim OldLen As Long
Dim NewLen As Long
Dim ReadPos As Long
Dim WritePos As Long
Dim CopyLen As Long
Dim Buffer As String
Dim BufferLen As Long
Dim BufferPosNew As Long
Dim BufferPosNext As Long
Dim Search As String

'   Note:   Using Byte versions of string functions so string must be pure AINSI.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Search = pTheString
    pCountReplaced = 0
    
    If Len(pSearchFor) = 0 Then Exit Sub
    
    
    If pStart < 2 Then
        pStart = InStrB(Search, pSearchFor)
    Else
        pStart = InStrB(pStart + pStart - 1, Search, pSearchFor)
    End If
  
    If pStart Then
    
        ReadPos = 1
        WritePos = 1

        Buffer = Space$(Len(Search))

        OldLen = LenB(pSearchFor)
        NewLen = LenB(pReplaceWith)
        
        Select Case NewLen
        
            Case OldLen ' ... query and replace are same size, no effect on string length.
            
                Buffer = Search
                
                Do Until pStart = 0
                    MidB$(Buffer, pStart) = pReplaceWith
                    pStart = InStrB(pStart + OldLen, Buffer, pSearchFor)
                    pCountReplaced = pCountReplaced + 1
                Loop
            
            Case Is < OldLen ' ... replace is smaller than query string, string length will shrink.
            
                TextLen = LenB(Search)
                
                If NewLen Then
                
                    Do Until pStart = 0
                        CopyLen = pStart - ReadPos
                        If CopyLen Then
                            BufferPosNew = WritePos + CopyLen
                            MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                            MidB$(Buffer, BufferPosNew) = pReplaceWith
                            WritePos = BufferPosNew + NewLen
                        Else
                            MidB$(Buffer, WritePos) = pReplaceWith
                            WritePos = WritePos + NewLen
                        End If
                        ReadPos = pStart + OldLen
                        pStart = InStrB(ReadPos, Search, pSearchFor)
                        pCountReplaced = pCountReplaced + 1
                    Loop
                
                Else    ' ... replace is empty.
                
                    Do Until pStart = 0
                        CopyLen = pStart - ReadPos
                        If CopyLen Then
                            MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                            WritePos = WritePos + CopyLen
                        End If
                        ReadPos = pStart + OldLen
                        pStart = InStrB(ReadPos, Search, pSearchFor)
                        pCountReplaced = pCountReplaced + 1
                    Loop
                
                End If
      
                If ReadPos > TextLen Then
                    Buffer = LeftB$(Buffer, WritePos - 1)
                Else
                    MidB$(Buffer, WritePos) = MidB$(Search, ReadPos)
                    Buffer = LeftB$(Buffer, WritePos + TextLen - ReadPos)
                End If
                
            Case Else   ' ... replace is larger than query, string length will grow.
                
                TextLen = LenB(Search)
                
                BufferPosNew = TextLen + NewLen
                
                If BufferPosNew > BufferLen Then
                    Buffer = Space$(BufferPosNew)
                    BufferLen = LenB(Buffer)
                End If
                      
                Do Until pStart = 0
                    
                    CopyLen = pStart - ReadPos
                    
                    If CopyLen Then
                        
                        BufferPosNew = WritePos + CopyLen
                        BufferPosNext = BufferPosNew + NewLen
                        
                        If BufferPosNext > BufferLen Then
                            ' ... grow the buffer.
                            Buffer = Buffer & Space$(BufferPosNext)
                            BufferLen = LenB(Buffer)
                        End If
                        
                        MidB$(Buffer, WritePos) = MidB$(Search, ReadPos, CopyLen)
                        MidB$(Buffer, BufferPosNew) = pReplaceWith
                    
                    Else
                        
                        BufferPosNext = WritePos + NewLen
                        
                        If BufferPosNext > BufferLen Then
                            ' ... grow the buffer.
                            Buffer = Buffer & Space$(BufferPosNext)
                            BufferLen = LenB(Buffer)
                        End If
                        
                        MidB$(Buffer, WritePos) = pReplaceWith
                    
                    End If
                    
                    WritePos = BufferPosNext
                    ReadPos = pStart + OldLen
                    pStart = InStrB(ReadPos, Search, pSearchFor)
                    pCountReplaced = pCountReplaced + 1
                
                Loop
      
                If ReadPos > TextLen Then
                    Buffer = LeftB$(Buffer, WritePos - 1)
                Else
                    BufferPosNext = WritePos + TextLen - ReadPos
                    If BufferPosNext < BufferLen Then
                        MidB$(Buffer, WritePos) = MidB$(Search, ReadPos)
                        Buffer = LeftB$(Buffer, BufferPosNext)
                    Else
                        Buffer = LeftB$(Buffer, WritePos - 1) & MidB$(Search, ReadPos)
                    End If
                End If
            
        End Select
        
        Let pTheString = Buffer
        
    End If
    
    pOK = True ' ... executed, no errors, doesn't mean anything was replaced though.
    
ResumeError:

Exit Sub

ErrHan:
    
    pOK = False
    Let pErrMsg = Err.Description
    
    Debug.Print "StringArray.Replace.Error: " & Err.Description

    Resume ResumeError:

End Sub ' ... Replace.

Private Sub pSortNumber(ByRef pArray() As Long, ByVal lngLBound As Long, ByVal lngUbound As Long)

' -------------------------------------------------------------------
' Thank You Squirm.
' -------------------------------------------------------------------

Dim iTemp As Long
Dim iOuter As Long
Dim iMax As Long
    
    On Error GoTo ErrHan:
    
    If (lngUbound - lngLBound) Then
        
        'Move the largest value to the rightmost position, otherwise
        'we need to check that iLeftCur does not exceed the bounds of the
        'array on EVERY pass (time consuming)
        
        For iOuter = lngLBound To lngUbound
        
            If pArray(iOuter) > pArray(iMax) Then
                Let iMax = iOuter
            End If
            
        Next iOuter
        
        Let iTemp = pArray(iMax)
        Let pArray(iMax) = pArray(lngUbound)
        Let pArray(lngUbound) = iTemp
        
        Call pQuickSortOnLongArray(pArray, lngLBound, lngUbound)
    
    End If

Exit Sub

ErrHan:

    Debug.Print "StringArray.pSortNumber.Error: " & Err.Number & "; " & Err.Description

End Sub

' V3: Added.
' Sub:             pSplitB04
' Description:     Optimised Split method to return the user string as an array of string data.

Private Sub pSplitB04(pTheString As String, ByRef pDestinationArray() As String, Optional ByRef pTheDelimiter As String = " ", Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)

'... Parameters.
'    R_A pDestinationArray: String   ' ... The array which shall be returned with the string data.
'    RO_ pTheDelimiter: String       ' ... The Delimiter upon which to Split the user string..
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' By Chris Lucas, cdl1051@earthlink.net, 20011208
' http://www.xbeat.net/vbspeed/c_Split.htm

    Dim c&, SLen&, DelLen&, tmp&, Results&()


    On Error GoTo ErrHan:

' -------------------------------------------------------------------
'Dim Expression As String
'    Expression = Me
' -------------------------------------------------------------------

    SLen = LenB(pTheString) \ 2
    DelLen = LenB(pTheDelimiter) \ 2

    ' Bail if we were passed an empty delimiter or an empty pTheString
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve pDestinationArray(0 To 0)
        pDestinationArray(0) = pTheString
        Exit Sub
    End If

    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    tmp = InStr(pTheString, pTheDelimiter)

    Do While tmp
        Results(c) = tmp
        c = c + 1
        tmp = InStr(Results(c - 1) + 1, pTheString, pTheDelimiter)
    Loop

    ' Size our return array
    ReDim Preserve pDestinationArray(0 To c)

    ' Populate the array
    If c = 0 Then
'         lazy man 's call
        pDestinationArray(0) = pTheString
    Else
        ' typical call
        pDestinationArray(0) = Left$(pTheString, Results(0) - 1)
        For c = 0 To c - 2
            pDestinationArray(c + 1) = Mid$(pTheString, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
        Next c
        pDestinationArray(c + 1) = Right$(pTheString, SLen - Results(c) - DelLen + 1)
    End If

    Let pErrMsg = vbNullString
    Let pOK = True


Exit Sub
ErrHan:

    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "StringArray.pSplitB04", Err.Number, Err.Description

End Sub ' ... pSplitB04.

' Function:        pTestArray
' Returns:         Boolean.
' Description:     Returns a Boolean indicating that the internal array is OK to use (True) or Not (False).

Private Function pTestArray() As Boolean

    On Error GoTo ErrHan:
    
    Let pTestArray = UBound(m_StringArray) >= 0 ' ... option base default = 0.
    
Exit Function
ErrHan:

End Function ' ... pTestArray: Boolean.

' Function:        pWriteTextFile
' Returns:         Boolean.
' Description:     Write the contents of a string to a file, optionally in Append mode.

Private Function pWriteTextFile(ByVal pTheText As String, ByVal pTheFileName As String, Optional ByVal pAppendMode As Boolean = False, Optional ByVal pConfirmOverWrite As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean

'... Parameters.
'    V__ pTheFileName: String        ' ... The Full Path Name and Location of the FIle to Write.
'    VO_ pAppendMode: Boolean        ' ... When True, attempts to Append the Text to an existing File.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
Dim iFileNumber As Integer
Dim bFileIsOpen As Boolean
Dim bFileExists As Boolean
Dim iAnswer As VbMsgBoxResult

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    bOK = pCheckFileName(pTheFileName, sErrMsg)
    ' -------------------------------------------------------------------
    
    If bOK Then
    
        If pAppendMode = False And pConfirmOverWrite = True Then
            
            Let bFileExists = pFileExists(pTheFileName, sErrMsg)
            
            If bFileExists = True Then
            
                Let iAnswer = MsgBox("The File [ " & pTheFileName & " ] Exists;" & vbCrLf & "Would you like to over-write this File?", vbYesNo + vbQuestion + vbDefaultButton1, "Over-Write File?")
                
                If iAnswer = vbYes Then
                    
                    Kill pTheFileName
                    
                End If
                
            End If
            
        End If
        
        ' -------------------------------------------------------------------
        ' ... get the next free file number.
        Let iFileNumber = FreeFile()
        
        If pAppendMode Then
            Open pTheFileName For Append As #iFileNumber
        Else
            Open pTheFileName For Output As #iFileNumber
        End If
        
        ' ... if execution flow got here, the file has been open correctly.
        Let bFileIsOpen = True
        
        ' ... print to the file in one single operation.
        Print #iFileNumber, pTheText
        
        Let bOK = True
                
    End If
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let pWriteTextFile = bOK

    On Error Resume Next    ' ... don't fall over because we couldn't release the file handle.
    
    If bFileIsOpen = True Then
        
        Close #iFileNumber
    
    End If

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.pWriteTextFile", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... pWriteTextFile: Boolean.

' Function:        Remove
' Returns:         Boolean.
' Description:     Deletes an Item from the internal buffer (optionally deletes all occurrances or first only).

Public Function Remove(ByVal pItemString As String, Optional ByVal pRemoveAll As Boolean = False, Optional ByRef pRemoveCount As Long = 0, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute Remove.VB_Description = "Deletes an Item from the internal buffer (optionally deletes all occurrances or first only)."

'... Parameters.
'    V__ pItemString: String         ' ... The item to find.
'    VO_ pRemoveAll: Boolean         ' ... Instruction to remove all occurrances of the search string (True).
'    RO_ pRemoveCount: Long          ' ... Returns the number of items removed.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
'Dim tmpFields() As String   ' ... use this n vb6
Dim tmpFields               ' ... use this in vb5
Dim lngRemoveCount As Long

' -------------------------------------------------------------------
' Note:     ' ... uses exact match to find item/s to remove.

    On Error GoTo ErrHan:

    If pTestArray = True Then
            
        Call Compact
        
        Let tmpFields = m_StringArray
        
        Call Clear
                
        If pRemoveAll = True Then
        
            For m_Row = 0 To UBound(tmpFields) 'm_Count - 1
                
                If tmpFields(m_Row) <> pItemString Then
                    
                    Call AddItemString(tmpFields(m_Row))
                
                Else
                                    
                    Let lngRemoveCount = lngRemoveCount + 1
                    
                    Let bOK = True
                
                End If
            
            Next m_Row
        
        Else
        
            ' ... same loop as above but with added lngRemoveCount > 0 added to match condition.
            ' ... wanted to avoid including extra condition testing for RemoveAll/RemoveFirst
            ' ... with no other ideas.
            ' ... V3, yet another bug, had the correct code then changed it then buggered up by adding call to clear above which made m_Count 0!
            For m_Row = LBound(tmpFields) To UBound(tmpFields) '0 To m_Count - 1 '
                
                If tmpFields(m_Row) <> pItemString Or lngRemoveCount > 0 Then
                    
                    Call AddItemString(tmpFields(m_Row))
                
                Else
                                    
                    Let lngRemoveCount = lngRemoveCount + 1
                    
                    Let bOK = True
                
                End If
            
            Next m_Row
        
        End If
        
    End If

ErrResume:

    Let pRemoveCount = lngRemoveCount
    Let pErrMsg = sErrMsg
    Let Remove = bOK


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.Remove", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... Remove: Boolean.

' Function:        Sort
' Returns:         Boolean.
' Description:     Sorts the items in the buffer as text using a Selection Sort algorithm.

Public Function Sort(Optional ByVal pAscending As Boolean = True, Optional ByRef pSorted As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute Sort.VB_Description = "Sorts the items in the buffer as text using a Selection Sort algorithm."

'... Parameters.
'    VO_ pAscending: Boolean         ' ... Instruction to Sort Ascending (True) or Descending (False).
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
'Dim List() As String    ' ... use this in vb6.
Dim List                ' ... use this in vb5.
Dim bSorted As Boolean

' -------------------------------------------------------------------
' Note:     This method returns True if an error was not trapped / generated.
'           Use parameter pSorted to see whether it executed a Sort.
'           Using a Quick Sort algorithm.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    If pTestArray = True Then
        
        If m_Sortable = True Then
            
            Call Compact
            
            Call pQuickSortOnStringArray(m_StringArray, LBound(m_StringArray), UBound(m_StringArray))
            
            If pAscending = False Then
                
                Let List = m_StringArray
                
                For m_Row = LBound(List) To UBound(List)
                    
                    Let m_StringArray(m_Row) = List(UBound(List) - m_Row)
                
                Next m_Row
            
            End If
                        
        Else
            
            Let sErrMsg = "The String buffer may not be Sorted, try setting Sortable to True."
        
        End If
    
    Else
        
        Let sErrMsg = "There is no data to Sort."
    
    End If
    
    Let bOK = True

ErrResume:
    
    Let pSorted = bSorted
    Let pErrMsg = sErrMsg
    Let Sort = bOK

Exit Function
ErrHan:

    Let bSorted = False
    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.Sort", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... Sort: Boolean.

' Property Get:    Sortable
' Type:            Boolean.
' Description:     Write-Once Property to allow Sorting (True) or not (False).

Public Property Get Sortable() As Boolean
Attribute Sortable.VB_Description = "Write-Once Property to allow Sorting (True) or not (False)."

    On Error GoTo ErrHan:

    Let Sortable = m_Sortable

Exit Property
ErrHan:

    Debug.Print "StringArray.Sortable", Err.Number, Err.Description

End Property ' ... Sortable: Boolean.

' Property Let:    Sortable

Public Property Let Sortable(ByVal pNewValue As Boolean)

    On Error GoTo ErrHan:

    Let m_Sortable = pNewValue
'   Call PropertyChanged("Sortable")

Exit Property
ErrHan:

    Debug.Print "StringArray.Sortable", Err.Number, Err.Description

End Property ' ... Sortable: Boolean.

' Function:        SortAsDate
' Returns:         Boolean.
' Description:     Attempts to sort buffer by date value.

Public Function SortAsDate(Optional ByVal pAscending As Boolean = True, Optional ByRef pSorted As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute SortAsDate.VB_Description = "Attempts to sort buffer by date value."

' -------------------------------------------------------------------

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
Dim bSorted As Boolean
Dim numArray() As Long
Dim iLBound As Long
Dim iUBound As Long

' -------------------------------------------------------------------
' Note:     Not only does the sort loop around the buffer a zillion times,
'           before that even happens and after it loops through the entire
'           buffer first to create an array of Longs and then to re-create
'           the array of strings.  The same process is applied to SortAsNumber.
'           It seems to work and the overhead doesn't seem to be too bad
'           perhaps because Longs are more easily managed by VB?
'           Something was required in order to allow dates and numbers to
'           be sorted correctly as sorting their string counterparts couldn't.
' -------------------------------------------------------------------
'           Used RDe's Format$ to replace vb6 FormatDateTime after reading his psc submission ...
'           A Complete set of VB6 functions not available in VB5
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=71812&lngWId=1

    On Error GoTo ErrHan:

    If pTestArray = True Then
        
        If m_Sortable = True Then
            
            Call Compact
            
            Let iLBound = LBound(m_StringArray)
            Let iUBound = UBound(m_StringArray)
            
            ReDim numArray(iUBound)
            
            For m_Row = iLBound To iUBound
                Let numArray(m_Row) = CLng(DateValue(m_StringArray(m_Row)))
            Next m_Row
            
            Call pSortNumber(numArray, iLBound, iUBound)
                        
            If pAscending = True Then
                For m_Row = iLBound To iUBound
                    Let m_StringArray(m_Row) = Format$(numArray(m_Row), "Long Date") ' FormatDateTime(numArray(m_Row), vbLongDate)
                Next m_Row
            Else
                For m_Row = iUBound To iLBound Step -1
                    m_StringArray(iUBound - m_Row) = Format$(numArray(m_Row), "Long Date") ' FormatDateTime(numArray(m_Row), vbLongDate)
                Next m_Row
            End If
            
            Let bSorted = True
                        
        Else
            
            Let sErrMsg = "The String buffer may not be Sorted, try setting Sortable to True."
        
        End If
    
    Else
        
        Let sErrMsg = "There is no data to Sort."
    
    End If

    Let bOK = True

ErrResume:

    Let pSorted = bSorted
    Let pErrMsg = sErrMsg
    Let SortAsDate = bOK


Exit Function
ErrHan:
    
    Let bSorted = False
    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.SortAsDate", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... SortAsDate: Boolean.

' Function:        SortAsNumber
' Returns:         Boolean.
' Description:     Attempts to sort buffer by numeric value.

Public Function SortAsNumber(Optional ByVal pAscending As Boolean = True, Optional ByRef pSorted As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute SortAsNumber.VB_Description = "Attempts to sort buffer by numeric value."

' -------------------------------------------------------------------

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
Dim bSorted As Boolean
Dim numArray() As Long
Dim iLBound As Long
Dim iUBound As Long

' -------------------------------------------------------------------
' Note:     Not only does the sort loop around the buffer a zillion times,
'           before that even happens and after it loops through the entire
'           buffer first to create an array of Longs and then to re-create
'           the array of strings.  The same process is applied to SortAsDate.
'           It seems to work and the overhead doesn't seem to be too bad
'           perhaps because Longs are more easily managed by VB?
'           Something was required in order to allow dates and numbers to
'           be sorted correctly as sorting their string counterparts couldn't.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:

    If pTestArray = True Then
        
        If m_Sortable = True Then
                        
            Call Compact
            
            Let iLBound = LBound(m_StringArray)
            Let iUBound = UBound(m_StringArray)
            
            ReDim numArray(iUBound)
            
            For m_Row = iLBound To iUBound
                Let numArray(m_Row) = CLng(m_StringArray(m_Row))
            Next m_Row
            
            Call pSortNumber(numArray, iLBound, iUBound)
                        
            If pAscending = True Then
                For m_Row = iLBound To iUBound
                    Let m_StringArray(m_Row) = CStr(numArray(m_Row))
                Next m_Row
            Else
                For m_Row = iUBound To iLBound Step -1
                    m_StringArray(iUBound - m_Row) = CStr(numArray(m_Row))
                Next m_Row
            End If
            
            Let bSorted = True
            
        Else
            
            Let sErrMsg = "The String buffer may not be Sorted, try setting Sortable to True."
        
        End If
    
    Else
        
        Let sErrMsg = "There is no data to Sort."
    
    End If

    Let bOK = True

ErrResume:

    Let pSorted = bSorted
    Let pErrMsg = sErrMsg
    Let SortAsNumber = bOK


Exit Function
ErrHan:
    
    Let bSorted = False
    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.SortAsNumber", Err.Number, Err.Description
    Resume ErrResume:

Resume
End Function ' ... SortAsNumber: Boolean.

' Property Get:    ToArray
' Type:            String.
' Description:     This is the internal array of strings currently stored.

'Public Property Get ToArray() As String()
Public Property Get ToArray()
Attribute ToArray.VB_Description = "Provides access to a copy of the internal array of strings currently stored."

' Note:
'       In vb5 change declaration to the one without a return value.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    Call Compact
    
    Let ToArray = m_StringArray

Exit Property
ErrHan:

    Debug.Print "StringArray.ToArray", Err.Number, Err.Description

End Property ' ... ToArray: String.

' Note:
' the default value of the parameter pEscapeChar in ToFile and ToString
' has been changed from "¬" to ":" on advice from TonyYang who said his
' Windows(Chinese) picked up a ? instead and this messed things up,
' ideally the delimiter and escape chars are different and i chose : rather than ,
' thanks Tony.

' Function:        ToFile
' Returns:         Boolean.
' Description:     Writes the contents of the buffer to disk.

Public Function ToFile(ByVal pTheFileName As String, Optional ByVal pTheDelimiter As String = ",", Optional ByVal pEscapeChar As String = ":", Optional ByVal pWrapInParenthesis As Boolean = False, Optional ByVal pEachItemOnNewRow As Boolean = False, Optional ByVal pItemPrefix As String = vbNullString, Optional ByVal pWrapItemsInQuotes As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Boolean
Attribute ToFile.VB_Description = "Writes the contents of the buffer to disk."

'... Parameters.
'    V__ pTheFileName: String        ' ... The Full Path Name and Location of the FIle to Write.
'    VO_ pTheDelimiter: String       ' ... The Delimiter to insert between items (defaults to comma).
'    VO_ pEscapeChar: String         ' ... The escape character if pDelimiter is found in the item string (defaults to ¬ ).
'    VO_ pWrapInParenthesis: Boolean ' ... When True returns the entire string wrapped in brackets e.g. (a, b, c).
'    VO_ pEachItemOnNewRow: Boolean  ' ... When True, adds a new line (vbCRLF) after each item.
'    VO_ pItemPrefix: String         ' ... A Prefix to add ahead of each item e.g. spaces " " (a,b,c to  a, b, c).
'    VO_ pWrapItemsInQuotes: Boolean ' ... Instruction to wrap items with quote characters e.g. (abc to "a","b","c" when delimiter is comma).
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
Dim sTheFileText As String

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
            
    Let sTheFileText = ToString(pTheDelimiter, pEscapeChar, pWrapInParenthesis, pEachItemOnNewRow, pItemPrefix, pWrapItemsInQuotes, bOK, sErrMsg)
    
    If bOK = True Then
        ' ... pWriteTextFile will check that there are no illegal characters in the file name passed.
        Let bOK = pWriteTextFile(sTheFileText, pTheFileName, False, , sErrMsg)
    End If
    
ErrResume:

    Let pErrMsg = sErrMsg
    Let ToFile = bOK

    If bOK = False Then
        Call MsgBox("The Program could not save the data to file:" & vbCrLf & sErrMsg, vbCritical, "String Array")
    End If

Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "StringArray.ToFile", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... ToFile: Boolean.

' Function:        ToString
' Returns:         String.
' Description:     Returns a concatenation of the buffer items with various options.

Public Function ToString(Optional ByVal pTheDelimiter As String = ",", Optional ByVal pEscapeChar As String = ":", Optional ByVal pWrapInParenthesis As Boolean = False, Optional ByVal pEachItemOnNewRow As Boolean = False, Optional ByVal pItemPrefix As String = vbNullString, Optional ByVal pWrapItemsInQuotes As Boolean = False, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As String
Attribute ToString.VB_Description = "Returns a concatenation of the buffer items with various options."

'... Parameters.
'    VO_ pTheDelimiter: String       ' ... The Delimiter to insert between items (defaults to comma).
'    VO_ pEscapeChar: String         ' ... The escape character if pDelimiter is found in the item string (defaults to ¬ ).
'    VO_ pWrapInParenthesis: Boolean ' ... When True returns the entire string wrapped in brackets e.g. (a, b, c).
'    VO_ pEachItemOnNewRow: Boolean  ' ... When True, adds a new line (vbCRLF) after each item.
'    VO_ pItemPrefix: String         ' ... A Prefix to add ahead of each item e.g. spaces " " (a,b,c to  a, b, c).
'    VO_ pWrapItemsInQuotes: Boolean ' ... Instruction to wrap items with quote characters e.g. (abc to "a","b","c" when delimiter is comma).
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' Note:     The code in this method is entirely based upon G. van den Hoven's StringArrayToString method (see the original commented below).
'           Thank you G. van den Hoven.
' -------------------------------------------------------------------

Dim bTemp() As Byte
Dim lSize As Long
Dim lNow As Long
Dim lTotal As Long
Dim lNowArray As Long
Dim tTemp As String
Dim lNow2 As Long
Dim lTotal2 As Long

Dim strPrefix As String
Dim strSuffix As String
Dim lngExtra As Long
'Dim arrTmp() As String  ' ... use this in vb6.
Dim arrTmp              ' ... use this in vb5.
Dim sItem As String

' -------------------------------------------------------------------
    
    On Error GoTo ErrHan:
    
    If pTestArray = False Then
        Let pOK = False
        Let pErrMsg = "There is no data to convert to string."
        Exit Function
    End If
    
    '-----
    ' Calculate size of inputarray
    '-----
    
    Call Compact
    
    Let arrTmp = m_StringArray
    
    lSize = 0
    
    lTotal = UBound(arrTmp)
    
    Let strPrefix = pItemPrefix
    
    Let strSuffix = pTheDelimiter & IIf(pEachItemOnNewRow = True, vbNewLine, vbNullString)
    
    Let lngExtra = Len(strSuffix)

    For lNow = 0 To lTotal
        
        Let sItem = arrTmp(lNow)
'        Let sItem = VBA.Replace$(sItem, pTheDelimiter, pEscapeChar)
        pReplaceChars sItem, pTheDelimiter, pEscapeChar
        
        Let sItem = strPrefix & sItem
        
        If pWrapItemsInQuotes = True Then
            Let sItem = Chr$(34) & sItem & Chr$(34)
        End If
        
        Let arrTmp(lNow) = sItem
        
        Let lSize = lSize + Len(sItem)
        
        If lNow < lTotal Then
            Let lSize = lSize + lngExtra
        End If
        
    Next
    
    '-----
    ' Create byte array which is big
    ' enough to hold all the bytes
    '-----
    ReDim bTemp(0 To lSize - 1) ' ... -1 to stop chr(0)being added to end of output string.
    
    '-----
    ' Convert the string array to a byte array
    '-----
    Let lNow = 0
    Let lNowArray = 0


    While lNow <> lSize
    
        Let tTemp = arrTmp(lNowArray)
                
        If lNowArray < lTotal Then

            Let tTemp = arrTmp(lNowArray) & strSuffix

        End If
        
        
        Let lTotal2 = Len(tTemp)
        
        '-----
        ' Loop through the temp string
        ' and place the byte character
        ' in the correct position
        ' Mid$(...) is faster then Mid(...)
        '-----

        For lNow2 = 0 To lTotal2 - 1
            Let bTemp(lNow + lNow2) = Asc(Mid$(tTemp, lNow2 + 1, 1))
        Next
        
        Let lNow = lNow + lTotal2
        Let lNowArray = lNowArray + 1
    
    Wend
    
    '-----
    ' Convert byte array to string
    '-----
    Let tTemp = StrConv(bTemp(), vbUnicode)
        
    If pWrapInParenthesis = True Then
        Let tTemp = "(" & tTemp & ")"
    End If
    
    Let ToString = tTemp

    Let pOK = True
    
Exit Function

ErrHan:
    
    Let pOK = False
    Let pErrMsg = Err.Description
    
    Debug.Print "StringArray.ToString.Error: " & Err.Number & "; " & Err.Description

End Function ' ... ToString: String

