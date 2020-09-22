VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Class to manage an Array of Strings, no API/TLB or VB6 only VBA functions."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' SArray
' a class to help the creation and use of an array of string items.
' attempts a couple of optimisations for keeping a unique list of items and
' attempts to sort and persist item data.

Option Explicit

' -------------------------------------------------------------------
' ...Properties
Public Name As String           ' whatever purpose
Attribute Name.VB_VarDescription = "Sets / Returns a name for an instance of this class."
Public Tag As String            ' whatever purpose
Attribute Tag.VB_VarDescription = "Sets / Returns a tag for an instance of this class.  The tag is a user defined string and can be used for any purpose."
' -------------------------------------------------------------------
Private mPos() As Long          ' positions of elements within string buffer

Private mArr() As String        ' the array
Private mDupesOK As Boolean     ' duplicates allowed
Private mData As String         ' test string for duplicate entries
Private mDataLen As Long        ' the length of the dupe test data string
Private mBuffLen As Long        ' the length of the dupe test buffer string
Private mBuffChunk As Long      ' the size to add to the dupe test buffer when too small
Private mCount As Long          ' number of items added to array
Private mSortable As Boolean    ' describes if array will be sortable or fixed
Private mSorted As Boolean      ' describes if array has been sorted, adding a new item on top of a sorted array will set this to false

Private Const cBuffChunk As Long = 100      ' default chunk size when adding array elements
Private Const cDupesOK As Boolean = True    ' default duplicates allowed setting

Sub Remove(pIndex As Long)
Attribute Remove.VB_Description = "Removes an item from the array by (1 based) index."
Dim bSorted As Boolean
Dim tmp '() As String
Dim i As Long
Dim j As Long
Dim c As Long

    If IndexExists(pIndex) Then
        c = mCount
        j = pIndex - 1
        bSorted = mSorted
        tmp = mArr
        Clear
        For i = 0 To c - 1
            If i <> j Then
                AddItem tmp(i)
            End If
        Next i
        If bSorted And mCount Then
            Sort
        End If
    End If
    Erase tmp
    
End Sub

Property Get Sortable() As Boolean
Attribute Sortable.VB_Description = "Sets / Returns a Boolean describing whether the array buffer items are allowed to be sorted."
    Sortable = mSortable
End Property

Property Let Sortable(pVal As Boolean)
    mSortable = pVal
End Property

Property Get ItemAsSArray(ByVal pIndex As Long, _
                 Optional ByVal pDelimiter As String = ",") As SArray
Attribute ItemAsSArray.VB_Description = "Returns an item in the array buffer as a new SArray object.  This requires the item to include delimited items and therefore needs a delimiter character to split the item data."
                 
Dim sTmp As String
    
    sTmp = Item(pIndex)
    
    Set ItemAsSArray = New SArray
    ItemAsSArray.FromString sTmp, pDelimiter
    
    sTmp = vbNullString
    
End Property

Property Get ItemAsNumberValue(ByVal pIndex As Long) As Double
Attribute ItemAsNumberValue.VB_Description = "Returns an item in the array buffer as a number, more specifically a double.  Zero will be returned when the value is 0 or when the element is not numeric."

Dim sTmp As String

    sTmp = Item(pIndex)
    If IsNumeric(sTmp) Then
        ItemAsNumberValue = CDbl(sTmp)
    End If
    
    sTmp = vbNullString
    
End Property

Property Get Item(ByVal pIndex As Long) As String
Attribute Item.VB_Description = "Sets / Returns an Item in the array by its 1 based index."
Attribute Item.VB_UserMemId = 0
    
    On Error GoTo ErrHan:
    
    If pIndex > mCount Then Err.Raise vbObjectError + 1000, , "Item Index " & pIndex & " not available"
    If pIndex < 1 Then pIndex = 1
    
    Item = mArr(pIndex - 1)

Exit Property
ErrHan:
    Debug.Print "SArray.Item.Error: " & Err.Number & "; " & Err.Description

End Property

Property Let Item(ByVal pIndex As Long, pVal As String)

    On Error GoTo ErrHan:
    
    If pIndex > mCount Then Err.Raise vbObjectError + 1000, , "Item Index not available"
    If pIndex < 1 Then pIndex = 1
    
    mArr(pIndex - 1) = pVal

Exit Property
ErrHan:
    Debug.Print "SArray.Item.Error: " & Err.Number & "; " & Err.Description

End Property

Property Get DuplicatesAllowed() As Boolean
Attribute DuplicatesAllowed.VB_Description = "Sets / Returns a Boolean descrbing whether exact duplicate entries are allowed to be add, True = Yes, False = No."
    DuplicatesAllowed = mDupesOK
End Property

Property Let DuplicatesAllowed(pVal As Boolean)
    mDupesOK = pVal
End Property

Property Get Count() As Long
Attribute Count.VB_Description = "Returns the no. of items in the buffer array of strings"
    Count = mCount
End Property

Property Get ChunkSize() As Long
Attribute ChunkSize.VB_Description = "Sets / Returns the no. of items to add to the array  buffer when it needs to expand (also no. of chars to add to string buffer)."
    ChunkSize = mBuffChunk
End Property

Property Let ChunkSize(pVal As Long)
    If pVal > 0 Then
        mBuffChunk = pVal
    End If
End Property


' -------------------------------------------------------------------

' Function:        AddItemString
' Returns:         Boolean.
' Description:     Add a String value to the internal array.

Public Function AddItemString(ByVal pItemString As String, Optional pExistsTest As Byte = 0, Optional ByRef pErrMsg As String = vbNullString) As Boolean

     AddItemString = Me.AddItem(pItemString)

End Function ' ... AddItemString: Boolean.


Function AddItem(ByVal pItem As String) As Boolean
Attribute AddItem.VB_Description = "Adds a string to the internal array of strings (doesn't add empty strings)."

Dim sItem As String
Dim iLen As Long
Dim iBuffLen As Long
Dim iStart As Long

    On Error GoTo ErrHan:
    
    sItem = pItem & Chr$(2)
    If mCount = 0 Then sItem = Chr$(2) & sItem
    ' binary existence check in string buffer ---------------------------
    If InStr(1, mData, Chr$(2) & pItem & Chr$(2)) > 0 Then
        If mDupesOK = False Then Exit Function
    End If
    iLen = Len(sItem)
    ' add item to string buffer for faster searching -------------------
    If mDataLen + iLen > mBuffLen Then
        iBuffLen = mBuffChunk
        If iLen > mBuffChunk Then iBuffLen = iBuffLen + iLen
        ' -------------------------------------------------------------------
        mData = mData & String$(iBuffLen, 0)
        mBuffLen = Len(mData)
    End If
    iStart = 1
    If mDataLen Then iStart = iStart + mDataLen
    ' -------------------------------------------------------------------
    Mid$(mData, iStart, iLen) = sItem
    mDataLen = mDataLen + iLen
    ' add item to array -------------------------------------------------
    If mCount Mod mBuffChunk = 0 Then
        ReDim Preserve mArr(mCount + mBuffChunk)    ' item
        ReDim Preserve mPos(mCount + mBuffChunk)    ' item position in buffer
    End If
    ' -------------------------------------------------------------------
    mArr(mCount) = pItem:           mPos(mCount) = iStart
    mCount = mCount + 1
    ' -------------------------------------------------------------------
    AddItem = True
    mSorted = False
    
Exit Function
ErrHan:
    AddItem = False
    Debug.Print "SArray.AddItem.Error: " & Err.Number & "; " & Err.Description
    
End Function

Function FromFile(ByRef pFileName As String, _
         Optional ByVal pDelimiter As String = ",") As Boolean
Attribute FromFile.VB_Description = "Loads data from a file when passed a filename and row delimiter."
   
Dim f As Long
Dim l As Long
Dim b() As Byte
Dim sTxt As String

    On Error GoTo ErrHan:
            
    f = FreeFile
    
    Open pFileName For Binary As #f
        
    l = LOF(f)
    If l Then
        ReDim b(0 To l - 1)
        Get #f, , b
        sTxt = StrConv(b, vbUnicode)
    End If
    
    FromString sTxt, pDelimiter

    FromFile = True
    
ResumeError:
    
    On Error GoTo 0
    Close #f
    
    Erase b
    sTxt = vbNullString
    
Exit Function
ErrHan:
    Debug.Print "SArray.FromFile.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Function ToFile(ByRef pFileName As String, _
       Optional ByVal pDelimiter As String = ",") As Boolean
Attribute ToFile.VB_Description = "Convers internal array to a string with a delimiter between elements and saves this to file."

Dim sTmp As String
Dim iFN As Long

    On Error GoTo ErrHan:
    
    If mCount Then
        
        sTmp = ToString(pDelimiter)
        
        If Len(sTmp) Then
        
            iFN = FreeFile
            Open pFileName For Output As #iFN
            Print #iFN, sTmp;
            
        End If
    
    End If

ResumeError:
    
    On Error GoTo 0
    If iFN Then Close #iFN
    
    sTmp = vbNullString
    
Exit Function
ErrHan:
    Debug.Print "SArray.ToFile.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Function

Function IndexExists(pIndex As Long) As Boolean
Attribute IndexExists.VB_Description = "Returns true if (1 based) index exists in array else false."
    If mCount Then
        If pIndex <= mCount Then
            If pIndex Then
                IndexExists = True
            End If
        End If
    End If
End Function

Function ItemExists(pItem As String, _
     Optional ByRef pIndex As Long = 0, _
     Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare) As Boolean
Attribute ItemExists.VB_Description = "Checks to see if an item exists in the buffer array and if so, returns its 1 based index."

Dim sTest As String
Dim sSource As String
Dim iFind As Long
Dim i As Long
    
    pIndex = 0
    If Len(pItem) = 0 Then Exit Function
    
    sTest = Chr$(2) & pItem & Chr$(2)
    ' check item in string buffer ---------------------------------------
    If pCompare <> vbBinaryCompare Then
        sSource = LCase$(mData)
        sTest = LCase$(sTest)
        iFind = InStr(1, sSource, sTest)
    Else
        iFind = InStr(1, mData, sTest)
    End If
        
    Select Case iFind
        Case 1
            pIndex = 1
        Case Is > 1
            ' get index from pos buffer --------------------------------------
            iFind = iFind + 1
            For i = 0 To mCount - 1
                If mPos(i) = iFind Then
                    pIndex = i + 1
                    Exit For
                End If
            Next i
    End Select
    
    ItemExists = iFind
    
    sSource = vbNullString
    sTest = vbNullString
    
End Function

Sub Clear()
Attribute Clear.VB_Description = "Clears all internal resources and sets up a new array buffer."
    
    Erase mArr
    Erase mPos

    mCount = 0
    mData = vbNullString
    mDataLen = 0
    mBuffLen = 0
    mSorted = False
    
    ReDim mArr(0)
    ReDim mPos(0)
    
End Sub

Sub FromString(ByRef pString As String, _
      Optional ByVal pDelimiter As String = ",")
Attribute FromString.VB_Description = "Loads data from some text when a string and delimiter are passed."

Dim tmp() As String

Dim c As Long
Dim iFind As Long
Dim iDLen As Long
Dim iSLen As Long
Dim iPos() As Long

    On Error GoTo ErrHan:
    
    Clear
    ' -------------------------------------------------------------------
    If Len(pString) = 0 Then Err.Raise vbObjectError + 1000, , "No String Data"
    If Len(pDelimiter) = 0 Then Err.Raise vbObjectError + 1000, , "No Delimiter"
    ' -------------------------------------------------------------------
    iSLen = Len(pString)
    iDLen = Len(pDelimiter)
    ' -------------------------------------------------------------------
    ReDim iPos(0 To iSLen)
    iFind = InStr(pString, pDelimiter)
    ' -------------------------------------------------------------------
    Do While iFind
        iPos(c) = iFind
        c = c + 1
        iFind = InStr(iPos(c - 1) + 1, pString, pDelimiter)
    Loop
    ' -------------------------------------------------------------------
    ReDim tmp(0 To c)

    If c = 0 Then
        AddItem pString
    Else
        tmp(0) = Left$(pString, iPos(0) - 1)
        For c = 0 To c - 2
            tmp(c + 1) = Mid$(pString, _
                iPos(c) + iDLen, _
                iPos(c + 1) - iPos(c) - iDLen)
        Next c
        tmp(c + 1) = Right$(pString, iSLen - iPos(c) - iDLen + 1)
        ' -------------------------------------------------------------------
        For c = 0 To UBound(tmp)
            AddItem tmp(c)
        Next c
    End If
    
ResumeError:
    
    Erase tmp
    Erase iPos
    
Exit Sub
ErrHan:
    Debug.Print "SArray.FromString.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Sub Sort()
Attribute Sort.VB_Description = "When Sortable = True, sorts the array buffer items in ascending order.  When Sortable = False, no sorting is performed."

    On Error GoTo ErrHan:
    
    If mCount = 0 Or mCount = 1 Then Exit Sub ' Err.Raise vbObjectError + 1000, , "Nothing to Sort"
    If mSortable = False Then Err.Raise vbObjectError + 1000, , "Array not Sortable"
    
    If UBound(mArr) >= mCount Then
        ' compact is required to set the correct upper limit to the array
        ReDim Preserve mArr(mCount - 1)
    End If
    
    pSortAsString LBound(mArr), UBound(mArr)
    
    pUpdateBuffer
    
    mSorted = True
    
Exit Sub
ErrHan:
    Debug.Print "SArray.Sort.Error: " & Err.Number & "; " & Err.Description

End Sub

Function ToString(Optional ByVal pDelimiter As String = vbCrLf) As String
Attribute ToString.VB_Description = "Joins the array buffer elements into a single string with a given item delimiter."
' see sarray 1.cls for alternate method

Dim sTmp As String

Dim i As Long
Dim tl As Long
Dim s As Long
Dim ml As Long
Dim xl() As Long

    ReDim xl(mCount - 1)
    
    For i = 0 To mCount - 1
        xl(i) = Len(mArr(i))
        tl = tl + xl(i)
    Next i
    
    tl = tl + ((mCount - 1) * Len(pDelimiter))
    s = 1
    sTmp = String$(tl, 0)
    For i = 0 To mCount - 1
        ml = xl(i) + Len(pDelimiter)
        Mid$(sTmp, s, ml) = mArr(i) & pDelimiter
        s = s + ml
    Next i
    
    ToString = sTmp
    
    sTmp = vbNullString
    Erase xl
    
End Function

Private Sub pUpdateBuffer()

' reorders items in the string buffer and position array
' context: follows a sort on the array items
' this is required to re-enable searching and returning correct item index

Dim i As Long
Dim iStart As Long
Dim iLen As Long

    If mCount = 0 Then Exit Sub
    
    iStart = 2
    For i = 0 To mCount - 1
        mPos(i) = iStart - 1
        If i Then mPos(i) = mPos(i) + 1
        iLen = Len(mArr(i)) + 1
        Mid$(mData, iStart, iLen) = mArr(i) & Chr$(2)
        iStart = iStart + iLen
    Next i

End Sub

Private Sub pSortAsString(ByVal lngLBound As Long, ByVal lngUbound As Long)

' With thanks to Optikon:
' XtremeVBTalk: http://www.xtremevbtalk.com/showthread.php?t=78889&highlight=Quick+Sort
' recursive method

Dim lngPivot As String
Dim k As Long
Dim p As Long
Dim lngTemp As String
    
    On Error GoTo ErrHan:
    
    If lngLBound >= lngUbound Then
        Exit Sub
    End If
    
    k = lngLBound + 1
    
    If lngUbound = k Then
        If LCase$(mArr(lngLBound)) > LCase$(mArr(lngUbound)) Then
            lngTemp = mArr(lngLBound)
            mArr(lngLBound) = mArr(lngUbound)
            mArr(lngUbound) = lngTemp
        End If
        Exit Sub
    End If
    
    lngPivot = LCase$(mArr(lngLBound))
    p = lngUbound
    
    Do Until (LCase$(mArr(k)) > lngPivot) Or (k >= lngUbound)
        k = k + 1
    Loop
    
    Do Until LCase$(mArr(p)) <= lngPivot
        p = p - 1
    Loop
    
    Do While k < p
    
        lngTemp = mArr(k)
        mArr(k) = mArr(p)
        mArr(p) = lngTemp
        
        Do
            k = k + 1
        Loop Until LCase$(mArr(k)) > lngPivot
        
        Do
            p = p - 1
        Loop Until LCase$(mArr(p)) <= lngPivot
    
    Loop
    
    lngTemp = mArr(p)
    mArr(p) = mArr(lngLBound)
    mArr(lngLBound) = lngTemp
    
    pSortAsString lngLBound, p - 1
    pSortAsString p + 1, lngUbound

Exit Sub
ErrHan:
    Debug.Print "SArray.pSortAsString.Error: " & Err.Number & "; " & Err.Description

End Sub

Private Sub Class_Initialize()
    
    mDupesOK = cDupesOK         ' set duplicates allowed
    mBuffChunk = cBuffChunk     ' set chunk size for adding array elements
    ReDim mArr(0)               ' initialise array with 1 element
    ReDim mPos(0)
    
End Sub

Private Sub Class_Terminate()
    
    Erase mPos
    Erase mArr
    mData = vbNullString

End Sub


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
    
    Let pKnownItem = mArr(pItemIndex - 1)

Exit Function

ErrHan:

    Debug.Print "StringArray.pKnownItem.Error: " & Err.Number & "; " & Err.Description

End Function

' Function:        ItemAsStringArray
' Returns:         StringArray.
' Description:     Returns a string converted into a StringArray via a Delimiter that defaults to a colon ( : ) if omitted.

Public Function ItemAsStringArray(ByVal pItemIndex As Long, Optional ByVal pTheDelimiter As String = ":", Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As SArray

'... Parameters.
'    V__ pItemIndex: Long            ' ... The index to query.
'    VO_ pTheDelimiter: String       ' ... The Delimiter to use to shred the data content.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

' -------------------------------------------------------------------

Dim objReturn As SArray ' StringArray         ' ... a return value to this function.
Dim bOK As Boolean                   ' ... Returns success or failure of this method.
Dim sErrMsg As String                ' ... Returns an error message trapped / generated here-in.
Dim sTmpItem As String
Dim bIndexed As Boolean

' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    ' ... create a new string array regardless, even if it will have no data,
    ' ... personal preference rather than return nothing.
    
    Set objReturn = New SArray ' StringArray
    
    ' -------------------------------------------------------------------
    
    Let bIndexed = IndexExists(pItemIndex) ', sErrMsg)
        
    ' -------------------------------------------------------------------
    
    If bIndexed = True Then
    
        ' ... use pKnownItem as have already tested index above.
        Let sTmpItem = pKnownItem(pItemIndex)
        
        ' -------------------------------------------------------------------
        ' ... re-use FromString on the Item.
'        Let bOK = objReturn.FromString(sTmpItem, pTheDelimiter) ', , , , sErrMsg)
        objReturn.FromString sTmpItem, pTheDelimiter
        bOK = True
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
