VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A provides a simple interface to the vb Long data type adding simple math functions such as Add, Subtract, Multiply etc."
' what?
'  a small class for handling a long number
' why?
'  originally intended to help increment a number
'  withing a loop
' when?
'  doing repetitive stuff to a long
'  such as incrementing or decrementing
' how?
'   example of incrementing a number value
'   Dim x As LongNumber
'   Dim y As Loop
'       For y = 1 To 10
'           x.Increment
'       Next y
'
'       Debug.Print x ' ... or Debug.Print x.Value ' ...value is the default property.
'
'       Set x = Nothing
'
' who?
'  d.c.

Option Explicit

' -------------------------------------------------------------------
' Note:
'   Busker's version, 1.
'   Haven't really used anything other than incrementing.
'   Just brought it in when enhancing the search project function
'   to save all this x = x + 1, y = y + 1, kind of stuff.
'   instead: x.Increment: y.Increment' ... it reads better.
' -------------------------------------------------------------------
'   DivideBy and MultiplyBy differ in so far as they can return the
'   result without updating the internal value, default is to update
'   use False on second parameter to ignore update and just do the math.
' -------------------------------------------------------------------

Private m_Value As Long

Public Property Get Value() As Long
Attribute Value.VB_Description = "(Default Property).  Returns the internal number."
Attribute Value.VB_UserMemId = 0

' Property: Value (Long).
' Purpose: accesses the number.

    On Error GoTo ErrHan:
    Value = m_Value
    
Exit Property
ErrHan:
    Debug.Print "LongNumber.Value.Error." & Err.Number & "." & Err.Description

End Property

Public Property Let Value(ByVal dValue As Long)

' Property: Value (Long).
' Purpose: sets the number.
    
    On Error GoTo ErrHan:
    m_Value = dValue
    
Exit Property
ErrHan:
    Debug.Print "LongNumber.Value.Error." & Err.Number & "." & Err.Description

End Property

Public Function Add(ValueToAdd As Long) As Long
Attribute Add.VB_Description = "Returns the result of Adding 'ValueToAdd' to the internal number."

' Function: Add (Long).
' Purpose: add ValueToAdd to the number and return new value for the number.

    On Error GoTo ErrHan:
    
    m_Value = m_Value + ValueToAdd
    
    Add = m_Value

Exit Function
ErrHan:
    Debug.Print "LongNumber.Add.Error." & Err.Number & "." & Err.Description

End Function

Public Function Subtract(ValueToSubtract As Long) As Long
Attribute Subtract.VB_Description = "Returns the result of Subtracting 'ValueToSubtract' from the internal number."

' Function: Subtract (Long).
' Purpose: subtract ValueToSubtract from the number and return new value for the number.
    
    On Error GoTo ErrHan:
    
    m_Value = m_Value - ValueToSubtract
    
    Subtract = m_Value

Exit Function
ErrHan:
    Debug.Print "LongNumber.Subtract.Error." & Err.Number & "." & Err.Description

End Function

Public Function MultiplyBy(ByVal pValueToMultiplyBy As Long, _
                  Optional ByVal pUpdateClassValue As Boolean = True) As Long

' Function: MultiplyBy (Long).
' Purpose: multiply the number by pValueToMultiplyBy and return new value for the number.

Dim lngValue As Long

    On Error GoTo ErrHan:
    
    ' -------------------------------------------------------------------
    lngValue = m_Value   ' ... default, return.
    ' -------------------------------------------------------------------
    
    lngValue = lngValue * pValueToMultiplyBy

ResumeErr:

    If lngValue <> m_Value Then
        If pUpdateClassValue Then
            ' -------------------------------------------------------------------
            ' ... force internal value to be updated.
            m_Value = lngValue
        End If
    End If
    ' -------------------------------------------------------------------
    ' ... return result.
    MultiplyBy = lngValue
    ' -------------------------------------------------------------------
    ' ... clean up
    lngValue = 0&
       
Exit Function
ErrHan:
    Debug.Print "LongNumber.MultiplyBy.Error." & Err.Number & "." & Err.Description
    Resume ResumeErr:

End Function

Public Function DivideBy(ByVal pValueToDivideBy As Long, _
                Optional ByVal pUpdateClassValue As Boolean = True) As Long

' Function: DivideBy (Long).
' Purpose: divide the number by pValueToDivideBy and return new value for the number.

Dim lngValue As Long

    On Error GoTo ErrHan:
    
    
    If pValueToDivideBy <> 0 Then
        ' ... ensure divider is not zero.
        lngValue = lngValue / pValueToDivideBy
    End If

ResumeErr:
    
    If lngValue <> m_Value Then
        If pUpdateClassValue Then
            ' -------------------------------------------------------------------
            ' ... force internal value to be updated.
            m_Value = lngValue
        End If
    End If
    ' -------------------------------------------------------------------
    ' ... return result.
    DivideBy = lngValue
    ' -------------------------------------------------------------------
    ' ... clean up
    lngValue = 0&

Exit Function
ErrHan:
    Debug.Print "LongNumber.DivideBy.Error." & Err.Number & "." & Err.Description
    Resume ResumeErr:
    
End Function

Public Function Increment(Optional IncrementalValue As Long = 1) As Long
Attribute Increment.VB_Description = "Returns the result of Adding 'IncrementalValue' (default = 1) to the internal number."

' Function: Increment (Long).
' Purpose: increment the number by IncrementalValue and return new value for the number.
    
    On Error GoTo ErrHan:
    
    Increment = Add(IncrementalValue)

Exit Function
ErrHan:
    Debug.Print "LongNumber.Increment.Error." & Err.Number & "." & Err.Description

End Function

Public Function Decrement(Optional DecrementalValue As Long = 1) As Long
Attribute Decrement.VB_Description = "Returns the result of Subtracting 'DecrementalValue' (default = 1) from the internal number."

' Function: Decrement (Long).
' Purpose: Reduce the number by DecrementalValue and return new value for the number.
    
    On Error GoTo ErrHan:
    Decrement = Subtract(DecrementalValue)

Exit Function
ErrHan:
    Debug.Print "LongNumber.Decrement.Error." & Err.Number & "." & Err.Description

End Function
