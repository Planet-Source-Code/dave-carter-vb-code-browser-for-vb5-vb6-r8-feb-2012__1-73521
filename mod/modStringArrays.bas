Attribute VB_Name = "modStringArrays"
Attribute VB_Description = "A module to help manage arrays of strings."
Option Explicit


Public Sub QuickSortOnStringArray(ByRef MyArray() As String, ByVal lngLBound As Long, ByVal lngUbound As Long)

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
    
    QuickSortOnStringArray MyArray, lngLBound, p - 1
    QuickSortOnStringArray MyArray, p + 1, lngUbound

Exit Sub

ErrHan:

    Debug.Print "modStringArrays.QuickSortOnStringArray.Error: " & Err.Number & "; " & Err.Description

End Sub


Private Function CompareResult(Value1 As String, Value2 As String, Optional pDescending As Boolean = False)
    CompareResult = (StrComp(Value1, Value2, vbTextCompare) = 1)
    CompareResult = CompareResult Xor pDescending
End Function

Sub ShellSort(pTheStringArray() As String, Optional pDescending As Boolean = False)

' ... http://msdn.microsoft.com/en-us/library/aa155630%28office.10%29.aspx
' ... A Better Shell Sort: Part I
' ... How to Sort Arrays of Any Data Type
' ... By Romke Soldaat

' ... adapted by me to use for strings only.

' ... Note:
' ... Trying to avoid Recursion (QuickSort), API RtlMoveMemory and Pointers
' ... to make it easier to port outside VB.


Dim TempVal As String
Dim i As Long
Dim GapSize As Long
Dim CurPos As Long
Dim FirstRow As Long
Dim LastRow As Long
Dim NumRows As Long

    FirstRow = LBound(pTheStringArray)
    LastRow = UBound(pTheStringArray)
    NumRows = LastRow - FirstRow + 1
    
    Do
        GapSize = GapSize * 3 + 1
    Loop Until GapSize > NumRows
    
    Do
        GapSize = GapSize \ 3
        For i = (GapSize + FirstRow) To LastRow
            CurPos = i
            TempVal = pTheStringArray(i)
            Do While CompareResult(pTheStringArray(CurPos - GapSize), TempVal, pDescending)
                pTheStringArray(CurPos) = pTheStringArray(CurPos - GapSize)
                CurPos = CurPos - GapSize
                If (CurPos - GapSize) < FirstRow Then Exit Do
            Loop
            pTheStringArray(CurPos) = TempVal
        Next
    Loop Until GapSize = 1
    
End Sub

Public Sub SplitString(ByRef TheString As String, ByRef TheArray() As String, Optional ByVal TheDelimiter As String = vbCrLf, Optional ByRef NoOfItems As Long = 0, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)

' ... Createa a String Array from a Source String using the
' ... delimiter passed.

' Note:
'       Not a VB6.VBA.Spli Clone for VB5
'       This is a Sub not a function and the
'       array we're after is passed by ref as TheArray.

'... Parameters.
'    R__ TheString: String           ' ... The String data to be split.
'    R_A TheArray: String            ' ... The String Array returned from the split operation.
'    RO_ TheDelimiter: String        ' ... The Delimiter with which to split the String data.
'    RO_ NoOfItems: Long             ' ... The number of array elements (base 1) generated from the split operation.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.
' -------------------------------------------------------------------
Dim lngFound As Long
Dim lngLoop As Long
Dim lngDelCount As Long
Dim lngStart As Long
Dim lngRet() As Long
Dim lngDelLength As Long
Dim lngTextLength As Long
' -------------------------------------------------------------------
' Note:
' Thanks to Chris Lucas (SplitB04 on VBSpeed).
' -------------------------------------------------------------------
    On Error GoTo ErrHan:
    
    lngDelLength = Len(TheDelimiter)
    lngTextLength = Len(TheString)
    
    If lngTextLength = 0 Or lngDelLength = 0 Then
        GoTo NothingDoing:
    End If
    
    lngStart = 1
    lngFound = InStr(lngStart, TheString, TheDelimiter)
    
    If lngFound = 0 Then
        GoTo NothingDoing:
    End If
    
    ReDim lngRet(lngTextLength)
    lngRet(0) = 1                                                   ' ... first pos is 1.
    
' -------------------------------------------------------------------
' v6, attempted optimisation, reduce work within loop
'     readability is second to performance.
'     leaving last effort in for now.

'    Do While lngFound > 0
'        lngStart = lngFound + lngDelLength
'        lngDelCount = lngDelCount + 1
'        lngRet(lngDelCount) = lngStart
'        lngFound = InStr(lngStart, TheString, TheDelimiter)
'    Loop
'    Do While tmp
'        Results(c) = tmp
'        c = c + 1
'        tmp = InStr(Results(c - 1) + 1, pTheString, pTheDelimiter)
'    Loop
    
'    Do While lngFound ' > 0 ' ... you see that > 0 , well it slows this down real bad!!!
'        lngDelCount = lngDelCount + 1
'        lngRet(lngDelCount) = lngFound + lngDelLength
'        lngFound = InStr(lngRet(lngDelCount), TheString, TheDelimiter) ' ... take the vba. out of instr
'    Loop
    Do While lngFound ' > 0 ' ... you see that > 0 , well it slows this down real bad!!!
        lngDelCount = lngDelCount + 1
        lngRet(lngDelCount) = lngFound + lngDelLength
        lngFound = InStr(lngRet(lngDelCount), TheString, TheDelimiter)
    Loop
    
' -------------------------------------------------------------------

    lngDelCount = lngDelCount + 1
    ReDim Preserve lngRet(lngDelCount)
    lngRet(lngDelCount) = lngTextLength                             ' ... last pos is length of text.
    
    ReDim TheArray(lngDelCount - 1)
    For lngLoop = 0 To lngDelCount - 2
        TheArray(lngLoop) = Mid$(TheString, lngRet(lngLoop), (lngRet(lngLoop + 1) - lngDelLength) - lngRet(lngLoop))
    Next lngLoop
    TheArray(lngLoop) = Mid$(TheString, lngRet(lngDelCount - 1))
    
    Let pErrMsg = vbNullString
    Let pOK = True
Quit:
    NoOfItems = lngDelCount
    lngDelCount = 0: lngLoop = 0: lngStart = 0
    lngDelLength = 0: lngTextLength = 0
    
    On Error GoTo 0
    Erase lngRet
Exit Sub
ErrHan:
    Let pErrMsg = Err.Description
    Let pOK = False
    Debug.Print "modTypes.SplitString", Err.Number, Err.Description
NothingDoing:
    lngDelCount = 1
    ReDim TheArray(0 To 0)
    TheArray(0) = TheString
    GoTo Quit:
End Sub ' ... SplitString.

