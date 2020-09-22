VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBPSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "draft class researching searching the files of a vbp for a substring."
Option Explicit

Private mInitialised As Boolean
Private moVBPInfo As VBPInfo
Private mAttributeDelimiter As String

Public Sub Clear()
Attribute Clear.VB_Description = "Clears resources and resets module fields to empty."

    On Error GoTo ErrHan:

    mInitialised = False
    Set moVBPInfo = Nothing
    
Exit Sub
ErrHan:

    Debug.Print "VBPSearch.Clear", Err.Number, Err.Description
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
    
Exit Sub
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "VBPSearch.Init", Err.Number, Err.Description
    Resume ErrResume:


End Sub ' ... Init.

Public Property Get Initialised() As Boolean
Attribute Initialised.VB_Description = "Returns a boolean describing whether the tree view has been set up etc."
    
    Initialised = mInitialised
    
End Property

' F   SearchString: Long.  ' ... Searches the text in the files of a vbp for a string and returns number of matches.

' F   SearchFiles: Long.  ' ... Searches the text in the files of a vbp for a string and returns number of matches.

' Function:        SearchFiles
' Returns:         Long.
' Description:     Searches the text in the files of a vbp for a string and returns number of matches.

Public Function SearchFiles(ByRef pSearch As String, _
                            ByRef pSSAFiles As StringArray, _
                   Optional ByRef pStart As Long = 0, _
                   Optional ByVal pCompare As VbCompareMethod = vbBinaryCompare, _
                   Optional ByVal pWholeWordOnly As Boolean = False, _
                   Optional ByRef pOK As Boolean = False, _
                   Optional ByRef pErrMsg As String = vbNullString) As Long



'... Parameters.
'    R__ pSearch: String             ' ... The substring to search for in the source data.
'    R__ pSSAFiles: StringArray          ' ... A StringArray object to return array of file names where search was found.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

'Dim lngReturn As Long ' ... a return value to this function.
Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
'Dim sFileText As String
Dim lngCount As Long
Dim tDataInfo As DataInfo
Dim sTmpArray As StringArray
Dim lngLoop As Long
Dim sTmpString As StringWorker
Dim lngFiles() As Long
Dim lngFindCount As Long
Dim lngFound As Long
Dim lngTotal As Long
Dim lngPos() As Long
Dim lngThisCount As Long
'Dim lngCompMethod As VbCompareMethod
'Dim zT As zTimer

    On Error GoTo ErrHan:
    ' used
    If mInitialised = False Then
        Err.Raise vbObjectError + 1000, , "Unable to Search without a valid VBPInof instance."
    End If
    
    If Len(pSearch) = 0 Then
        Err.Raise vbObjectError + 1000, , "Unable to Search without a valid Search String."
    End If
    
'    Set zT = New zTimer
'    zT.Name = "VBPSearch.SearchFiles: " & pSearch
'    zT.StartTiming
    
    Set pSSAFiles = New StringArray
    
    Set sTmpArray = moVBPInfo.FilesData
    lngCount = sTmpArray.Count
    
    If lngCount > 0 Then
    
        For lngLoop = 1 To lngCount
        
            modGeneral.ParseDataInfoItem sTmpArray, lngLoop, tDataInfo, mAttributeDelimiter
            
            If tDataInfo.Exists Then
                
                Set sTmpString = New StringWorker
                sTmpString.FromFile tDataInfo.ExtraInfo
                
                lngFound = sTmpString.Find(pSearch, pStart, pCompare, pWholeWordOnly, bOK, sErrMsg)
                
                If lngFound > 0 Then
                    
                    ReDim Preserve lngFiles(lngFindCount)
                    lngFiles(lngFindCount) = lngLoop
                    lngFindCount = lngFindCount + 1
                    
                    lngThisCount = sTmpString.FindAllMatches(pSearch, lngPos, pStart, pCompare, pWholeWordOnly, sErrMsg)
                    lngTotal = lngTotal + lngThisCount ' + sTmpString.FindAllMatches(pSearch, lngPos)
                    
                    pSSAFiles.AddItemString sTmpArray(lngLoop) & mAttributeDelimiter & lngThisCount
                    Erase lngPos
                    
                End If
                
                If Not sTmpString Is Nothing Then Set sTmpString = Nothing
                
            End If
            
        Next lngLoop
        
        If Not pSSAFiles Is Nothing Then
            If pSSAFiles.Count > 0 Then pSSAFiles.Sort
        End If
        
    End If
    
'    Debug.Print "Files Read: " & lngCount
'    Debug.Print "Files Containing " & pSearch & " = " & lngFindCount
'    lngReturn = lngTotal

    Let sErrMsg = vbNullString
    Let bOK = True

ErrResume:
    
    On Error GoTo 0
    
    Let pErrMsg = sErrMsg
    Let pOK = bOK
    Let SearchFiles = lngTotal
    
    If Not sTmpArray Is Nothing Then
        Set sTmpArray = Nothing
    End If
    
    Erase lngPos
    
    lngFindCount = 0&
    lngFound = 0&
    lngTotal = 0&
    lngThisCount = 0&
    lngLoop = 0&
    sErrMsg = vbNullString
    
Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "VBPSearch.SearchFiles", Err.Number, Err.Description
    Resume ErrResume:
Resume

End Function ' ... SearchFiles: Long.



' Function:        SearchString
' Returns:         Long.
' Description:     Searches the text in the files of a vbp for a string and returns number of matches.

Public Function SearchString(pSearch As String, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString) As Long
Attribute SearchString.VB_Description = "Searches the text in the files of a vbp for a string and returns number of matches."

'... Parameters.
'    RO_ pOK: Boolean                ' ... Returns the Success, True, or Failure, False, of this method.
'    RO_ pErrMsg: String             ' ... Returns an error message trapped / generated here-in.

Dim lngReturn As Long ' ... a return value to this function.
Dim bOK As Boolean ' ... Returns success or failure of this method.
Dim sErrMsg As String ' ... Returns an error message trapped / generated here-in.
'Dim sFileText As String
Dim lngCount As Long
Dim tDataInfo As DataInfo
Dim sTmpArray As StringArray
Dim lngLoop As Long
Dim sTmpString As StringWorker
Dim lngFiles() As Long
Dim lngFindCount As Long
Dim lngFound As Long
Dim lngTotal As Long
Dim lngPos() As Long

    On Error GoTo ErrHan:
    
    If mInitialised = False Then
        Err.Raise vbObjectError + 1000, , "Unable to Search without a valid VBPInof instance."
    End If
    
    If Len(pSearch) = 0 Then
        Err.Raise vbObjectError + 1000, , "Unable to Search without a valid Search String."
    End If
    
    Set sTmpArray = moVBPInfo.FilesData
    lngCount = sTmpArray.Count
    
    If lngCount > 0 Then
        For lngLoop = 1 To lngCount
        
            modGeneral.ParseDataInfoItem sTmpArray, lngLoop, tDataInfo, mAttributeDelimiter
            
            If tDataInfo.Exists Then
                Set sTmpString = New StringWorker
                sTmpString.FromFile tDataInfo.ExtraInfo
                lngFound = sTmpString.Find(pSearch)
                If lngFound > 0 Then
                    ReDim Preserve lngFiles(lngFindCount)
                    lngFiles(lngFindCount) = lngLoop
                    lngFindCount = lngFindCount + 1
                    lngTotal = lngTotal + sTmpString.FindAllMatches(pSearch, lngPos)
                End If
                If Not sTmpString Is Nothing Then Set sTmpString = Nothing
            End If
        Next lngLoop
        
    End If
    Debug.Print "Files Read: " & lngCount
    Debug.Print "Files Containing " & pSearch & " = " & lngFindCount
    lngReturn = lngTotal
    Let sErrMsg = vbNullString
    Let bOK = True

ErrResume:

    Let pErrMsg = sErrMsg
    Let pOK = bOK
    Let SearchString = lngReturn


Exit Function
ErrHan:

    Let sErrMsg = Err.Description
    Let bOK = False
    Debug.Print "VBPSearch.SearchString", Err.Number, Err.Description
    Resume ErrResume:


End Function ' ... SearchString: Long.


Private Sub Class_Terminate()
    Clear
End Sub