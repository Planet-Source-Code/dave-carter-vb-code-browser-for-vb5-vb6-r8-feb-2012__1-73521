VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------------
' Name:         zTimer
' Purpose:      Aid to capturing code execution performance times.
' Author:       Dave Carter.
' Date:         Thursday 15 July 2010
' ------------------------------------------------------------

' Notes:        This is not a high resolution timer; it simply records start and end times using VB6's Timer function
'               and subtracts start from end to provide an elapsed tick count and then formats this into hrs:mins:secs.000000.
'               Preference for no api / typelib references.
'
' Use:          Dim z As New zTimer, lngLoop As Long
'
'                   z.Name = "First Test"
'                   z.StartTiming
'                   ' ... do stuff.
'                   z.StopTiming
'
'                   z.Name = "Second Test"
'                   z.StartTiming
'                   ' ... do other stuff.
'                   z.StopTiming
'
'                   ' ... Print out all results.
'                   For lngLoop = 1 To z.Count
'                       Debug.Print z.ToString(lngLoop)
'                   Next lngLoop
'
'                   ' ... Print Fastest, Slowest and Average Times.
'                   Debug.Print "Fastest: " & z.ToString(z.Fastest), "Slowest: " & z.ToString(z.Slowest), "Average: " & z.AverageAsString
'

Option Explicit

Private m_Start As Single
Private m_End As Single
Private m_Count As Long
Private m_Timings() As Single
Private m_WithDebug As Boolean
Private m_Duration As Single
Private m_lngLoop As Long
Private m_lngReturn As Long

Public Name As String

Public Function AverageAsString() As String
    AverageAsString = pFormatDuration(Me.AverageTime)
End Function

Public Function AverageTime() As Single
Dim sngTmp As Single
    If m_Count > 0 Then
        For m_lngLoop = 0 To UBound(m_Timings)
            sngTmp = sngTmp + m_Timings(m_lngLoop)
        Next m_lngLoop
        sngTmp = sngTmp / m_Count
        AverageTime = sngTmp
    End If
End Function

Private Sub Class_Terminate()
    If Running = True And m_WithDebug = True Then StopTiming
    Me.Clear
End Sub

Public Sub Clear()
    Erase m_Timings
    m_Count = 0
    m_End = 0
    m_Start = 0
End Sub

Public Function Count() As Long
    Count = m_Count
End Function

Private Property Let Duration(ByVal pNewValue As Single)
    m_Duration = pNewValue
End Property

Public Property Get Duration() As Single
    Duration = m_Duration
End Property

Public Function Fastest() As Long
Dim sngTmp As Single
    m_lngReturn = -1
    If m_Count > 0 Then
        sngTmp = m_Timings(0)
        m_lngReturn = 0
        For m_lngLoop = 1 To UBound(m_Timings)
            If m_Timings(m_lngLoop) < sngTmp Then
                sngTmp = m_Timings(m_lngLoop)
                m_lngReturn = m_lngLoop
            End If
        Next m_lngLoop
        m_lngReturn = m_lngReturn + 1
    End If
    Fastest = m_lngReturn
End Function

Private Function pFormatDuration(pSeconds As Single) As String
Dim sngSeconds As Single, sngMinutes As Single, sngHours As Single, sngDays As Single
    sngSeconds = pSeconds
    If sngSeconds >= 60 Then
        sngMinutes = sngSeconds \ 60
        sngSeconds = sngSeconds Mod 60
        If sngMinutes >= 60 Then
            sngHours = sngMinutes \ 60
            sngMinutes = sngMinutes Mod 60
            If sngHours >= 24 Then
                sngDays = sngHours \ 24
                sngHours = sngHours Mod 24
            End If
        End If
    End If
    pFormatDuration = IIf(sngDays > 0, Format(sngDays, "#,##0") & " ", "") & Format(sngHours, "00") & ":" & Format(sngMinutes, "00") & ":" & Format(sngSeconds, "00.000000")
End Function

Public Function Running() As Boolean
    Running = CBool(m_Start > m_End)
End Function

Public Function Slowest() As Long
Dim sngTmp As Single
    m_lngReturn = -1
    If m_Count > 0 Then
        m_lngReturn = 0
        sngTmp = m_Timings(0)
        For m_lngLoop = 1 To UBound(m_Timings)
            If m_Timings(m_lngLoop) > sngTmp Then
                sngTmp = m_Timings(m_lngLoop)
                m_lngReturn = m_lngLoop
            End If
        Next m_lngLoop
        m_lngReturn = m_lngReturn + 1
    End If
    Slowest = m_lngReturn
End Function

Public Sub StartTiming(Optional ByVal pWIthDebug As Boolean = True)
    m_WithDebug = pWIthDebug
    m_End = 0
    If pWIthDebug = True Then
        Debug.Print "Timer Start: " & Name & ".  " & pFormatDuration(Timer)
    End If
    m_Start = Timer
End Sub

Public Sub StopTiming()
    m_End = Timer
    Duration = m_End - m_Start
    If m_WithDebug = True Then
        Debug.Print "Timer End: " & Name & ".  " & pFormatDuration(m_End)
        Debug.Print "Timer Duration: " & Name & ".  " & Duration
    End If
    ReDim Preserve m_Timings(m_Count)
    m_Timings(m_Count) = Duration
    m_Count = m_Count + 1
End Sub

Public Property Get TimingData() As Single()
Dim tmpTimes() As Single
    If m_Count > 0 Then
        ReDim tmpTimes(UBound(m_Timings))
        For m_lngLoop = 0 To UBound(m_Timings)
            tmpTimes(m_lngLoop) = m_Timings(m_lngLoop)
        Next m_lngLoop
    End If
    TimingData = tmpTimes
End Property

Public Function ToString(Optional pIndex As Long = -1) As String
Dim sngSeconds As Single
    If pIndex >= 0 Then
        If m_Count > 0 Then
            sngSeconds = m_Timings(pIndex - 1)
        End If
    Else
        sngSeconds = Duration
    End If
    ToString = pFormatDuration(sngSeconds)
End Function

Public Function TotalTime() As Single
Dim sngTmp As Single
    If m_Count > 0 Then
        For m_lngLoop = 0 To UBound(m_Timings)
            sngTmp = sngTmp + m_Timings(m_lngLoop)
        Next m_lngLoop
        TotalTime = sngTmp
    End If
End Function

Public Function TotalTimeAsString() As String
    TotalTimeAsString = pFormatDuration(Me.TotalTime)
End Function

