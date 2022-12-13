VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class System.Diagnostics.StopWatch
'The datatype Currency in VB is a 64Bit-Datatype with a fixed point
'part. this means at the Bit.value &H2710 (resp decimal 10000) the value 1
'is defined. The smallest positive value of currency is 0.0001

'because ticks contains the factor 10000 as well, so the VB-datatype Currency
'is the right datatype for time measurements wit hthe QueryPerformance-API

'the correcting factor for the datatype Currency:
Private Const CurCorrect          As Long = 10000
Private Const TicksPerMillisecond As Long = 10000 / CurCorrect '= 1 '&H2710
Private Const TicksPerSecond      As Long = 10000000 / CurCorrect   '&H989680

'in the .net class of the same name here the datatype long resp Int64 will be used.
Private m_elapsed          As Currency
Private m_startTimeStamp   As Currency
Private m_isRunning        As Boolean
Private m_frequency        As Currency
Private m_isHighResolution As Boolean
Private m_tickFrequency    As Double

Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long

Private Sub Class_Initialize()
    If Not (QueryPerformanceFrequency(m_frequency) = 1) Then
        m_isHighResolution = False
        m_frequency = TicksPerSecond
        m_tickFrequency = 1
    Else
        m_isHighResolution = True
        m_tickFrequency = (TicksPerSecond / CDbl(m_frequency))
    End If
End Sub

' #################### '   Public Shared Properties   ' #################### '
Public Property Get Frequency() As Currency
    'Diese Multiplikation mit CurCorrect ist nur dazu da, die Klasse
    'nach auﬂen hin konsistent zu halten mit der gleichnamigen Klasse
    'aus dem .netFX
    Frequency = m_frequency * CurCorrect
End Property
'Shared-Member Properties
Public Property Get IsHighResolution() As Boolean
    IsHighResolution = m_isHighResolution
End Property

' #################### '      Public Properties       ' #################### '
Public Property Get Elapsed() As Currency
    Elapsed = GetElapsedDateTimeTicks
End Property

Public Property Get ElapsedToString() As String 'TimeSpan
    Dim edt As Currency: edt = GetElapsedDateTimeTicks
    ElapsedToString = TimeSpanToString(edt)
End Property

Public Property Get ElapsedMilliseconds() As Currency 'Long 'LongInt
    ElapsedMilliseconds = GetElapsedDateTimeTicks / TicksPerMillisecond ' = 1
End Property

Public Property Get ElapsedTicks() As Currency 'As Long 'LongInt
    ElapsedTicks = GetRawElapsedTicks * CurCorrect
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = m_isRunning
End Property

' #################### '   Public Functions    ' #################### '
Public Function GetTimestamp() As Currency
    GetTimestamp = pGetTimestamp * CurCorrect
End Function

Public Function GetElapsedTime(ByVal startTimeStamp As Currency, Optional ByVal endTimeStamp As Currency = -1) As Currency
    If endTimeStamp < 0 Then endTimeStamp = GetTimestamp
    GetElapsedTime = endTimeStamp - startTimeStamp
End Function

Public Sub Start()
    If Not m_isRunning Then
        m_startTimeStamp = pGetTimestamp
        m_isRunning = True
    End If
End Sub

Public Function StartNew() As StopWatch
    Set StartNew = New StopWatch
    StartNew.Start
End Function

Public Sub SStop()
    If m_isRunning Then
        Dim pts As Currency: pts = pGetTimestamp
        Dim sts As Currency: sts = m_startTimeStamp
        Dim c As Currency: c = (pts - sts)
        m_elapsed = m_elapsed + c
        m_isRunning = False
    End If
End Sub

Public Sub Reset()
    m_elapsed = 0
    m_isRunning = False
    m_startTimeStamp = 0
End Sub

' #################### '   Private Functions    ' #################### '
Private Function pGetTimestamp() As Currency
    If m_isHighResolution Then
        QueryPerformanceCounter pGetTimestamp
    Else
        pGetTimestamp = Timer 'DateTime.Now 'or Timer
    End If
End Function

Private Function TimeSpanToString(ByVal ticks As Currency) As String
    Dim b As String
    Dim h As Long, m As Long, s As Long, n As Long
    Dim days As Long:  days = CLng((ticks / 86400000))
    Dim time As Currency: time = (ticks Mod 86400000)
    
    If (ticks < 0) Then
        b = b & "-"
        days = -days
        time = -time
    End If
    
    If (days <> 0) Then
        b = b & CStr(days) & "."
    End If
    
    'hours
    h = CInt(((time \ 3600000) Mod 24))
    b = b & IntToString(h, 2) & ":"
    
    'minutes
    m = CInt(((time \ 60000) Mod 60))
    b = b & IntToString(m, 2) & ":"
    
    'seconds
    s = CInt(((time \ 1000) Mod 60))
    b = b & IntToString(s, 2)
    
    n = (ticks - (CCur(h) * CCur(3600000)) _
               - (CCur(m) * CCur(60000)) _
               - (CCur(s) * CCur(1000))) * CurCorrect '10000
               
    If (n <> 0) Then
        b = b & "."
        b = b & IntToString(n, 7)
    End If
    TimeSpanToString = b
End Function

Private Function IntToString(ByVal n As Long, ByVal digits As Long) As String
    IntToString = CStr(n)
    Dim l As Long: l = Len(IntToString)
    If l < digits Then IntToString = String$(digits - l, "0") & IntToString
End Function

Private Function GetElapsedDateTimeTicks() As Currency
    Dim rawElapsedTicks As Currency: rawElapsedTicks = GetRawElapsedTicks
    If m_isHighResolution Then
        Dim d As Double: d = CDbl(rawElapsedTicks)
        d = (d * m_tickFrequency) ' m_tickFrequency = 1
        GetElapsedDateTimeTicks = CCur(d)
    Else
        GetElapsedDateTimeTicks = rawElapsedTicks
    End If
End Function

Private Function GetRawElapsedTicks() As Currency
    Dim Elapsed As Currency: Elapsed = m_elapsed
    If m_isRunning Then
        Dim ts As Currency: ts = pGetTimestamp
        Dim c As Currency: c = (ts - m_startTimeStamp)
        Elapsed = (Elapsed + c)
    End If
    GetRawElapsedTicks = Elapsed
End Function
