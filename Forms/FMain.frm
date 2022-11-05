VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Stopwatch"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   LinkTopic       =   "FMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSplittime 
      Caption         =   "Splittime"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton BtnReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton BtnStartStop 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label LblRuntime 
      Alignment       =   2  'Zentriert
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label LblSplittime 
      Alignment       =   2  'Zentriert
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   3975
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Tester for the stopwatch-class
Private mSW As StopWatch
Private mdt As Double

Private Sub Form_Load()
    Me.Caption = "Stopwatch v" & App.Major & "." & App.Minor & "." & App.Revision
    Set mSW = New StopWatch
End Sub

Private Sub BtnTest1_Click()
    Dim z As Long, v As Double
    mSW.Start
    For z = 1 To 10 ^ 7
        v = Rnd(1)
    Next
    MsgBox mSW.ElapsedToString
    For z = 1 To 10 ^ 7
        v = Rnd(1)
    Next
    MsgBox mSW.ElapsedToString
    For z = 1 To 10 ^ 7
        v = Rnd(1)
    Next
    mSW.SStop
    MsgBox mSW.ElapsedToString
End Sub

Private Sub BtnStartStop_Click()
    Timer1.Enabled = Not Timer1.Enabled
    If Timer1.Enabled Then
        mSW.Start
        BtnStartStop.Caption = "Stop"
        BtnReset.Enabled = False
    Else
        mSW.SStop
        BtnStartStop.Caption = "Start"
        BtnReset.Enabled = True
        LblRuntime.Caption = mSW.ElapsedToString
        Dim s As String
        s = s & "mSW.ElapsedTicks       : " & mSW.ElapsedTicks & vbCrLf
        s = s & "mSW.ElapsedMilliseconds: " & mSW.ElapsedMilliseconds & vbCrLf
        s = s & "StopWatch.Frequency    : " & StopWatch.Frequency & vbCrLf
        s = s & "Stopwatch.GetTimestamp : " & StopWatch.GetTimestamp & vbCrLf
        Label1.Caption = s
    End If
End Sub

Private Sub BtnReset_Click()
    mSW.Reset
    mdt = 0#
    Label1.Caption = ""
End Sub

Private Sub BtnSplittime_Click()
    Dim dt As Double: dt = mdt
    mdt = (CDbl(mSW.Elapsed) / 1000)
    dt = mdt - dt
    LblSplittime.Caption = mSW.ElapsedToString & vbCrLf & "dt: " & Format(dt, "0.0######### s")
End Sub

Private Sub Timer1_Timer()
    LblRuntime.Caption = mSW.ElapsedToString
End Sub
