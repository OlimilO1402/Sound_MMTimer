VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MMTimer"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnStartAllTimers 
      Caption         =   "StartAllTimers"
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Text            =   "5"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnReset2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Text            =   "1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnReset1 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton BtnStartStop2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1485
         TabIndex        =   4
         Top             =   120
         Width           =   90
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1485
         TabIndex        =   2
         Top             =   120
         Width           =   90
      End
   End
   Begin VB.CommandButton BtnStartStop1 
      Caption         =   "Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label LblTimerID2 
      Caption         =   "Label3"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label LblTimerID1 
      Caption         =   "Label3"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label LblTimerIndex2 
      Caption         =   "Label4"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label LblTimerIndex1 
      Caption         =   "Label3"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Interval [ms]:"
      Height          =   195
      Left            =   2760
      TabIndex        =   15
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Tim2Lbl0 
      Caption         =   "0"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Tim2Lbl1 
      Caption         =   "0"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Interval [ms]:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Tim1Lbl1 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Tim1Lbl0 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Works safely only compiled
Private m_TimerIndex1 As Long
Private m_TimerIndex2 As Long


Private Sub Form_Load()

#If withoutTlb Then
    Me.Caption = "MMTimer with PostMessage declared in Module MTimer.bas"
#ElseIf withTlb Then
    Me.Caption = "MMTimer with PostMessage in .tlb (usesgetlasterror=off)"
#ElseIf withAsmFnc Then
    Me.Caption = "MMTimer with PostMessage in VirtualMemory Asm-function"
#End If
    m_TimerIndex1 = -1
    m_TimerIndex2 = -1
    BtnStartStop1.Caption = "Start"
    BtnStartStop2.Caption = "Start"
End Sub

Private Sub Form_Unload(Cancel As Integer)
#If withAsmFnc Then
    MVirtualMem.DeleteAsmCallback
#End If
End Sub

Private Sub BtnStartAllTimers_Click()
    'BtnStartStop1.Value = True
    BtnStartStop1_Click
    'BtnStartStop2.Value = True
    BtnStartStop2_Click
End Sub

Private Sub BtnStartStop1_Click()
    m_TimerIndex1 = OnBtnStartStop(BtnStartStop1, Text1, Picture1, Tim1Lbl0, LblTimerIndex1, LblTimerID1, m_TimerIndex1)
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    OnKeyPress MTimer.Timers(m_TimerIndex1), Label1, Tim1Lbl1
End Sub
Private Sub BtnReset1_Click()
    OnReset m_TimerIndex1, Label1, Tim1Lbl0, Tim1Lbl1
End Sub


Private Sub BtnStartStop2_Click()
    m_TimerIndex2 = OnBtnStartStop(BtnStartStop2, Text2, Picture2, Tim2Lbl0, LblTimerIndex2, LblTimerID2, m_TimerIndex2)
End Sub
Private Sub Picture2_KeyPress(KeyAscii As Integer)
    OnKeyPress MTimer.Timers(m_TimerIndex2), Label2, Tim2Lbl1
End Sub
Private Sub BtnReset2_Click()
    OnReset m_TimerIndex2, Label2, Tim2Lbl0, Tim2Lbl1
End Sub

Private Function OnBtnStartStop(BtnStartStop As CommandButton, Txt As TextBox, PB As PictureBox, LblDT As Label, LblI As Label, LblID As Label, ByVal TimerIndex_in As Long) As Long
    If BtnStartStop.Caption = "Start" Then
        BtnStartStop.Caption = "Stop"
        Dim Interval As Long: Interval = Txt.Text
        LblDT.Caption = Timer
        OnBtnStartStop = MTimer.StartTimer(PB.hWnd, Interval, TimerIndex_in)
        LblI.Caption = OnBtnStartStop
        LblID.Caption = MTimer.Timers(OnBtnStartStop).TimerID
    Else
        MTimer.StopTimer TimerIndex_in
        OnBtnStartStop = TimerIndex_in
        BtnStartStop.Caption = "Start"
    End If
End Function
Private Sub OnKeyPress(tim As TimerHandle, Lbl As Label, TimLbl As Label)
    With tim
        Lbl.Caption = CStr(.Counter)
        If .mMod <> 0 Then
            If .Counter Mod .mMod = 0 Then
                TimLbl.Caption = Timer - .StartTime
            End If
        End If
    End With
End Sub
Private Sub OnReset(TimerIndex As Long, Lbl As Label, TimLbl0 As Label, TimLbl1 As Label)
    MTimer.ResetTimer TimerIndex
    Lbl.Caption = "0"
    TimLbl0.Caption = "0"
    TimLbl1.Caption = "0"
End Sub

'    With Timers(m_TimerIndex1)
'        Label2.Caption = CStr(.Counter)
'        If .Counter Mod .mMod = 0 Then
'            Tim2Lbl1.Caption = Timer - d2
'        End If
'    End With

