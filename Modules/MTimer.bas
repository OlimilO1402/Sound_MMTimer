Attribute VB_Name = "MTimer"
Option Explicit
Declare Function timeSetEvent Lib "winmm.dll" ( _
    ByVal uDelay As Long, _
    ByVal uResolution As Long, _
    ByVal lpFunction As Long, _
    ByVal dwUser As Long, _
    ByVal uFlags As Long) As Long

Declare Function timeKillEvent Lib "winmm.dll" ( _
    ByVal uID As Long) As Long
        
'no! Do not Declare here in VB-Module directly, use tlb instead
'In tlb you have to use usesgetlasterror=off
#If withoutTlb Then
'not the way to go, do not do this at all in your projects!
Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" ( _
                 ByVal hwnd As Long, _
                 ByVal wMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long) As Long
#End If

Public Const TIME_ONESHOT  As Long = 0&
Public Const TIME_PERIODIC As Long = 1&
Public Const TIME_CALLBACK_FUNCTION    As Long = &H0&
Public Const TIME_CALLBACK_EVENT_SET   As Long = &H10&
Public Const TIME_CALLBACK_EVENT_PULSE As Long = &H20&
Public Const TIME_KILL_SYNCHRONOUS     As Long = &H100&

Public Type TimerHandle
    TimerID   As Long
    hwnd      As Long
    Counter   As Long
    Interval  As Long
    mMod      As Long
    StartTime As Single
    IsRunning As Boolean
End Type
Private TimerHandleEmpty As TimerHandle
Private Const C_MaxTimerCount As Long = 32
Public Timers(0 To C_MaxTimerCount - 1) As TimerHandle
Private m_CountTimers As Long
'
Private Const WM_LBUTTONDOWN As Long = &H201

Private Const WM_CHAR        As Long = &H102

Public Function StartTimer(ByVal hwnd As Long, ByVal Interval As Long, Optional ByVal iTimer As Long = -1) As Long
    If m_CountTimers > C_MaxTimerCount Then Exit Function
    Dim b As Boolean: b = iTimer < 0
    If b Then iTimer = m_CountTimers
    With Timers(iTimer)
        .Interval = Interval
        .hwnd = hwnd
        If Interval <> 0 Then .mMod = 1000 \ Interval
        .StartTime = Timer
        .IsRunning = True
        .TimerID = timeSetEvent(Interval, _
                                0, _
                                AddressOf TimerCallback, _
                                iTimer, _
                                TIME_PERIODIC Or TIME_KILL_SYNCHRONOUS)
    End With
    StartTimer = iTimer
    If b Then m_CountTimers = m_CountTimers + 1
End Function

Private Sub TimerCallback(ByVal ID As Long, _
                          ByVal Msg As Long, _
                          ByVal LngUser As Long, _
                          ByVal Lng1 As Long, _
                          ByVal Lng2 As Long)
    With Timers(LngUser)
        .Counter = .Counter + 1
        PostMessage .hwnd, WM_CHAR, &H0, 0
    End With

End Sub
Public Sub StopTimer(ByVal Index As Long)
    With Timers(Index)
        If .TimerID >= 0 Then
            Dim hr As Long
            hr = timeKillEvent(.TimerID)
            .IsRunning = False
        End If
    End With
End Sub
Public Sub ResetTimer(ByVal Index As Long)
    If Timers(Index).IsRunning Then
        StopTimer Index
    End If
    Timers(Index) = TimerHandleEmpty
End Sub
