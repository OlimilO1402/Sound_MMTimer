Attribute VB_Name = "MTimer"
Option Explicit
'https://docs.microsoft.com/en-us/previous-versions/dd757634(v=vs.85)
Declare Function timeSetEvent Lib "winmm.dll" ( _
    ByVal uDelay As Long, _
    ByVal uResolution As Long, _
    ByVal lpFunction As Long, _
    ByVal dwUser As Long, _
    ByVal uFlags As Long) As Long

'https://docs.microsoft.com/en-us/previous-versions/dd757630(v=vs.85)
Declare Function timeKillEvent Lib "winmm.dll" ( _
    ByVal uID As Long) As Long
        
'no! Do not Declare here in VB-Module directly, use tlb instead
'In tlb you have to use usesgetlasterror=off
#If withoutTlb Then
'not the way to go, do not do this at all in your projects!

    'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea
    'Places (posts) a message in the message queue associated with the thread that created the specified window and returns without waiting for the thread to process the message.
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'SendMessage in contrary
    'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessage
    'Sends the specified message to a window or windows. The SendMessage function calls the window procedure for the specified window and does not return until the window procedure has processed the message.
    Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
#End If

Public Const TIME_ONESHOT  As Long = &H0&
Public Const TIME_PERIODIC As Long = &H1&
Public Const TIME_CALLBACK_FUNCTION    As Long = &H0&
Public Const TIME_CALLBACK_EVENT_SET   As Long = &H10&
Public Const TIME_CALLBACK_EVENT_PULSE As Long = &H20&
Public Const TIME_KILL_SYNCHRONOUS     As Long = &H100&

Public Const TIMERR_NOERROR      As Long = &H0&
Public Const MMSYSERR_BASE       As Long = &H0&
Public Const MMSYSERR_INVALPARAM As Long = (MMSYSERR_BASE + 11)

'TODO: trennen von Timer-zeugs und allem anderen
'Public Type TimerHnd
'  hWnd      As Long 'Window-Handle of the control to send the Message
'  TimerID   As Long
'  Enabled   As Boolean ' = IsRunning
'  Interval  As Long    '
'  StartTime As Single  '
'End Type 'Timer
'Public Type MMTimer
'    Hnd     As TimerHnd
'    Counter As Long
'    mMod    As Long
'End Type
Public Type TimerHandle
    TimerID   As Long
    hWnd      As Long
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

Public Const WM_CHAR        As Long = &H102

Public Function StartTimer(ByVal hWnd As Long, ByVal Interval As Long, Optional ByVal iTimer As Long = -1) As Long
    If m_CountTimers > C_MaxTimerCount Then Exit Function
    Dim b As Boolean: b = iTimer < 0
    If b Then iTimer = m_CountTimers
    With Timers(iTimer)
        .Interval = Interval
        .hWnd = hWnd
        If Interval <> 0 Then .mMod = 1000 \ Interval
        .StartTime = Timer
        .IsRunning = True
        .TimerID = timeSetEvent(Interval, _
                                0, _
                                AddressOf_TimerCallback(hWnd), _
                                iTimer, _
                                TIME_PERIODIC Or TIME_KILL_SYNCHRONOUS)
    End With
    StartTimer = iTimer
    If b Then m_CountTimers = m_CountTimers + 1
End Function
Private Function AddressOf_TimerCallback(ByVal hWnd As Long) As Long
#If withasmfnc Then
    AddressOf_TimerCallback = MVirtualMem.AddressOfAsmCallback(hWnd)
#Else
    AddressOf_TimerCallback = FncPtr(AddressOf TimerCallback)
#End If
End Function
'https://docs.microsoft.com/en-us/previous-versions//dd757631(v=vs.85)
'https://docs.microsoft.com/en-us/previous-versions/ff728861(v=vs.85)
Private Sub TimerCallback(ByVal ID As Long, _
                          ByVal Msg As Long, _
                          ByVal LngUser As Long, _
                          ByVal Lng1 As Long, _
                          ByVal Lng2 As Long)
    With Timers(LngUser)
        .Counter = .Counter + 1
        PostMessage .hWnd, WM_CHAR, &H0, 0
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

Public Function FncPtr(AddressOf_Fnc As Long) As Long
    FncPtr = AddressOf_Fnc
End Function
