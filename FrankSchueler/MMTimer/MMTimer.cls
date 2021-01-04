VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MMTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----==== Const ====----
Private Const MEM_COMMIT As Long = &H1000&
Private Const PAGE_READWRITE As Long = &H4&
Private Const PAGE_EXECUTE_READ As Long = &H20&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
Private Const MEM_RELEASE As Long = &H8000&
Private Const GWL_WNDPROC As Long = (-4&)
Private Const HWND_MESSAGE As Long = (-3&)
Private Const WM_USER As Long = &H400&
Private Const WM_DESTROY As Long = &H2&
Private Const TIME_ONESHOT As Long = &H0&
Private Const TIME_PERIODIC As Long = &H1&

' ----==== Kernel32.dll Deklaration ====----
Private Declare Sub CopyMemory Lib "kernel32.dll" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Private Declare Function VirtualAlloc Lib "kernel32.dll" ( _
                         ByRef lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
                         
Private Declare Function VirtualProtect Lib "kernel32" ( _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long, _
                         ByVal flNewProtect As Long, _
                         ByRef lpflOldProtect As Long) As Long
                         
Private Declare Function VirtualFree Lib "kernel32.dll" ( _
                         ByRef lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
                         
Private Declare Function VirtualLock Lib "kernel32.dll" ( _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long) As Long
                         
Private Declare Function GetProcAddress Lib "kernel32.dll" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
                         
Private Declare Function GetModuleHandle Lib "kernel32.dll" _
                         Alias "GetModuleHandleA" ( _
                         ByVal lpModuleName As String) As Long


' ----==== User32.dll Deklaration ====----
Private Declare Function CreateWindowEx Lib "user32.dll" _
                         Alias "CreateWindowExW" ( _
                         ByVal dwExStyle As Long, _
                         ByVal lpClassName As Long, _
                         ByVal lpWindowName As Long, _
                         ByVal dwStyle As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal nWidth As Long, _
                         ByVal nHeight As Long, _
                         ByVal hWndParent As Long, _
                         ByVal hMenu As Long, _
                         ByVal hInstance As Long, _
                         ByRef lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32.dll" ( _
                         ByVal hwnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32.dll" _
                         Alias "GetWindowLongW" ( _
                         ByVal hwnd As Long, _
                         ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" _
                         Alias "SetWindowLongW" ( _
                         ByVal hwnd As Long, _
                         ByVal nIndex As Long, _
                         ByVal dwNewLong As Long) As Long


' ----==== Winmm.dll Deklaration ====----
Private Declare Function timeKillEvent Lib "winmm.dll" ( _
                         ByVal uID As Long) As Long

Private Declare Function timeSetEvent Lib "winmm.dll" ( _
                         ByVal uDelay As Long, _
                         ByVal uResolution As Long, _
                         ByVal lpFunction As Long, _
                         ByVal dwUser As Long, _
                         ByVal uFlags As Long) As Long

' ----==== Variablen ====----
Private m_hMsgWnd As Long
Private m_hPrevProc As Long
Private m_pASM As Long
Private m_pASMWrapper As Long
Private m_TimerID As Long
Private m_Interval As Long
Private m_OneShot As Boolean
Private m_Enabled As Boolean

Public Event Timer()

' Callback-Funktion,
Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If Msg = WM_USER + 4711 Then

        RaiseEvent Timer

    End If

'    If Msg = WM_DESTROY Then
'
'        Call StopTimer
'
'    End If

    WindowProc = 0&

End Function

Public Property Get Interval() As Long
    
    Interval = m_Interval

End Property

Public Property Let Interval(ByVal Value As Long)
    
    If Value < 0& Then Value = 0&
    
    m_Interval = Value
    
    If m_Interval = 0& Then
    
        Call StopTimer
        
    Else
    
        If m_Enabled = True Then
        
            Call StartTimer
        
        End If
    
    End If
    
End Property

Public Property Get OneShot() As Boolean
    
    OneShot = m_OneShot

End Property

Public Property Let OneShot(ByVal Value As Boolean)

    m_OneShot = Value
    
End Property

Public Property Get Enabled() As Boolean
    
    Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal Value As Boolean)

    m_Enabled = Value
    
    If m_Enabled = True Then
    
        Call StartTimer
        
    Else
    
        Call StopTimer
    
    End If

End Property

Private Sub StartTimer()

    Call StopTimer

    If m_hMsgWnd <> 0& Then

        m_hPrevProc = SetWindowLong(m_hMsgWnd, GWL_WNDPROC, m_pASMWrapper)
        
        If m_TimerID = 0& Then
    
            If m_Interval > 0& Then

                If m_OneShot = True Then

                    m_TimerID = timeSetEvent(m_Interval, 1&, m_pASM, 0&, TIME_ONESHOT)

                Else

                    m_TimerID = timeSetEvent(m_Interval, 1&, m_pASM, 0&, TIME_PERIODIC)

                End If
                
            End If
            
        End If

    End If

End Sub

Private Sub StopTimer()

    If m_hMsgWnd <> 0 Then

        Call SetWindowLong(m_hMsgWnd, GWL_WNDPROC, m_hPrevProc)

    End If

    If m_TimerID <> 0& Then
    
        If timeKillEvent(m_TimerID) <> 0& Then

            m_TimerID = 0&
            
            If m_OneShot = True Then
            
                m_Enabled = False
            
            End If
            
        End If
    
    End If

End Sub

Private Sub Class_Initialize()

    Dim lAsm(0 To 7) As Long
    Dim bAsm(0 To 103) As Byte
    Dim pVar As Long
    Dim ThisClass As Long
    Dim CallbackFunction As Long
    Dim pVirtualFree As Long

    m_TimerID = 0&
    m_Enabled = False
    m_OneShot = False
    m_Interval = 1000&

    m_hMsgWnd = CreateWindowEx(0&, StrPtr("Static"), 0&, 0&, 0&, 0&, 0&, 0&, HWND_MESSAGE, 0&, 0&, ByVal 0&)

    If m_hMsgWnd <> 0 Then

        ' // --- Asm-Code aufbauen ---
        ' //lea eax,[esp+4)    Pointer auf ersten Funktionsparameter ermitteln und als...
        ' //push eax           LPARAM auf den Stack schieben
        ' //push 0             WPARAM auf den Stack schieben
        ' //push 0             Message auf den Stack schieben
        ' //push 0             HWND auf den Stack schieben
        ' //mov eax,0          Function-Pointer auf SendMessage in Register schieben und...
        ' //call eax           SendMessage-Funktion aufrufen
        ' //ret 20             Asm-Funktion verlassen und Stack aufräumen

        ' &H0424448D
        lAsm(0) = &H424448D
        ' &H00006850
        lAsm(1) = &H6850
        ' &H00680000
        lAsm(2) = &H680000
        ' &H68000000
        lAsm(3) = &H68000000
        ' &H00000000
        lAsm(4) = &H0
        ' &H000000B8
        lAsm(5) = &HB8
        ' &HC2D0FF00
        lAsm(6) = &HC2D0FF00
        ' &H00000014
        lAsm(7) = &H14

        ' //WPARAM. Kann beliebig gewählt werden
        Dim lWParam As Long

        lWParam = 0
        Call CopyMemory(ByVal CLng(VarPtr(lAsm(1)) + 2), lWParam, Len(lWParam))

        ' //Beliebige, zu verwendende Window-Message
        Dim lMessage As Long

        lMessage = WM_USER + 4711
        Call CopyMemory(ByVal CLng(VarPtr(lAsm(2)) + 3), lMessage, Len(lMessage))

        ' //Fensterhandle des Zielfensters
        Dim lHWnd As Long

        lHWnd = m_hMsgWnd
        Call CopyMemory(ByVal CLng(VarPtr(lAsm(4)) + 0), lHWnd, Len(lHWnd))

        ' //Adresse von SendMessage
        Dim lRet As Long

        'lRet = GetProcAddress(GetModuleHandle("user32.dll"), "SendMessageW")
        lRet = GetProcAddress(GetModuleHandle("user32.dll"), "PostMessageW")
        Call CopyMemory(ByVal CLng(VarPtr(lAsm(5)) + 1), lRet, Len(lRet))

        m_pASM = VirtualAlloc(ByVal 0, 32, MEM_COMMIT, PAGE_READWRITE)
        Call VirtualLock(ByVal m_pASM, 32)
        Call CopyMemory(ByVal m_pASM, lAsm(0), 32)

        ' //--- Memory-Protection setzen ---
        Dim old As Long

        Call VirtualProtect(ByVal m_pASM, 32, PAGE_EXECUTE_READ, old)
        
        ' Virtuellen Speicher anfordern
        m_pASMWrapper = VirtualAlloc(ByVal 0&, 104, MEM_COMMIT, PAGE_EXECUTE_READWRITE)

        If m_pASMWrapper <> 0 Then

            ' Instanzzeiger der Klasse auslesen
            ThisClass = ObjPtr(Me)

            ' Zeiger auf die Callback-Funktion auslesen
            Call CopyMemory(pVar, ByVal ThisClass, 4)
            Call CopyMemory(CallbackFunction, ByVal (pVar + 28), 4)

            ' Zeiger auf die VirtualFree-Funktion ermitteln
            pVirtualFree = GetProcAddress(GetModuleHandle("kernel32.dll"), "VirtualFree")

            ' ASM-Wrapper mit Maschinencode befüllen
            ' &Hcc int 3 (Software Interrupt zum debuggen), &H90=nop (No Operation Point)
            bAsm(0) = &H90

            ' inc (Zähler)
            bAsm(1) = &HFF
            bAsm(2) = &H5

            ' push 0
            bAsm(7) = &H6A
            bAsm(8) = &H0

            ' push esp
            bAsm(9) = &H54

            ' push (esp+18h) (laram)
            bAsm(10) = &HFF
            bAsm(11) = &H74
            bAsm(12) = &H24
            bAsm(13) = &H18

            ' push (esp+18h) (wParam)
            bAsm(14) = &HFF
            bAsm(15) = &H74
            bAsm(16) = &H24
            bAsm(17) = &H18

            ' push (esp+18h) (msg)
            bAsm(18) = &HFF
            bAsm(19) = &H74
            bAsm(20) = &H24
            bAsm(21) = &H18

            ' push (esp+18h) (hwnd)
            bAsm(22) = &HFF
            bAsm(23) = &H74
            bAsm(24) = &H24
            bAsm(25) = &H18

            ' push Instanzzeiger
            bAsm(26) = &H68

            ' mov eax, Adresse WindowProc
            bAsm(31) = &HB8

            ' call eax
            bAsm(36) = &HFF
            bAsm(37) = &HD0

            ' dec (Zähler)
            bAsm(38) = &HFF
            bAsm(39) = &HD

            ' mov eax, (Signal)
            bAsm(44) = &HA1

            ' test eax, eax
            bAsm(49) = &H85
            bAsm(50) = &HC0

            ' jne
            bAsm(51) = &H75
            bAsm(52) = &H4

            ' pop eax (Rückgabewert)
            bAsm(53) = &H58

            ' ret &H10
            bAsm(54) = &HC2
            bAsm(55) = &H10
            bAsm(56) = &H0

            ' mov eax, (Zähler)
            bAsm(57) = &HA1

            ' test eax, eax
            bAsm(62) = &H85
            bAsm(63) = &HC0

            ' je
            bAsm(64) = &H74
            bAsm(65) = &H4

            ' pop eax (Rückgabewert)
            bAsm(66) = &H58

            ' ret &H10
            bAsm(67) = &HC2
            bAsm(68) = &H10
            bAsm(69) = &H0

            ' pop eax retval
            bAsm(70) = &H58

            ' pop ecx (Rücksprungzeiger)
            bAsm(71) = &H59

            ' pop eax hwnd
            bAsm(72) = &H58

            ' pop eax msg
            bAsm(73) = &H58

            ' pop eax wparam
            bAsm(74) = &H58

            ' pop eax lparam
            bAsm(75) = &H58

            ' push MEM_RELEASE
            bAsm(76) = &H68
            bAsm(77) = &H0
            bAsm(78) = &H80
            bAsm(79) = &H0
            bAsm(80) = &H0

            ' push 0
            bAsm(81) = &H6A
            bAsm(82) = &H0

            ' push Zeiger auf den Wrapper
            bAsm(83) = &H68

            ' push ecx (Rücksprungzeiger)
            bAsm(88) = &H51

            ' mov eax, VirtualFree Adresse
            bAsm(89) = &HB8

            ' jmp eax
            bAsm(94) = &HFF
            bAsm(95) = &HE0

            ' Speicher für Zähler
            bAsm(96) = &H0
            bAsm(97) = &H0
            bAsm(98) = &H0
            bAsm(99) = &H0

            ' Speicher für Signal
            bAsm(100) = &H0
            bAsm(101) = &H0
            bAsm(102) = &H0
            bAsm(103) = &H0

            ' Zähler Variable setzen
            pVar = m_pASMWrapper + 96

            Call CopyMemory(bAsm(3), pVar, 4)
            Call CopyMemory(bAsm(40), pVar, 4)
            Call CopyMemory(bAsm(58), pVar, 4)

            ' Flag Variable setzen
            pVar = m_pASMWrapper + 100
            Call CopyMemory(bAsm(45), pVar, 4)

            ' Wrapper Adresse setzen
            pVar = m_pASMWrapper
            Call CopyMemory(bAsm(84), pVar, 4)

            ' Instanzzeiger setzen
            pVar = ThisClass
            Call CopyMemory(bAsm(27), pVar, 4)

            ' Funktionszeiger setzen
            pVar = CallbackFunction
            Call CopyMemory(bAsm(32), pVar, 4)

            ' VirtualFree Adresse setzen
            pVar = pVirtualFree
            Call CopyMemory(bAsm(90), pVar, 4)

            ' fertigen Wrapper in DEP-kompatiblen Speicher kopieren
            Call CopyMemory(ByVal m_pASMWrapper, bAsm(0), 104)

        End If

    End If

End Sub

Private Sub Class_Terminate()

    Dim lCounter As Long
    Dim lFlag As Long

    Call StopTimer

    If m_pASMWrapper <> 0& Then

        If m_hMsgWnd <> 0& Then

            Call SetWindowLong(m_hMsgWnd, GWL_WNDPROC, m_hPrevProc)

            ' Zähler auslesen
            Call CopyMemory(lCounter, ByVal (m_pASMWrapper + 104), 4&)

            If lCounter = 0& Then

                ' Wrapper kann von VB aus gelöscht werden

                Call VirtualFree(ByVal m_pASMWrapper, 0&, MEM_RELEASE)

            Else

                ' Wrapper befindet sich noch innerhalb einer Rekursion und muss sich
                ' selbst löschen Flag setzen

                lFlag = 1&

                Call CopyMemory(ByVal (m_pASMWrapper + 108), lFlag, 4&)

            End If

            Call VirtualFree(ByVal m_pASM, 0&, MEM_RELEASE)

            Call DestroyWindow(m_hMsgWnd)

        End If

    End If

End Sub

