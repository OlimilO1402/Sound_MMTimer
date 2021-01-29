Attribute VB_Name = "MVirtualMem"
Option Explicit
' http://www.activevb.de/tipps/vb6tipps/tipp0101.html

Private Const MEM_COMMIT As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const MEM_RELEASE As Long = &H8000&

Public Declare Function VirtualAlloc Lib "kernel32.dll" ( _
                         ByRef lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
                         
Public Declare Function VirtualFree Lib "kernel32.dll" ( _
                         ByRef lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
                         
Public Declare Function VirtualLock Lib "kernel32.dll" ( _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
                
Public pAsm As Long

Public Function AddressOfAsmCallback(ByVal hWnd As Long) As Long
    Dim Arr(0 To 7) As Long
    Dim Tmp As Long
    ' push lParam
    ' push wParam
    ' push Msg
    ' push hWnd
    ' mov eax, SendMessage
    ' call eax
    ' ret 20
    Arr(0) = &H33221168
    Arr(1) = &H22116844
    Arr(2) = &H11684433
    Arr(3) = &H68443322
    Arr(4) = &H44332211
    Arr(5) = &H332211B8
    Arr(6) = &HC2D0FF44
    Arr(7) = &H14
    
    Tmp = 0       ' lParam setzen
    RtlMoveMemory ByVal CLng(VarPtr(Arr(0)) + 1), Tmp, Len(Tmp)
    
    Tmp = 18      ' wParam setzen; 18 ist Charcode für ein nicht darstellbares
                  ' Zeichen, das nun als ID für den Timer zweckentfremdet wird
    RtlMoveMemory ByVal CLng(VarPtr(Arr(1)) + 2), Tmp, Len(Tmp)
    
    Tmp = WM_CHAR ' Message; damit wird das KeyPress-Event des angegebenen
                  ' Fensters angesprungen
    RtlMoveMemory ByVal CLng(VarPtr(Arr(2)) + 3), Tmp, Len(Tmp)
    
    Tmp = hWnd    ' Fenster, an den die Timermessage geleitet wird
    RtlMoveMemory ByVal CLng(VarPtr(Arr(4)) + 0), Tmp, Len(Tmp)
    
                  ' Nachrichtenfunktion festlegen
    Tmp = GetProcAddress(GetModuleHandleA("user32.dll"), "PostMessageA")
    RtlMoveMemory ByVal CLng(VarPtr(Arr(5)) + 1), Tmp, Len(Tmp)
    
    ' verschieben in NX-kompatiblen Speicher
    pAsm = VirtualAlloc(ByVal 0, 32, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock ByVal pAsm, 32
    
    RtlMoveMemory ByVal pAsm, Arr(0), 32
    AddressOfAsmCallback = pAsm
End Function

Public Sub DeleteAsmCallback()
    VirtualFree ByVal pAsm, 0, MEM_RELEASE
End Sub

'oder auch:
'Option Explicit
'Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)
'
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
'
'Dim BytArr1(0 To 31) As Byte
'Dim BytArr2(0 To 31) As Byte
'
'Private Const WM_CHAR As Long = &H102
'
'Private ASM As Long
'
'Private Sub Form_Load()
'    PrepareBytArr1
'    PrepareBytArr2
'
'    Text1.Text = BytArr_ToHex(BytArr1)
'    Text2.Text = BytArr_ToHex(BytArr2)
'
'    Debug.Print BytArr_Compare(BytArr1, BytArr2)
'
'End Sub
'
'Sub PrepareBytArr1()
'    Dim LngArr(0 To 7) As Long
'    Dim Tmp As Long
'    ' push lParam
'    ' push wParam
'    ' push Msg
'    ' push hWnd
'    ' mov eax, SendMessage
'    ' call eax
'    ' ret 20
'    LngArr(0) = &H33221168
'    LngArr(1) = &H22116844
'    LngArr(2) = &H11684433
'    LngArr(3) = &H68443322
'    LngArr(4) = &H44332211
'    LngArr(5) = &H332211B8
'    LngArr(6) = &HC2D0FF44
'    LngArr(7) = &H14
'
'
'    Tmp = 0 ' lParam setzen
'    RtlMoveMemory ByVal CLng(VarPtr(LngArr(0)) + 1), Tmp, Len(Tmp)
'
'    Tmp = 18 ' wParam setzen; 18 ist Charcode für ein nicht darstellbares
'             ' Zeichen, das nun als ID für den Timer zweckentfremdet wird
'    RtlMoveMemory ByVal CLng(VarPtr(LngArr(1)) + 2), Tmp, Len(Tmp)
'
'    Tmp = WM_CHAR ' Message; damit wird das KeyPress-Event des angegebenen
'                  ' Fensters angesprungen
'    RtlMoveMemory ByVal CLng(VarPtr(LngArr(2)) + 3), Tmp, Len(Tmp)
'
'    Tmp = Picture1.hWnd ' Fenster, an den die Timermessage geleitet wird
'    RtlMoveMemory ByVal CLng(VarPtr(LngArr(4)) + 0), Tmp, Len(Tmp)
'
'    ' Nachrichtenfunktion festlegen
'    Tmp = GetProcAddress(GetModuleHandleA("user32.dll"), "PostMessageA")
'    RtlMoveMemory ByVal CLng(VarPtr(LngArr(5)) + 1), Tmp, Len(Tmp)
'
'    RtlMoveMemory BytArr1(0), LngArr(0), 32
'End Sub
'
'
Function GetAsm(ByVal aHwnd As Long) As Byte()

    Dim i As Long
    ReDim BytArr(0 To 31)
    
    ' push lParam
    BytArr(i) = &H68:                              i = i + 1

    ' lParam setzen
    Dim lParam As Long ': lParam = 0
    RtlMoveMemory BytArr(i), lParam, LenB(lParam): i = i + 4

    ' push wParam
    BytArr(i) = &H68:                              i = i + 1

    ' wParam setzen; 18 ist Charcode für ein nicht darstellbares Zeichen, das nun als ID für den Timer zweckentfremdet wird
    Dim wParam As Long: wParam = &H12 ' = ""
    RtlMoveMemory BytArr(i), wParam, LenB(wParam): i = i + 4

    ' push wMsg
    BytArr(i) = &H68:                              i = i + 1

    ' Message WM_CHAR; damit wird das KeyPress-Event des angegebenen Fensters angesprungen
    Dim wMsg   As Long: wMsg = WM_CHAR
    RtlMoveMemory BytArr(i), wMsg, LenB(wMsg):     i = i + 4

    ' push hWnd
    BytArr(i) = &H68:                              i = i + 1

    ' das Fensterhandle an das die Message gesendet wird
    Dim hWnd   As Long: hWnd = aHwnd
    RtlMoveMemory BytArr(i), hWnd, LenB(hWnd):     i = i + 4

    ' mov eax, PostMessage
    BytArr(i) = &HB8:                              i = i + 1

    ' Zeiger auf die Funktion ermitteln die aufgerufen werden soll, hier PostMessage
    Dim pFnc   As Long: pFnc = GetProcAddress(GetModuleHandleA("user32.dll"), "PostMessageA")
    RtlMoveMemory BytArr(i), pFnc, LenB(pFnc):     i = i + 4

    'call eax
    BytArr(i) = &HFF:                              i = i + 1
    BytArr(i) = &HD0:                              i = i + 1
    BytArr(i) = &HC2:                              i = i + 1

    'ret 20
    BytArr(i) = &H14

End Function
'
'Function BytArr_Compare(BArr1() As Byte, BArr2() As Byte) As Long
''retval  condition
''  0     if equal
'' -1     if barr1<barr2
''  1     if barr2<barr1
'
'    Dim u1 As Long: u1 = UBound(BArr1)
'    Dim u2 As Long: u2 = UBound(BArr2)
'    If u1 <> u2 Then
'        BytArr_Compare = IIf(u1 < u2, -1, 1)
'        Exit Function
'    End If
'    Dim i As Long, u As Long: u = u1
'    For i = 0 To u
'        If BArr1(i) <> BArr2(i) Then
'            BytArr_Compare = IIf(BArr1(i) < BArr2(i), -1, 1)
'            Debug.Print i & ": &H" & Hex(BArr1(i)) & " &H" & Hex(BArr2(i))
'            'Exit Function
'        End If
'    Next
'End Function
'
'Function BytArr_ToHex(BArr() As Byte) As String
'    Dim i As Long
'    Dim s As String: s = " " & Hex2(BArr(i))
'    For i = 1 To UBound(BArr)
'        s = s & " " & Hex2(BArr(i))
'        If (i + 1) Mod 5 = 0 Then s = s & vbCrLf
'    Next
'    BytArr_ToHex = s
'End Function
'
'Function Hex2(b As Byte) As String
'    Hex2 = Hex(b)
'    If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
'End Function
'
