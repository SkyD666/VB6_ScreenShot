Attribute VB_Name = "modSystemKeyHook"
' by CLE

' APIs
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const WH_MOUSE_LL = 14
Private Const WH_KEYBOARD_LL = 13
Private Const WH_MOUSE = 7
Private Const WH_KEYBOARD = 2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_MBUTTONUP = &H208
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105

Private Const VK_SHIFT As Byte = &H10
Private Const VK_CAPITAL As Byte = &H14
Private Const VK_NUMLOCK As Byte = &H90

Public Type Point
    x As Long
    y As Long
End Type

Private Type KeyboardHookStruct
    vkCode As Long
    ScanCode As Long
    flags As Long
    Time As Long
    DwExtraInfo As Long
End Type

Dim hKeyboardHook As Long
Public SpeKeycodename As String
Public FinalKeyName As String

Sub RegKeyHook()
    If hKeyboardHook = 0 Then
        hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardHookProcHookProc, App.hInstance, 0)
        If hKeyboardHook = 0 Then
            ' ���ﴦ��ע�����
            MsgBox "ע��ϵͳ���̹���ʧ�ܣ�"
        Else
            'MsgBox "ע����̹�����ɣ�"
        End If
    End If
End Sub

Sub UnKeyHook()
    If hKeyboardHook <> 0 Then
        Dim retKeyboard As Long
        retKeyboard = UnhookWindowsHookEx(hKeyboardHook)
        hKeyboardHook = 0
        If retKeyboard = 0 Then
            ' ���ﴦ��ж�ش���
            MsgBox "ж��ϵͳ���̹���ʧ�ܣ�"
        End If
    End If
End Sub

' �ص�����
Private Function KeyboardHookProcHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode >= 0 Then
        Dim ks As KeyboardHookStruct
        CopyMemory ks, ByVal lParam, 20
        
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
            ' ���ﴦ����̰��µ��¼�
            
            If ks.vkCode = HotKeyCodeInt And frmMain.Enabled = True Then HotKeyPressedBoo = True: frmMain.timerHotKey.Enabled = True
        End If
        
        If wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            ' ���ﴦ����̵����¼�
            'frmMain.key_up ks.vkCode
        End If
        
        ' ��Ҫ�ػ񰴼��¼�������ֱ������ KeyboardHookProcHookProc = 1
        ' ���������һ������
        CallNextHookEx hKeyboardHook, nCode, wParam, lParam
    End If
End Function

'Sub key_down(ByVal code As Long)
'    Select Case code
'    Case 8
'        SpeKeycodename = "Backspace��"
'
'    Case 9
'        SpeKeycodename = "Tab��"
'
'    Case 13
'        SpeKeycodename = "�س���"
'
'    Case 19
'        SpeKeycodename = "Pause��"
'
'    Case 20
'        SpeKeycodename = "CapsLock��"
'
'    Case 27
'        SpeKeycodename = "ESC��"
'
'    Case 32
'        SpeKeycodename = "�ո��"
'
'    Case 33
'        SpeKeycodename = "PageUp��"
'
'    Case 34
'        SpeKeycodename = "PageDown��"
'
'    Case 35
'        SpeKeycodename = "End��"
'
'    Case 36
'        SpeKeycodename = "Home��"
'
'    Case 37
'        SpeKeycodename = "�����(��)"
'
'    Case 38
'        SpeKeycodename = "�����(��)"
'
'    Case 39
'        SpeKeycodename = "�����(��)"
'
'    Case 40
'        SpeKeycodename = "�����(��)"
'
'    Case 44
'        SpeKeycodename = "PrtSc��"
'
'    Case 45
'        SpeKeycodename = "Insert��"
'
'    Case 46
'        SpeKeycodename = "Delete��"
'
'    Case 93
'        SpeKeycodename = "�˵���(appskey)"
'
'    Case 106
'        SpeKeycodename = "*��"
'
'    Case 107
'        SpeKeycodename = "+��"
'
'    Case 110
'        SpeKeycodename = ".��"
'
'    Case 144
'        SpeKeycodename = "NumLK��"
'
'    Case 145
'        SpeKeycodename = "ScrLK��"
'
'    Case 160
'        SpeKeycodename = "��Shift��"
'
'    Case 161
'        SpeKeycodename = "��Shift��"
'
'    Case 162
'        SpeKeycodename = "��Ctrl��"
'
'    Case 163
'        SpeKeycodename = "��Ctrl��"
'
'    Case 164
'        SpeKeycodename = "��Alt��"
'
'    Case 165
'        SpeKeycodename = "��Alt��"
'
'    Case 189
'        SpeKeycodename = "-��"
'
'    Case 109
'        SpeKeycodename = "-��"
'
'    Case 187
'        SpeKeycodename = "=��"
'
'    Case 192
'        SpeKeycodename = "`��"
'
'    Case 219
'        SpeKeycodename = "[��"
'
'    Case 221
'        SpeKeycodename = "]��"
'
'    Case 186
'        SpeKeycodename = ";��"
'
'    Case 222
'        SpeKeycodename = "'��"
'
'    Case 220
'        SpeKeycodename = "\��"
'
'    Case 188
'        SpeKeycodename = ",��"
'
'    Case 190
'        SpeKeycodename = ".��"
'
'    Case 191
'        SpeKeycodename = "/��"
'
'    Case 111
'        SpeKeycodename = "/��"
'
'    Case 193
'        SpeKeycodename = "\��"
'
'    Case 112
'        SpeKeycodename = "F1"
'
'    Case 113
'        SpeKeycodename = "F2"
'
'    Case 114
'        SpeKeycodename = "F3"
'
'    Case 115
'        SpeKeycodename = "F4"
'
'    Case 116
'        SpeKeycodename = "F5"
'
'    Case 117
'        SpeKeycodename = "F6"
'
'    Case 118
'        SpeKeycodename = "F7"
'
'    Case 119
'        SpeKeycodename = "F8"
'
'    Case 120
'        SpeKeycodename = "F9"
'
'    Case 121
'        SpeKeycodename = "F10"
'
'    Case 122
'        SpeKeycodename = "F11"
'
'    Case 123
'        SpeKeycodename = "F12"
'
'    Case 97
'        SpeKeycodename = "С����1"
'
'    Case 98
'        SpeKeycodename = "С����2"
'
'    Case 99
'        SpeKeycodename = "С����3"
'
'    Case 100
'        SpeKeycodename = "С����4"
'
'    Case 101
'        SpeKeycodename = "С����5"
'
'    Case 102
'        SpeKeycodename = "С����6"
'
'    Case 103
'        SpeKeycodename = "С����7"
'
'    Case 104
'        SpeKeycodename = "С����8"
'
'    Case 105
'        SpeKeycodename = "С����9"
'
'    Case Else
'        SpeKeycodename = ""
'    End Select
'End Sub

