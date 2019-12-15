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
            ' 这里处理注册错误
            MsgBox "注册系统键盘钩子失败！"
        Else
            'MsgBox "注册键盘钩子完成！"
        End If
    End If
End Sub

Sub UnKeyHook()
    If hKeyboardHook <> 0 Then
        Dim retKeyboard As Long
        retKeyboard = UnhookWindowsHookEx(hKeyboardHook)
        hKeyboardHook = 0
        If retKeyboard = 0 Then
            ' 这里处理卸载错误
            MsgBox "卸载系统键盘钩子失败！"
        End If
    End If
End Sub

' 回调函数
Private Function KeyboardHookProcHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode >= 0 Then
        Dim ks As KeyboardHookStruct
        CopyMemory ks, ByVal lParam, 20
        
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
            ' 这里处理键盘按下的事件
            
            If ks.vkCode = HotKeyCodeInt And frmMain.Enabled = True Then HotKeyPressedBoo = True: frmMain.timerHotKey.Enabled = True
        End If
        
        If wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            ' 这里处理键盘弹起事件
            'frmMain.key_up ks.vkCode
        End If
        
        ' 想要截获按键事件，可以直接设置 KeyboardHookProcHookProc = 1
        ' 否则呼叫下一个钩子
        CallNextHookEx hKeyboardHook, nCode, wParam, lParam
    End If
End Function

'Sub key_down(ByVal code As Long)
'    Select Case code
'    Case 8
'        SpeKeycodename = "Backspace键"
'
'    Case 9
'        SpeKeycodename = "Tab键"
'
'    Case 13
'        SpeKeycodename = "回车键"
'
'    Case 19
'        SpeKeycodename = "Pause键"
'
'    Case 20
'        SpeKeycodename = "CapsLock键"
'
'    Case 27
'        SpeKeycodename = "ESC键"
'
'    Case 32
'        SpeKeycodename = "空格键"
'
'    Case 33
'        SpeKeycodename = "PageUp键"
'
'    Case 34
'        SpeKeycodename = "PageDown键"
'
'    Case 35
'        SpeKeycodename = "End键"
'
'    Case 36
'        SpeKeycodename = "Home键"
'
'    Case 37
'        SpeKeycodename = "方向键(←)"
'
'    Case 38
'        SpeKeycodename = "方向键(↑)"
'
'    Case 39
'        SpeKeycodename = "方向键(→)"
'
'    Case 40
'        SpeKeycodename = "方向键(↓)"
'
'    Case 44
'        SpeKeycodename = "PrtSc键"
'
'    Case 45
'        SpeKeycodename = "Insert键"
'
'    Case 46
'        SpeKeycodename = "Delete键"
'
'    Case 93
'        SpeKeycodename = "菜单键(appskey)"
'
'    Case 106
'        SpeKeycodename = "*键"
'
'    Case 107
'        SpeKeycodename = "+键"
'
'    Case 110
'        SpeKeycodename = ".键"
'
'    Case 144
'        SpeKeycodename = "NumLK键"
'
'    Case 145
'        SpeKeycodename = "ScrLK键"
'
'    Case 160
'        SpeKeycodename = "左Shift键"
'
'    Case 161
'        SpeKeycodename = "右Shift键"
'
'    Case 162
'        SpeKeycodename = "左Ctrl键"
'
'    Case 163
'        SpeKeycodename = "右Ctrl键"
'
'    Case 164
'        SpeKeycodename = "左Alt键"
'
'    Case 165
'        SpeKeycodename = "右Alt键"
'
'    Case 189
'        SpeKeycodename = "-键"
'
'    Case 109
'        SpeKeycodename = "-键"
'
'    Case 187
'        SpeKeycodename = "=键"
'
'    Case 192
'        SpeKeycodename = "`键"
'
'    Case 219
'        SpeKeycodename = "[键"
'
'    Case 221
'        SpeKeycodename = "]键"
'
'    Case 186
'        SpeKeycodename = ";键"
'
'    Case 222
'        SpeKeycodename = "'键"
'
'    Case 220
'        SpeKeycodename = "\键"
'
'    Case 188
'        SpeKeycodename = ",键"
'
'    Case 190
'        SpeKeycodename = ".键"
'
'    Case 191
'        SpeKeycodename = "/键"
'
'    Case 111
'        SpeKeycodename = "/键"
'
'    Case 193
'        SpeKeycodename = "\键"
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
'        SpeKeycodename = "小键盘1"
'
'    Case 98
'        SpeKeycodename = "小键盘2"
'
'    Case 99
'        SpeKeycodename = "小键盘3"
'
'    Case 100
'        SpeKeycodename = "小键盘4"
'
'    Case 101
'        SpeKeycodename = "小键盘5"
'
'    Case 102
'        SpeKeycodename = "小键盘6"
'
'    Case 103
'        SpeKeycodename = "小键盘7"
'
'    Case 104
'        SpeKeycodename = "小键盘8"
'
'    Case 105
'        SpeKeycodename = "小键盘9"
'
'    Case Else
'        SpeKeycodename = ""
'    End Select
'End Sub

