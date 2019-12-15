Attribute VB_Name = "modSystemHotKey"
Option Explicit

Public DeclareHotKeyWayInt As Integer                                           '1为热键，2为钩子
Public HotKeyCodeInt As Integer, HotKeyPressedBoo As Boolean
'Public DeclareHotKeyWayInt As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
'在窗口结构中为指定的窗口设置信息
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
'上述五个API函数是注册系统级热键所必需的，具体实现过程如后文所示  '常数
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8
Public Const WM_HOTKEY = &H312
'Public Const GWL_WNDPROC = (-4)                                                 '定义系统的热键,原中断标示,被隐藏的项目句柄
Public preWinProcHotKey As Long, MyhWnd As Long, uVirtKey As Long

Public Function WndProcHotKey(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY And frmMain.Enabled = True Then                          '如果拦截到热键标志常数
        HotkeyPressed wParam                                                    '执行隐藏鼠标所指项目
    End If                                                                      '如果不是热键,或者不是我们设置的热键,交还控制权给系统,继续监测热键
    WndProcHotKey = CallWindowProc(preWinProcHotKey, hwnd, Msg, wParam, lParam)
End Function

Public Sub HotkeyPressed(ByVal id As Long)
    HotKeyPressedBoo = True
    frmMain.timerHotKey.Enabled = True
End Sub

Public Sub RegHotkeySub()                                                       '注册热键
    preWinProcHotKey = GetWindowLong(frmTray.hwnd, GWL_WNDPROC)
    SetWindowLong frmTray.hwnd, GWL_WNDPROC, AddressOf WndProcHotKey
    RegisterHotKey frmTray.hwnd, 1, 0, HotKeyCodeInt                            '按下热键
    '        RegisterHotKey Me.hWnd, 1, MOD_ALT, vbKeyF12                            '按下ATL+F12
    '        RegisterHotKey Me.hWnd, 2, MOD_CONTROL, vbKeyF12                        '按下CTRL+F12
    '        RegisterHotKey Me.hWnd, 3, MOD_SHIFT, vbKeyF12                          '按下SHIFT+F12
    '        RegisterHotKey Me.hWnd, 4, MOD_WIN, vbKeyF12                            '按下WINDOWS键+F12
    '        RegisterHotKey Me.hWnd, 5, 0, vbKeyF12                                  '直接按F12
    '        RegisterHotKey Me.hWnd, 6, MOD_ALT Or MOD_CONTROL, vbKeyF12             '按下ALT+CTRL+F12键
End Sub

Public Sub UnRegHotkeySub()                                                     '释放热键
    SetWindowLong frmTray.hwnd, GWL_WNDPROC, preWinProcHotKey
    UnregisterHotKey frmTray.hwnd, 1                                            '取消系统级热键,释放资源End Sub
End Sub
