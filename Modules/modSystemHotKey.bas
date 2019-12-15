Attribute VB_Name = "modSystemHotKey"
Option Explicit

Public DeclareHotKeyWayInt As Integer                                           '1Ϊ�ȼ���2Ϊ����
Public HotKeyCodeInt As Integer, HotKeyPressedBoo As Boolean
'Public DeclareHotKeyWayInt As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
'�ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
'�������API������ע��ϵͳ���ȼ�������ģ�����ʵ�ֹ����������ʾ  '����
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8
Public Const WM_HOTKEY = &H312
'Public Const GWL_WNDPROC = (-4)                                                 '����ϵͳ���ȼ�,ԭ�жϱ�ʾ,�����ص���Ŀ���
Public preWinProcHotKey As Long, MyhWnd As Long, uVirtKey As Long

Public Function WndProcHotKey(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY And frmMain.Enabled = True Then                          '������ص��ȼ���־����
        HotkeyPressed wParam                                                    'ִ�����������ָ��Ŀ
    End If                                                                      '��������ȼ�,���߲����������õ��ȼ�,��������Ȩ��ϵͳ,��������ȼ�
    WndProcHotKey = CallWindowProc(preWinProcHotKey, hwnd, Msg, wParam, lParam)
End Function

Public Sub HotkeyPressed(ByVal id As Long)
    HotKeyPressedBoo = True
    frmMain.timerHotKey.Enabled = True
End Sub

Public Sub RegHotkeySub()                                                       'ע���ȼ�
    preWinProcHotKey = GetWindowLong(frmTray.hwnd, GWL_WNDPROC)
    SetWindowLong frmTray.hwnd, GWL_WNDPROC, AddressOf WndProcHotKey
    RegisterHotKey frmTray.hwnd, 1, 0, HotKeyCodeInt                            '�����ȼ�
    '        RegisterHotKey Me.hWnd, 1, MOD_ALT, vbKeyF12                            '����ATL+F12
    '        RegisterHotKey Me.hWnd, 2, MOD_CONTROL, vbKeyF12                        '����CTRL+F12
    '        RegisterHotKey Me.hWnd, 3, MOD_SHIFT, vbKeyF12                          '����SHIFT+F12
    '        RegisterHotKey Me.hWnd, 4, MOD_WIN, vbKeyF12                            '����WINDOWS��+F12
    '        RegisterHotKey Me.hWnd, 5, 0, vbKeyF12                                  'ֱ�Ӱ�F12
    '        RegisterHotKey Me.hWnd, 6, MOD_ALT Or MOD_CONTROL, vbKeyF12             '����ALT+CTRL+F12��
End Sub

Public Sub UnRegHotkeySub()                                                     '�ͷ��ȼ�
    SetWindowLong frmTray.hwnd, GWL_WNDPROC, preWinProcHotKey
    UnregisterHotKey frmTray.hwnd, 1                                            'ȡ��ϵͳ���ȼ�,�ͷ���ԴEnd Sub
End Sub
