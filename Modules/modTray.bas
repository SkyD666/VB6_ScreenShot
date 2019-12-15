Attribute VB_Name = "modTray"
Option Explicit

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Public Const SW_RESTORE = 9
Public Const SW_SHOWNOACTIVATE = 4

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public pWndProc As Long
Public Type NOTIFYICONDATA
    cbSize As Long                                                              'NOTIFYICONDATA���͵Ĵ�С
    hwnd As Long                                                                '���Ӧ�ó����������
    uID As Long                                                                 'Ӧ�ó���ͼ����Դ��ID��
    uFlags As Long                                                              'ʹ��Щ������Ч��������ö�������е�NIF_MESSAGE��NIF_ICON��NIF_TIP��������
    uCallbackMessage As Long                                                    '����ƶ�ʱ�Ѵ���Ϣ������ͼ��Ĵ���
    hIcon As Long                                                               'ͼ������
    szTip As String * 128                                                       '�������ͼ����ʱ��ʾ��Tip�ı�
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutAndVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
    
Public Enum enm_NIM_Shell
    NIM_ADD = &H0                                                               '����ͼ��
    NIM_MODIFY = &H1                                                            '�޸�ͼ��
    NIM_DELETE = &H2                                                            'ɾ��ͼ��
    NIF_MESSAGE = &H1                                                           'ʹ����"NOTIFYICONDATA"�е�uCallbackMessage��Ч
    NIF_ICON = &H2                                                              'ʹ����"NOTIFYICONDATA"�е�hIcon��Ч
    NIF_TIP = &H4                                                               'ʹ����"NOTIFYICONDATA"�е�szTip��Ч
End Enum

Public Const WM_MOUSEMOVE = &H200                                               '��ͼ�����ƶ����
Public Const WM_LBUTTONDOWN = &H201                                             '����������
Public Const WM_LBUTTONUP = &H202                                               '�������ͷ�
Public Const WM_LBUTTONDBLCLK = &H203                                           '˫��������
Public Const WM_RBUTTONDOWN = &H204                                             '����Ҽ�����
Public Const WM_RBUTTONUP = &H205                                               '����Ҽ��ͷ�
Public Const WM_RBUTTONDBLCLK = &H206                                           '˫������Ҽ�
Public Const WM_SETHOTKEY = &H32                                                '��Ӧ��������ȼ�
Public nidProgramData As NOTIFYICONDATA
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
Private Const WM_USER As Long = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
Private Const NOTIFYICON_VERSION = 3
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Private Const WM_CLOSE As Long = &H10
    
Public Enum bFlag
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_GUID = &H5
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10                                                         '�ر���ʾ����־
End Enum

'����¼�
Public Enum TrayRetunEventEnum
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
    BalloonClick = (WM_USER + 5)
End Enum

Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" _
    (ByVal lpString As String) As Long
Public MsgTaskbarRestart As Long

'��������
Public Sub TrayAddIcon(ByVal MyForm As Form, ByVal ToolTip As String, Optional ByVal bFlag As bFlag, Optional ByVal Boo As Boolean)
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmMain.Icon
        .szTip = ToolTip & vbNullChar
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, nidProgramData)
    
    pWndProc = SetWindowLong(frmTray.hwnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub TrayRemoveIcon()
    
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    
    'pWndProc = SetWindowLong(frmTray.hWnd, GWL_WNDPROC, AddressOf WndProc)      '�����˳�ʱ����ֹͣ���С����������
End Sub

Public Sub TrayBalloon(ByVal MyForm As Form, ByVal sBaloonText As String, sBallonTitle As String, Optional ByVal bFlag As bFlag)
    
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_INFO
        .dwInfoFlags = bFlag
        .szInfoTitle = sBallonTitle & vbNullChar
        .szInfo = sBaloonText & vbNullChar
    End With
    
    Shell_NotifyIcon NIM_MODIFY, nidProgramData
    
End Sub

Public Sub TrayTip(ByVal MyForm As Form, ByVal sTipText As String)
    
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .szTip = sTipText & vbNullChar
    End With
    
    Shell_NotifyIcon NIM_MODIFY, nidProgramData
    
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Explorer.exe ����֮���ؽ�������ͼ��
    If Msg = MsgTaskbarRestart Then
        'ԭ��Explorer.exe �����������ؽ�ϵͳ����������ϵͳ������������ʱ�����ϵͳ������
        'ע�����TaskbarCreated ��Ϣ�Ķ������ڷ���һ����Ϣ������ֻ��Ҫ��׽�����Ϣ�����ؽ�ϵͳ���̵�ͼ�꼴�ɡ�
        With nidProgramData
            .cbSize = Len(nidProgramData)
            .hwnd = frmMain.hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = frmMain.Icon
            .szTip = App.Title & vbNullChar
        End With
        
        Call Shell_NotifyIcon(NIM_ADD, nidProgramData)                          '�ؼ���һ��,ʹͼ���ؽ�
        WndProc = True
    End If
    WndProc = CallWindowProc(pWndProc, hwnd, Msg, wParam, lParam)
End Function
