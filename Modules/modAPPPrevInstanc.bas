Attribute VB_Name = "modAPPPrevInstanc"
Option Explicit
Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim Ret As Long

'    hwnd Long�����ھ����Ҫ���������Ӧ����nCmdShowָ��������
'    nCmdShow Long��Ϊ����ָ�������Է����һ��������������κ�һ������
'    SW_HIDE ���ش��ڣ��״̬����һ������
'    SW_MINIMIZE ��С�����ڣ��״̬����һ������
'    SW_RESTORE ��ԭ���Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
'    SW_SHOW �õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
'    SW_SHOWMAXIMIZED ��󻯴��ڣ������伤��
'    SW_SHOWMINIMIZED ��С�����ڣ������伤��
'    SW_SHOWMINNOACTIVE ��С��һ�����ڣ�ͬʱ���ı�����
'    SW_SHOWNA �õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ����ı�����
'    SW_SHOWNOACTIVATE ������Ĵ�С��λ����ʾһ�����ڣ�ͬʱ���ı�����
'    SW_SHOWNORMAL ��SW_RESTORE��ͬ

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_FORCEMINIMIZE = 11
Private Const SW_MAX = 11

Public Sub APPPrevInstance()                                                    '��ֹ����������
    Dim WinHwnd As Long
    Ret = CreateMutex(ByVal 0, 1, App.Title)                                    '����ĳɳ���ı���
    If Err.LastDllError = 183 Then
        ReleaseMutex Ret
        CloseHandle Ret
        frmMain.Caption = ""
        frmMain.Hide
        MsgBox LoadResString(10811), vbExclamation + vbOKOnly                   '����������!
        'WinHwnd = FindWindow("ThunderRT6FormDC", "ScreenSnap")                  '�˴���Ĭ�ϵĴ��������ͱ��⣬���������Ҫ�޸�
        'ShowWindow WinHwnd, SW_SHOW                                             '�ò��ҵ��Ĵ���������ʾ�������Ƿ���С��
        'SetForegroundWindow WinHwnd                                             'ʹ�����ý���
        EndOrMinBoo = True
        Unload frmMain
        End
    End If
End Sub

