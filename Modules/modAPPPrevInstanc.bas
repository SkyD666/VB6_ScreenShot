Attribute VB_Name = "modAPPPrevInstanc"
Option Explicit
Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim Ret As Long

'    hwnd Long，窗口句柄，要向这个窗口应用由nCmdShow指定的命令
'    nCmdShow Long，为窗口指定可视性方面的一个命令。请用下述任何一个常数
'    SW_HIDE 隐藏窗口，活动状态给令一个窗口
'    SW_MINIMIZE 最小化窗口，活动状态给令一个窗口
'    SW_RESTORE 用原来的大小和位置显示一个窗口，同时令其进入活动状态
'    SW_SHOW 用当前的大小和位置显示一个窗口，同时令其进入活动状态
'    SW_SHOWMAXIMIZED 最大化窗口，并将其激活
'    SW_SHOWMINIMIZED 最小化窗口，并将其激活
'    SW_SHOWMINNOACTIVE 最小化一个窗口，同时不改变活动窗口
'    SW_SHOWNA 用当前的大小和位置显示一个窗口，不改变活动窗口
'    SW_SHOWNOACTIVATE 用最近的大小和位置显示一个窗口，同时不改变活动窗口
'    SW_SHOWNORMAL 与SW_RESTORE相同

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

Public Sub APPPrevInstance()                                                    '阻止程序多次启动
    Dim WinHwnd As Long
    Ret = CreateMutex(ByVal 0, 1, App.Title)                                    '这里改成程序的标题
    If Err.LastDllError = 183 Then
        ReleaseMutex Ret
        CloseHandle Ret
        frmMain.Caption = ""
        frmMain.Hide
        MsgBox LoadResString(10811), vbExclamation + vbOKOnly                   '程序已运行!
        'WinHwnd = FindWindow("ThunderRT6FormDC", "ScreenSnap")                  '此处是默认的窗体类名和标题，根据你的需要修改
        'ShowWindow WinHwnd, SW_SHOW                                             '让查找到的窗体正常显示，无论是否最小化
        'SetForegroundWindow WinHwnd                                             '使程序获得焦点
        EndOrMinBoo = True
        Unload frmMain
        End
    End If
End Sub

