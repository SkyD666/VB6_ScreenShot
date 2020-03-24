Attribute VB_Name = "modPublic"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RtlGetNtVersionNumbers& Lib "ntdll" (Major As Long, Minor As Long, Optional Build As Long) '获取系统版本
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const SWP_NOMOVE = &H2                                                   '不更动目前视窗位置
Public Const SWP_NOSIZE = &H1                                                   '不更动目前视窗大小
Public Const HWND_TOPMOST = -1                                                  '设定为最上层
Public Const HWND_NOTTOPMOST = -2                                               '取消最上层

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_OK = &H0&
'---------------------操作ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'---------------------
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_ITEMFROMPOINT = &H1A9

Public SysMajor As Long, SysMinor As Long, SysBuild As Long                     '保存系统版本信息
Public AutoSendToClipBoardBoo As Boolean                                        '热键截图后直接将图片复制到剪贴板
Public AutoSaveSnapFormatStr As String                                          '自动保存图片格式
Public AutoSaveSnapPathStr As String                                            '自动保存图片路径
Public AutoSaveSnapInt(4) As Integer                                            '截图是否保存，0为不保存，1为保存  (index 0,1,2,3,4 全屏截图，活动窗口，热键，鼠标，任何窗口)
Public HideWinCaptureInt(4) As Integer                                          '截图时是否隐藏窗口，0为不保存，1为保存  (index 0,1,2,3,4 全屏截图，活动窗口，热键，鼠标，任何窗口)
Public CloseFilesModeInt As Integer                                             '记录从哪里触发关闭子窗体事件，0为单独关闭，1为退出程序触发
Public NewMsgBoxInt As Integer                                                  '窗体消息框，0为取消，1为是，2为否，3为全部是，4为全部否
Public EndOrMinBoo As Boolean                                                   '关闭主窗口时直接退出程序而不是最小化到托盘，0是最小化，非零是退出
Public ActiveWindowSnapMode As Integer                                          '记录活动窗口截图方式，0为原始方法，1为新方法
Public CloseAllFilesUnsavedBoo As Boolean
Public ScreenShotHideFormInt As Integer                                         '是否截图时隐藏窗口
Public SoundPlayBoo As Boolean                                                  '是否播放提示音
Public DelayTimeInt(4) As Integer                                               '活动窗口截图延时，(index 0,1,2,3,4 全屏截图，活动窗口，热键，鼠标，任何窗口)
Public ChooseSoundPlayStr As String, SoundPlayInt As Integer                    '选择的提示音
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()                     'XP视觉样式
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public PicFilesCount As Long                                                    '文件计数，不减只加
Public SnapWhenTrayLng As Long, SnapWhenTrayBoo As Boolean                      '托盘时截图数
Public frmPicNum As Integer                                                     '窗体编号
Public PictureForms As New Collection

Public Sub UnloadfrmPic(ByVal FrmNum As Integer)
    Dim i As Integer, n As Integer                                              '计数
    frmPicNum = frmPicNum - 1
    'Unload PictureForms.Item(1 + FrmNum)
    PictureForms.Remove (1 + FrmNum)
    n = frmMain.listSnapPic.ListIndex
    frmMain.listSnapPic.RemoveItem (n)                                          '先移除此条
    If n > 0 Then frmMain.listSnapPic.Selected(n - 1) = True                    '再选中上一条，防止当在最后一条时无法自动选中下一条
    
    TrayTip frmMain, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '共   张截图        '刷新托盘提示文本
End Sub

Public Function GetPicturePath(ByVal frm As Form, ByVal FrmNum As Long) As String '选择保存的文件名(optval=1是，“全部是”保存)
    On Error GoTo Err
    Dim SaveToFileName As String, EXEFiles() As Byte
    SaveToFileName = GetDialog("save", "保存到文件", Format(Now, "yyyy-MM-dd_hh-mm-ss") & ".bmp", frm.hwnd)
    If SaveToFileName = "" Then
        GetPicturePath = ""
    Else
        GetPicturePath = Split(SaveToFileName, "\")(UBound(Split(SaveToFileName, "\")))
    End If
    Exit Function
Err:
    MsgBox "错误！GetPicturePath" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, vbCritical + vbOKOnly
End Function

Public Function SavePictures(ByVal Str As String, ByVal FrmNum As Long, Optional ByVal OptVal = 0) As Integer '保存文件(optval=1是，“全部是”保存)
    On Error GoTo Err
    SavePictures = 0
    Dim EXEFiles() As Byte
    If Str <> "" Then
        If OptVal = 0 Then
            If SnapWhenTrayBoo Then
                '托盘截图自动保存时不需要在此处添加条目，在frmmain显示时会添加
            Else
                PictureForms.Item(1 + FrmNum).PictureName = Str
                frmMain.listSnapPic.List(frmMain.listSnapPic.ListIndex) = Str
            End If
            PictureForms.Item(1 + FrmNum).PictureSaved = True
            
            Call SaveStdPicToFile(PictureForms.Item(1 + FrmNum).PictureData, Str, Split(Str, ".")(UBound(Split(Str, "."))))
            
            SavePictures = 1
        ElseIf OptVal = 1 Then
            Dim i As Long, NewName As String
            For i = 0 To frmPicNum
                Randomize                                                       '1000-99999随机数
                ShowProgressBar Format((i + 1) / (frmPicNum + 1), "0.000")      '进度条
                If PictureForms.Item(1 + i).PictureSaved = False Then
                    NewName = Mid(Str, 1, InStrRev(Str, ".") - 1) & (i + 1) & "_" & (1000 + Int(Rnd * 98999)) & "." & Split(Str, ".")(UBound(Split(Str, ".")))
                    Call SaveStdPicToFile(PictureForms.Item(1 + i).PictureData, NewName, Split(NewName, ".")(UBound(Split(NewName, "."))))
                    PictureForms.Item(1 + i).PictureSaved = True
                End If
            Next i
        End If
    End If
    Exit Function
Err:
    MsgBox "错误！SavePictures" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, vbCritical + vbOKOnly
    SavePictures = 0
End Function

Public Sub AutoSaveSnapSub(ByVal Value As Single, ByVal Num As Long)            '自动保存图片    0为全屏，1为活动，2为热键，3为光标，4为任何窗口
    Dim FilesName As String, IDStr As String
    Select Case Value
    Case 0
        IDStr = LoadResString(10601)
    Case 1
        IDStr = LoadResString(10602)
    Case 2
        IDStr = LoadResString(11311)
    Case 3
        IDStr = LoadResString(10812)
    Case 4
        IDStr = LoadResString(10813)
    End Select
    Randomize
    FilesName = IDStr & " - " & Format(Now, "yyyy-MM-dd-hh-mm-ss") & (frmPicNum + 1) & "_" & Int(Rnd * 98999) + 1000 'Int(Rnd * n) + m,生成m到n的随机数其中,n,m为integer类型
    '文件夹不存在  '在应用程序根目下，创建文件夹
    If Dir(AutoSaveSnapPathStr, vbDirectory) = "" Then MkDir AutoSaveSnapPathStr
    SavePictures AutoSaveSnapPathStr & "\" & FilesName & Replace(AutoSaveSnapFormatStr, "*", ""), frmPicNum, 0
    PictureForms.Item(1 + frmPicNum).PictureName = AutoSaveSnapPathStr & "\" & FilesName & Replace(AutoSaveSnapFormatStr, "*", "")
    PictureForms.Item(1 + frmPicNum).Caption = PictureForms.Item(1 + frmPicNum).PictureName
End Sub

Public Sub ShowProgressBar(ByVal Value As Single)                               '进度条        0-1的小数
    If frmProgressBar.Visible = False Then frmProgressBar.Show
    frmProgressBar.shpProgressBar.Width = Value * frmProgressBar.Shape1.Width
    frmProgressBar.Label1.Caption = Value * 100 & "%"
    If Value = 1 Then Unload frmProgressBar
End Sub

Public Sub DelayTimeSub(ByVal Index As Single)                                  '进度条        0,1,2,3,4全屏，活动窗口，热键，光标，任意窗口控件
    Dim EndTime As Date
    EndTime = DateAdd("s", DelayTimeInt(Index), Now)
    Do Until Now >= EndTime
        DoEvents
    Loop
End Sub
