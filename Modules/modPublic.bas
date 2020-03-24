Attribute VB_Name = "modPublic"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RtlGetNtVersionNumbers& Lib "ntdll" (Major As Long, Minor As Long, Optional Build As Long) '��ȡϵͳ�汾
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const SWP_NOMOVE = &H2                                                   '������Ŀǰ�Ӵ�λ��
Public Const SWP_NOSIZE = &H1                                                   '������Ŀǰ�Ӵ���С
Public Const HWND_TOPMOST = -1                                                  '�趨Ϊ���ϲ�
Public Const HWND_NOTTOPMOST = -2                                               'ȡ�����ϲ�

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_OK = &H0&
'---------------------����ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'---------------------
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_ITEMFROMPOINT = &H1A9

Public SysMajor As Long, SysMinor As Long, SysBuild As Long                     '����ϵͳ�汾��Ϣ
Public AutoSendToClipBoardBoo As Boolean                                        '�ȼ���ͼ��ֱ�ӽ�ͼƬ���Ƶ�������
Public AutoSaveSnapFormatStr As String                                          '�Զ�����ͼƬ��ʽ
Public AutoSaveSnapPathStr As String                                            '�Զ�����ͼƬ·��
Public AutoSaveSnapInt(4) As Integer                                            '��ͼ�Ƿ񱣴棬0Ϊ�����棬1Ϊ����  (index 0,1,2,3,4 ȫ����ͼ������ڣ��ȼ�����꣬�κδ���)
Public HideWinCaptureInt(4) As Integer                                          '��ͼʱ�Ƿ����ش��ڣ�0Ϊ�����棬1Ϊ����  (index 0,1,2,3,4 ȫ����ͼ������ڣ��ȼ�����꣬�κδ���)
Public CloseFilesModeInt As Integer                                             '��¼�����ﴥ���ر��Ӵ����¼���0Ϊ�����رգ�1Ϊ�˳����򴥷�
Public NewMsgBoxInt As Integer                                                  '������Ϣ��0Ϊȡ����1Ϊ�ǣ�2Ϊ��3Ϊȫ���ǣ�4Ϊȫ����
Public EndOrMinBoo As Boolean                                                   '�ر�������ʱֱ���˳������������С�������̣�0����С�����������˳�
Public ActiveWindowSnapMode As Integer                                          '��¼����ڽ�ͼ��ʽ��0Ϊԭʼ������1Ϊ�·���
Public CloseAllFilesUnsavedBoo As Boolean
Public ScreenShotHideFormInt As Integer                                         '�Ƿ��ͼʱ���ش���
Public SoundPlayBoo As Boolean                                                  '�Ƿ񲥷���ʾ��
Public DelayTimeInt(4) As Integer                                               '����ڽ�ͼ��ʱ��(index 0,1,2,3,4 ȫ����ͼ������ڣ��ȼ�����꣬�κδ���)
Public ChooseSoundPlayStr As String, SoundPlayInt As Integer                    'ѡ�����ʾ��
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()                     'XP�Ӿ���ʽ
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public PicFilesCount As Long                                                    '�ļ�����������ֻ��
Public SnapWhenTrayLng As Long, SnapWhenTrayBoo As Boolean                      '����ʱ��ͼ��
Public frmPicNum As Integer                                                     '������
Public PictureForms As New Collection

Public Sub UnloadfrmPic(ByVal FrmNum As Integer)
    Dim i As Integer, n As Integer                                              '����
    frmPicNum = frmPicNum - 1
    'Unload PictureForms.Item(1 + FrmNum)
    PictureForms.Remove (1 + FrmNum)
    n = frmMain.listSnapPic.ListIndex
    frmMain.listSnapPic.RemoveItem (n)                                          '���Ƴ�����
    If n > 0 Then frmMain.listSnapPic.Selected(n - 1) = True                    '��ѡ����һ������ֹ�������һ��ʱ�޷��Զ�ѡ����һ��
    
    TrayTip frmMain, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '��   �Ž�ͼ        'ˢ��������ʾ�ı�
End Sub

Public Function GetPicturePath(ByVal frm As Form, ByVal FrmNum As Long) As String 'ѡ�񱣴���ļ���(optval=1�ǣ���ȫ���ǡ�����)
    On Error GoTo Err
    Dim SaveToFileName As String, EXEFiles() As Byte
    SaveToFileName = GetDialog("save", "���浽�ļ�", Format(Now, "yyyy-MM-dd_hh-mm-ss") & ".bmp", frm.hwnd)
    If SaveToFileName = "" Then
        GetPicturePath = ""
    Else
        GetPicturePath = Split(SaveToFileName, "\")(UBound(Split(SaveToFileName, "\")))
    End If
    Exit Function
Err:
    MsgBox "����GetPicturePath" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
End Function

Public Function SavePictures(ByVal Str As String, ByVal FrmNum As Long, Optional ByVal OptVal = 0) As Integer '�����ļ�(optval=1�ǣ���ȫ���ǡ�����)
    On Error GoTo Err
    SavePictures = 0
    Dim EXEFiles() As Byte
    If Str <> "" Then
        If OptVal = 0 Then
            If SnapWhenTrayBoo Then
                '���̽�ͼ�Զ�����ʱ����Ҫ�ڴ˴������Ŀ����frmmain��ʾʱ�����
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
                Randomize                                                       '1000-99999�����
                ShowProgressBar Format((i + 1) / (frmPicNum + 1), "0.000")      '������
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
    MsgBox "����SavePictures" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
    SavePictures = 0
End Function

Public Sub AutoSaveSnapSub(ByVal Value As Single, ByVal Num As Long)            '�Զ�����ͼƬ    0Ϊȫ����1Ϊ���2Ϊ�ȼ���3Ϊ��꣬4Ϊ�κδ���
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
    FilesName = IDStr & " - " & Format(Now, "yyyy-MM-dd-hh-mm-ss") & (frmPicNum + 1) & "_" & Int(Rnd * 98999) + 1000 'Int(Rnd * n) + m,����m��n�����������,n,mΪinteger����
    '�ļ��в�����  '��Ӧ�ó����Ŀ�£������ļ���
    If Dir(AutoSaveSnapPathStr, vbDirectory) = "" Then MkDir AutoSaveSnapPathStr
    SavePictures AutoSaveSnapPathStr & "\" & FilesName & Replace(AutoSaveSnapFormatStr, "*", ""), frmPicNum, 0
    PictureForms.Item(1 + frmPicNum).PictureName = AutoSaveSnapPathStr & "\" & FilesName & Replace(AutoSaveSnapFormatStr, "*", "")
    PictureForms.Item(1 + frmPicNum).Caption = PictureForms.Item(1 + frmPicNum).PictureName
End Sub

Public Sub ShowProgressBar(ByVal Value As Single)                               '������        0-1��С��
    If frmProgressBar.Visible = False Then frmProgressBar.Show
    frmProgressBar.shpProgressBar.Width = Value * frmProgressBar.Shape1.Width
    frmProgressBar.Label1.Caption = Value * 100 & "%"
    If Value = 1 Then Unload frmProgressBar
End Sub

Public Sub DelayTimeSub(ByVal Index As Single)                                  '������        0,1,2,3,4ȫ��������ڣ��ȼ�����꣬���ⴰ�ڿؼ�
    Dim EndTime As Date
    EndTime = DateAdd("s", DelayTimeInt(Index), Now)
    Do Until Now >= EndTime
        DoEvents
    Loop
End Sub
