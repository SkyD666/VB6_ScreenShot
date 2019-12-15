Attribute VB_Name = "modPublic"
Option Explicit
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
Public PictureData() As Picture                                                 '�洢ͼƬ
Public frmPictureCopy() As New frmPicture, frmPicNum As Integer, frmPictureSaved() As Boolean, frmPictureName() As String, PicZoom() As Integer
'          �࿪����                                            ������                      ͼƬ�Ƿ񱣴�                           ÿ��������ͼƬ����
Public Sub UnloadfrmPic(ByVal FrmNum As Integer)
    Dim i As Integer, n As Integer                                              '����
    For i = FrmNum To frmPicNum - 1                                             '��ر���2����4�����ڣ���3�ű��2�ţ��ĺű������
        Set frmPictureCopy(i) = frmPictureCopy(i + 1)
        frmPictureCopy(i).labfrmi.Caption = i                                   '�����ĺ����Ŵ��ݸ�����
        frmPictureSaved(i) = frmPictureSaved(i + 1)                             '�Ƿ񱣴�ͼƬ ��Ϣ
        PicZoom(i) = PicZoom(i + 1)                                             '�Ƿ񱣴�ͼƬ ��Ϣ
        Set PictureData(i) = PictureData(i + 1)
        frmPictureName(i) = frmPictureName(i + 1)
    Next
    frmPicNum = frmPicNum - 1
    If frmPicNum > -1 Then
        ReDim Preserve frmPictureCopy(0 To frmPicNum) As New frmPicture
        ReDim Preserve frmPictureSaved(0 To frmPicNum) As Boolean
        ReDim Preserve PicZoom(0 To frmPicNum) As Integer
        ReDim Preserve PictureData(0 To frmPicNum) As Picture
        ReDim Preserve frmPictureName(0 To frmPicNum) As String
    End If
    n = frmMain.listSnapPic.ListIndex
    frmMain.listSnapPic.RemoveItem (n)                                          '���Ƴ�����
    If n > 0 Then frmMain.listSnapPic.Selected(n - 1) = True                    '��ѡ����һ������ֹ�������һ��ʱ�޷��Զ�ѡ����һ��
    
    TrayTip frmMain, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '��   �Ž�ͼ        'ˢ��������ʾ�ı�
End Sub

Public Function SaveFiles(ByVal frm As Form, ByVal FrmNum As Long) As String    'ѡ�񱣴���ļ���(optval=1�ǣ���ȫ���ǡ�����)
    On Error GoTo Err
    Dim SaveToFileName As String, EXEFiles() As Byte
    SaveToFileName = GetDialog("save", "���浽�ļ�", Format(Now, "yyyy-MM-dd_hh-mm-ss") & ".bmp", frm.hwnd)
    If SaveToFileName = "" Then
        SaveFiles = ""
    Else
        SaveFiles = Split(SaveToFileName, "\")(UBound(Split(SaveToFileName, "\")))
    End If
    Exit Function
Err:
    If (Err.Description = "�ļ�δ�ҵ��� gdiplus" And Err.Number = 53) Or (Err.Description = "File not found: gdiplus" And Err.Number = 53) _
        Or (Err.Description = "�ļ�δ�ҵ��� gdiplus" And Err.Number = 48) Or (Err.Description = "File not found: gdiplus" And Err.Number = 48) Then
        EXEFiles = LoadResData(101, "CUSTOM")
        '�Զ����Ʒ�ʽд�����ɣ�����ǰĿ¼
        Open App.path & "\GdiPlus.dll" For Binary As #1
        Put #1, , EXEFiles
        Close #1
        
        SaveFiles = Split(SaveToFileName, "\")(UBound(Split(SaveToFileName, "\")))
    Else
        MsgBox "����" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
    End If
End Function

Public Sub SaveFiles2(ByVal Str As String, ByVal FrmNum As Long, Optional ByVal OptVal = 0) '�����ļ�(optval=1�ǣ���ȫ���ǡ�����)
    On Error GoTo Err
nxt:
    Dim SelectedInt As Integer, EXEFiles() As Byte
    If Str = "" Then
        Exit Sub
    Else
        If OptVal = 0 Then
            frmPictureName(FrmNum) = Str
            frmPictureCopy(FrmNum).Caption = frmPictureName(CInt(frmPictureCopy(FrmNum).labfrmi.Caption))
            If SnapWhenTrayBoo Then
                '���̽�ͼ�Զ�����ʱ����Ҫ�ڴ˴�������Ŀ����frmmain��ʾʱ������
            Else
                frmMain.listSnapPic.AddItem frmPictureName(FrmNum), frmMain.listSnapPic.ListIndex
                SelectedInt = frmMain.listSnapPic.ListIndex - 1
                frmMain.listSnapPic.RemoveItem frmMain.listSnapPic.ListIndex
                frmMain.listSnapPic.Selected(SelectedInt) = True
            End If
            frmPictureSaved(FrmNum) = True
            
            Call SaveStdPicToFile(PictureData(FrmNum), Str, Split(Str, ".")(UBound(Split(Str, "."))))
        ElseIf OptVal = 1 Then
            Dim i As Long, NewName As String
            For i = 0 To frmPicNum
                Randomize                                                       '1000-99999�����
                ShowProgressBar Format((i + 1) / (frmPicNum + 1), "0.000")      '������
                If frmPictureSaved(i) = False Then
                    NewName = Mid(Str, 1, InStrRev(Str, ".") - 1) & (i + 1) & "_" & (1000 + Int(Rnd * 98999)) & "." & Split(Str, ".")(UBound(Split(Str, ".")))
                    Call SaveStdPicToFile(PictureData(i), NewName, Split(NewName, ".")(UBound(Split(NewName, "."))))
                    frmPictureSaved(i) = True
                End If
            Next i
        End If
    End If
    Exit Sub
Err:
    If (Err.Description = "�ļ�δ�ҵ��� gdiplus" And Err.Number = 53) Or (Err.Description = "File not found: gdiplus" And Err.Number = 53) _
        Or (Err.Description = "�ļ�δ�ҵ��� gdiplus" And Err.Number = 48) Or (Err.Description = "File not found: gdiplus" And Err.Number = 48) Then
        EXEFiles = LoadResData(101, "CUSTOM")
        '�Զ����Ʒ�ʽд�����ɣ�����ǰĿ¼
        Open App.path & "\GdiPlus.dll" For Binary As #1
        Put #1, , EXEFiles
        Close #1
        GoTo nxt
    Else
        MsgBox "����" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
    End If
End Sub

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
    
    If Dir(AutoSaveSnapPathStr, vbDirectory) = "" Then                          '�ļ��в�����
        MkDir AutoSaveSnapPathStr                                               '��Ӧ�ó����Ŀ�£������ļ���
    End If
    SaveFiles2 AutoSaveSnapPathStr & "\" & FilesName & Replace(AutoSaveSnapFormatStr, "*", ""), frmPicNum, 0
    
    'frmPictureSaved(Num) = True
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