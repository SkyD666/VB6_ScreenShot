VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   12225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12225
   ScaleWidth      =   19980
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5920
      Index           =   1
      Left            =   7920
      ScaleHeight     =   5895
      ScaleWidth      =   5895
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   5920
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   5415
         TabIndex        =   26
         Top             =   1440
         Width           =   5415
         Begin VB.OptionButton optDeclareHotKeyWay2 
            BackColor       =   &H80000005&
            Caption         =   "ϵͳ���̹���"
            Height          =   255
            Left            =   2520
            TabIndex        =   28
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton optDeclareHotKeyWay1 
            BackColor       =   &H80000005&
            Caption         =   "ϵͳ�ȼ�"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.TextBox txtHotKeyScreenShot 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'OFF
         Left            =   240
         TabIndex        =   15
         Text            =   "(��Ҫ����,�������,���°���)"
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label LabHotKeyAscii 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HotKeyAscii"
         Height          =   180
         Left            =   840
         TabIndex        =   34
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ȼ�ע�᷽ʽ��"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label HotKeyNow_lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ȼ���"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5920
      Index           =   0
      Left            =   1920
      ScaleHeight     =   5895
      ScaleWidth      =   5895
      TabIndex        =   7
      Top             =   120
      Width           =   5920
      Begin VB.CheckBox chkIncludeCursor 
         BackColor       =   &H80000005&
         Caption         =   "��ȡ���"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4080
         Width           =   5415
      End
      Begin VB.CheckBox chkAutoSendToClipBoard 
         BackColor       =   &H80000005&
         Caption         =   "�ȼ���ͼ��ֱ�ӽ�ͼƬ���Ƶ�������"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   5415
      End
      Begin VB.CheckBox chkAutoRun 
         BackColor       =   &H80000005&
         Caption         =   "��������"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   5415
      End
      Begin VB.CheckBox chkEndOrMin 
         BackColor       =   &H80000005&
         Caption         =   "�ر�������ʱֱ���˳������������С��������"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   5415
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   5415
         TabIndex        =   9
         Top             =   4860
         Width           =   5415
         Begin VB.OptionButton optActiveWinSnapMode0 
            BackColor       =   &H80000005&
            Caption         =   "�ɷ�ʽ"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   0
            Width           =   2175
         End
         Begin VB.OptionButton optActiveWinSnapMode1 
            BackColor       =   &H80000005&
            Caption         =   "�·�ʽ"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkSoundPlaySetFrm 
         BackColor       =   &H80000005&
         Caption         =   "��ͼ�󲥷���ʾ��"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5415
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   1880
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   5415
         Begin VB.HScrollBar hsbarDelayTime 
            Height          =   255
            LargeChange     =   3
            Left            =   120
            Max             =   10
            TabIndex        =   35
            Top             =   1460
            Width           =   5175
         End
         Begin VB.CheckBox chkHideWinValue 
            BackColor       =   &H80000005&
            Caption         =   "�����������"
            Height          =   255
            Left            =   2760
            TabIndex        =   32
            Top             =   720
            Width           =   2535
         End
         Begin VB.ComboBox cmbSnapName 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox chkAutoSaveSnapValue 
            BackColor       =   &H80000005&
            Caption         =   "�Զ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label labDelayTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ȡ����ڵȴ�ʱ�䣺"
            Height          =   180
            Left            =   120
            TabIndex        =   36
            Top             =   1160
            Width           =   1980
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ڽ�ͼ��ʽ��"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1620
      End
   End
   Begin VB.ListBox listSettings 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5920
      Index           =   2
      Left            =   13920
      ScaleHeight     =   5895
      ScaleWidth      =   5895
      TabIndex        =   3
      Top             =   120
      Width           =   5920
      Begin VB.CommandButton cmdOpenTheFolder 
         Caption         =   "�򿪴��ļ���..."
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox cmbAutoSaveSnapFormat 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtAutoSaveSnapPath 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   4935
      End
      Begin VB.CommandButton cmdAutoSaveSnapPath 
         Caption         =   "..."
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtSetJpgQuality 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Text            =   "80"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ������ͼ�ĸ�ʽ��"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ�����ͼƬ���ļ���(��Ŀ¼���������Զ�����)��"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   4320
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "��ͼ����ΪJPG��ʽƷ��(1-100),����Խ��,Ʒ��Խ��(һ��Ϊ80����)��"
         Height          =   420
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5920
      Index           =   3
      Left            =   1920
      ScaleHeight     =   5895
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   6120
      Width           =   5920
      Begin VB.ComboBox cmbChooseSoundPlay 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ���ͼ����ʾ����"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmLoadSoundPlayCmbBoo As Boolean                                           '�����ʱ��������ʾ��
Dim LoadfrmSettingsoptBoo As Boolean                                            '��־���봰��

Private Sub chkAutoRun_Click()
    Dim w As Object
    On Error GoTo Err:
    If chkAutoRun.Value = 1 Then
        Set w = CreateObject("wscript.shell")
        w.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\ScreenSnap", Chr(34) & App.path & "\" & "ScreenSnap" & ".exe" & Chr(34) & " AUTORUN", "REG_SZ"
        Set w = Nothing
    Else
        Set w = CreateObject("wscript.shell")
        w.Regdelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\ScreenSnap"
        Set w = Nothing
    End If
    Exit Sub
Err:
    '����ʧ��
    MsgBox LoadResString(11703) & vbCrLf & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbExclamation
    If chkAutoRun.Value = 1 Then
        chkAutoRun.Value = 0
    Else
        chkAutoRun.Value = 1
    End If
End Sub

Private Sub chkAutoSaveSnapValue_Click()
    If chkAutoSaveSnapValue.Value = 1 Then
        AutoSaveSnapInt(cmbSnapName.ListIndex) = 1
    Else
        AutoSaveSnapInt(cmbSnapName.ListIndex) = 0
    End If
End Sub

Private Sub chkAutoSendToClipBoard_Click()
    If chkAutoSendToClipBoard.Value = 1 Then
        AutoSendToClipBoardBoo = True
    Else
        AutoSendToClipBoardBoo = False
    End If
End Sub

Private Sub chkEndOrMin_Click()
    If chkEndOrMin.Value = 1 Then
        EndOrMinBoo = 1
    Else
        EndOrMinBoo = 0
    End If
End Sub

Private Sub chkHideWinValue_Click()
    If chkHideWinValue.Value = 1 Then
        HideWinCaptureInt(cmbSnapName.ListIndex) = 1
    Else
        HideWinCaptureInt(cmbSnapName.ListIndex) = 0
    End If
End Sub

Private Sub chkIncludeCursor_Click()
    If chkIncludeCursor.Value = 1 Then
        IncludeCursorBoo = True
    Else
        IncludeCursorBoo = False
    End If
End Sub

Private Sub chkSoundPlaySetFrm_Click()
    If chkSoundPlaySetFrm.Value = 1 Then
        frmMain.chkSoundPlay.Value = 1
    Else
        frmMain.chkSoundPlay.Value = 0
    End If
    SoundPlayInt = chkSoundPlaySetFrm.Value
End Sub

Private Sub cmbSnapName_Click()
    If cmbSnapName.Text = LoadResString(11311) Then
        chkHideWinValue.Visible = False
    Else
        chkHideWinValue.Visible = True
    End If
    chkAutoSaveSnapValue.Value = AutoSaveSnapInt(cmbSnapName.ListIndex)
    chkHideWinValue.Value = HideWinCaptureInt(cmbSnapName.ListIndex)
    labDelayTime.Caption = LoadResString(11307) & DelayTimeInt(cmbSnapName.ListIndex)
    hsbarDelayTime.Value = DelayTimeInt(cmbSnapName.ListIndex)
End Sub

Private Sub cmbChooseSoundPlay_Click()
    If frmLoadSoundPlayCmbBoo = False Then
        Select Case cmbChooseSoundPlay.List(cmbChooseSoundPlay.ListIndex)
        Case LoadResString(11602)
            ChooseSoundPlayStr = "XIANGJI"
        Case LoadResString(11603)
            ChooseSoundPlayStr = "FENGLING"
        Case LoadResString(11604)
            ChooseSoundPlayStr = "DIANJI"
        Case LoadResString(11605)
            ChooseSoundPlayStr = "JICHI"
        Case LoadResString(11606)
            ChooseSoundPlayStr = "QIAOJI"
        Case LoadResString(11607)
            ChooseSoundPlayStr = "BAOZHA"
        Case LoadResString(11608)
            ChooseSoundPlayStr = "JIGUANG"
        Case LoadResString(11609)
            ChooseSoundPlayStr = "DAZIJI"
        Case Else
            Exit Sub
        End Select
        Call SoundPlay
    End If
End Sub

Private Sub cmdAutoSaveSnapPath_Click()
    Dim FPath As String
    FPath = GetFolderName(Me.hwnd, LoadResString(11702))                        'ѡ���Զ�����ͼƬ���ļ���
    If FPath <> "" Then
        AutoSaveSnapPathStr = FPath
        txtAutoSaveSnapPath.Text = AutoSaveSnapPathStr
    End If
End Sub

Private Sub cmdOpenTheFolder_Click()
    Shell "explorer.exe " & txtAutoSaveSnapPath.Text, vbNormalFocus
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    
    Dim i As Integer                                                            '����
    
    Me.Width = 8040
    Me.Height = 6600
    
    LoadLanguages "frmSettings"
    
    listSettings.AddItem LoadResString(11201)                                   '��������
    listSettings.AddItem LoadResString(11202)                                   '�ȼ�����
    listSettings.AddItem LoadResString(11203)                                   'ͼƬ����
    listSettings.AddItem LoadResString(11204)                                   '��������
    
    cmbSnapName.AddItem LoadResString(10601)                                    'ȫ����ͼ
    cmbSnapName.AddItem LoadResString(10602)                                    '����ڽ�ͼ
    cmbSnapName.AddItem LoadResString(11311)                                    '�ȼ���ͼ
    cmbSnapName.AddItem LoadResString(11312)                                    '��ȡ���
    cmbSnapName.AddItem LoadResString(10813)                                    '��ȡ�κδ���/�ؼ�
    
    cmbSnapName.Text = cmbSnapName.List(0)
    chkAutoSaveSnapValue.Value = AutoSaveSnapInt(cmbSnapName.ListIndex)
    chkHideWinValue.Value = HideWinCaptureInt(cmbSnapName.ListIndex)
    labDelayTime.Caption = LoadResString(11307) & DelayTimeInt(cmbSnapName.ListIndex)
    hsbarDelayTime.Value = DelayTimeInt(cmbSnapName.ListIndex)
    
    cmbAutoSaveSnapFormat.AddItem "*.bmp"
    cmbAutoSaveSnapFormat.AddItem "*.jpg"
    cmbAutoSaveSnapFormat.AddItem "*.png"
    cmbAutoSaveSnapFormat.AddItem "*.gif"
    
    cmbChooseSoundPlay.AddItem LoadResString(11602)                             '���
    cmbChooseSoundPlay.AddItem LoadResString(11603)                             '����
    cmbChooseSoundPlay.AddItem LoadResString(11604)                             '���
    cmbChooseSoundPlay.AddItem LoadResString(11605)                             '����
    cmbChooseSoundPlay.AddItem LoadResString(11606)                             '�û�
    cmbChooseSoundPlay.AddItem LoadResString(11607)                             '��ը
    cmbChooseSoundPlay.AddItem LoadResString(11608)                             '����
    cmbChooseSoundPlay.AddItem LoadResString(11609)                             '���ֻ�
    
    For i = 0 To listSettings.ListCount - 1
        picSettings(i).Visible = False
    Next
    
    picSettings(0).Visible = True
    listSettings.Selected(0) = True
    
    '��������������ȡ����
    chkAutoRun.Value = AppAutoRun
    
    If ActiveWindowSnapMode = 0 Then
        optActiveWinSnapMode0.Value = True
    ElseIf ActiveWindowSnapMode = 1 Then
        optActiveWinSnapMode1.Value = True
    End If
    If EndOrMinBoo = 0 Then
        chkEndOrMin.Value = 0
    Else
        chkEndOrMin.Value = 1
    End If
    chkSoundPlaySetFrm.Value = SoundPlayInt
    txtSetJpgQuality.Text = SetJpgQuality
    txtAutoSaveSnapPath.Text = AutoSaveSnapPathStr
    chkAutoSendToClipBoard.Value = Abs(AutoSendToClipBoardBoo)                  'Trueת��ΪAbs(-1)��Falseת��Ϊ0
    chkIncludeCursor.Value = Abs(IncludeCursorBoo)                              'Trueת��ΪAbs(-1)��Falseת��Ϊ0
    
    cmbAutoSaveSnapFormat.Text = AutoSaveSnapFormatStr
    
    '��ʾ�ȼ�
    LabHotKeyAscii.Caption = AsciiToName(HotKeyCodeInt) & "  (Ascii:" & HotKeyCodeInt & ")"
    
    Select Case DeclareHotKeyWayInt                                             '1�ȼ�   2����
    Case 1
        LoadfrmSettingsoptBoo = True
        optDeclareHotKeyWay1.Value = True
        optDeclareHotKeyWay2.Value = False
    Case 2
        optDeclareHotKeyWay2.Value = True
        optDeclareHotKeyWay1.Value = False
    End Select
    
    
    Select Case ChooseSoundPlayStr
    Case "XIANGJI"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11602)
    Case "FENGLING"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11603)
    Case "DIANJI"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11604)
    Case "JICHI"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11605)
    Case "QIAOJI"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11606)
    Case "BAOZHA"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11607)
    Case "JIGUANG"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11608)
    Case "DAZIJI"
        frmLoadSoundPlayCmbBoo = True
        cmbChooseSoundPlay.Text = LoadResString(11609)
    End Select
    frmLoadSoundPlayCmbBoo = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err
    
    If txtSetJpgQuality.Text = "" Then
        MsgBox LoadResString(11700), vbExclamation + vbOKOnly                   'JPG��ʽƷ��ֵ��Ч��
        Cancel = 1
        Exit Sub
    End If
    
    AutoSaveSnapFormatStr = cmbAutoSaveSnapFormat.List(cmbAutoSaveSnapFormat.ListIndex)
    
    WritePrivateProfileString "Sound", "SoundPlay", CStr(chkSoundPlaySetFrm.Value), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "DelayTimeFullScreen", CStr(DelayTimeInt(0)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "DelayTimeActiveWindow", CStr(DelayTimeInt(1)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "DelayTimeHotKey", CStr(DelayTimeInt(2)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "DelayTimeCursor", CStr(DelayTimeInt(3)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "DelayTimeWindowCtrl", CStr(DelayTimeInt(4)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Picture", "SaveJpgQuality", CStr(CInt(txtSetJpgQuality.Text)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Sound", "ChooseSoundPlay", ChooseSoundPlayStr, App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "ActiveWindowSnapMode", CStr(ActiveWindowSnapMode), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "EndOrMin", CStr(CInt(EndOrMinBoo)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "HotKey", "HotKeyCode", CStr(HotKeyCodeInt), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Save", "AutoSaveSnapFullScreen", CStr(AutoSaveSnapInt(0)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Save", "AutoSaveSnapActiveWindow", CStr(AutoSaveSnapInt(1)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Save", "AutoSaveSnapHotKey", CStr(AutoSaveSnapInt(2)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Save", "AutoSaveSnapCursor", CStr(AutoSaveSnapInt(3)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Save", "AutoSaveSnapWindowCtrl", CStr(AutoSaveSnapInt(4)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "HideWinCaptureFullScreen", CStr(HideWinCaptureInt(0)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "HideWinCaptureActiveWindow", CStr(HideWinCaptureInt(1)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "HideWinCaptureCursor", CStr(HideWinCaptureInt(3)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "HideWinCaptureWindowCtrl", CStr(HideWinCaptureInt(4)), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Picture", "AutoSaveSnapPath", AutoSaveSnapPathStr, App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Picture", "AutoSaveSnapFormat", AutoSaveSnapFormatStr, App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "AutoSendToClipBoard", CStr(AutoSendToClipBoardBoo), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "Config", "IncludeCursor", CStr(IncludeCursorBoo), App.path & "\ScreenSnapConfig.ini"
    WritePrivateProfileString "HotKey", "DeclareHotKeyWay", CStr(DeclareHotKeyWayInt), App.path & "\ScreenSnapConfig.ini"
    Exit Sub
Err:
    MsgBox "����frmSetting.Form_Unload" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
    Cancel = 1
End Sub

Private Sub hsbarDelayTime_Change()
    labDelayTime.Caption = LoadResString(11307) & hsbarDelayTime.Value & "s"    '��ȡ����ڵȴ�ʱ�䣺
    DelayTimeInt(cmbSnapName.ListIndex) = hsbarDelayTime.Value
End Sub

Private Sub hsbarDelayTime_Scroll()
    labDelayTime.Caption = LoadResString(11307) & hsbarDelayTime.Value & "s"    '��ȡ����ڵȴ�ʱ�䣺
    DelayTimeInt(cmbSnapName.ListIndex) = hsbarDelayTime.Value
End Sub

Private Sub listSettings_Click()
    Dim i As Integer                                                            '����
    
    For i = 0 To listSettings.ListCount - 1
        If i = listSettings.ListIndex Then
            picSettings(listSettings.ListIndex).Top = 120
            picSettings(listSettings.ListIndex).Left = 1920
            picSettings(listSettings.ListIndex).Visible = True
        Else
            picSettings(i).Visible = False
        End If
    Next
End Sub

Private Sub optActiveWinSnapMode0_Click()
    If optActiveWinSnapMode0.Value = True Then ActiveWindowSnapMode = 0
End Sub

Private Sub optActiveWinSnapMode1_Click()
    If optActiveWinSnapMode1.Value = True Then ActiveWindowSnapMode = 1
End Sub

Private Sub optDeclareHotKeyWay1_Click()
    If optDeclareHotKeyWay1.Value And LoadfrmSettingsoptBoo = False Then
        DeclareHotKeyWayInt = 1
        
        UnKeyHook                                                               'ж�ع���
        
        RegHotkeySub                                                            'װ���ȼ�
    ElseIf optDeclareHotKeyWay1.Value And LoadfrmSettingsoptBoo Then
        LoadfrmSettingsoptBoo = False
    ElseIf optDeclareHotKeyWay1.Value = False Then
        UnRegHotkeySub                                                          'ж���ȼ�
    End If
End Sub

Private Sub optDeclareHotKeyWay2_Click()
    If optDeclareHotKeyWay2.Value = True Then
        DeclareHotKeyWayInt = 2
        
        UnRegHotkeySub                                                          'ж���ȼ�
        
        RegKeyHook                                                              'װ�ع���
    Else
        UnKeyHook                                                               'ж�ع���
    End If
End Sub

Private Sub txtAutoSaveSnapPath_Change()
    AutoSaveSnapPathStr = txtAutoSaveSnapPath.Text
End Sub

Private Sub txtHotKeyScreenShot_KeyUp(KeyCode As Integer, Shift As Integer)
    HotKeyCodeInt = KeyCode
    
    'ˢ���ȼ�
    If DeclareHotKeyWayInt = 1 Then
        SetWindowLong frmTray.hwnd, GWL_WNDPROC, preWinProcHotKey
        UnregisterHotKey frmTray.hwnd, 1                                        'ȡ��ϵͳ���ȼ�,�ͷ���Դ
        
        preWinProcHotKey = GetWindowLong(frmTray.hwnd, GWL_WNDPROC)
        SetWindowLong frmTray.hwnd, GWL_WNDPROC, AddressOf WndProcHotKey
        RegisterHotKey frmTray.hwnd, 1, 0, HotKeyCodeInt                        '�����ȼ�
    ElseIf DeclareHotKeyWayInt = 2 Then
        SetWindowLong frmTray.hwnd, GWL_WNDPROC, preWinProcHotKey
        UnregisterHotKey frmTray.hwnd, 1                                        'ȡ��ϵͳ���ȼ�,�ͷ���Դ
        
        RegKeyHook
    End If
    
    '��ʾ�ȼ�
    LabHotKeyAscii.Caption = AsciiToName(HotKeyCodeInt) & "  (Ascii:" & HotKeyCodeInt & ")"
End Sub

Private Sub txtSetJpgQuality_Change()
    If txtSetJpgQuality.Text = "" Then
        Exit Sub
    ElseIf CInt(txtSetJpgQuality.Text) > 100 Then
        MsgBox LoadResString(11701), vbExclamation + vbOKOnly                   '��Ч���֣�
        txtSetJpgQuality.Text = 100
    ElseIf CInt(txtSetJpgQuality.Text) < 1 Then
        MsgBox LoadResString(11701), vbExclamation + vbOKOnly                   '��Ч���֣�
        txtSetJpgQuality.Text = 1
    End If
    SetJpgQuality = CInt(txtSetJpgQuality.Text)
End Sub

Private Sub txtSetJpgQuality_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 And KeyAscii <> 8) Or (KeyAscii > 57 And KeyAscii <> 8) Then KeyAscii = 0
End Sub
