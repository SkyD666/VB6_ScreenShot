VERSION 5.00
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "ScreenSnap"
   ClientHeight    =   9360
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13860
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picSnapPic 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9060
      Left            =   10890
      ScaleHeight     =   9060
      ScaleWidth      =   2970
      TabIndex        =   6
      Top             =   0
      Width           =   2970
      Begin VB.ListBox listSnapPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8910
         ItemData        =   "frmMain.frx":4781A
         Left            =   0
         List            =   "frmMain.frx":4781C
         TabIndex        =   7
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H009B9B9B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   352.941
      ScaleMode       =   0  'User
      ScaleWidth      =   13860
      TabIndex        =   1
      Top             =   9060
      Width           =   13860
      Begin VB.ComboBox cmbZoom 
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
         Left            =   12240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.CheckBox chkSoundPlay 
         BackColor       =   &H009B9B9B&
         Caption         =   "��ͼ�󲥷���ʾ��"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4920
         TabIndex        =   2
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label labMousePos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���λ��: X:0  Y:0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9360
         TabIndex        =   5
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label labPicQuantity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��0��,ѡ�е�0��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6840
         TabIndex        =   4
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H009B9B9B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9060
      Left            =   0
      ScaleHeight     =   9060
      ScaleWidth      =   705
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   705
      Begin VB.Timer timerHotKey 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   3480
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
         Begin VB.Image imgSideBarPic 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   7
            Left            =   0
            Picture         =   "frmMain.frx":4781E
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Height          =   375
            Index           =   6
            Left            =   0
            Picture         =   "frmMain.frx":478CE
            Stretch         =   -1  'True
            Top             =   2160
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Height          =   375
            Index           =   5
            Left            =   0
            Picture         =   "frmMain.frx":4797E
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   4
            Left            =   0
            Picture         =   "frmMain.frx":47B12
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   0
            Picture         =   "frmMain.frx":47CA7
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   2
            Left            =   0
            Picture         =   "frmMain.frx":47E0F
            Stretch         =   -1  'True
            Top             =   720
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Height          =   375
            Index           =   3
            Left            =   0
            Picture         =   "frmMain.frx":47EB7
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image imgSideBarPic 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   0
            Picture         =   "frmMain.frx":47F5F
            Stretch         =   -1  'True
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Image imgAnyCtrlWindow 
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image imgCursor 
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image imgActiveWin 
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Width           =   495
      End
      Begin VB.Image imgScreenSnap 
         Height          =   495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpenTheFolder 
         Caption         =   "�򿪽�ͼ�ļ���..."
      End
      Begin VB.Menu mnuCloseAllFilesUnsaved 
         Caption         =   "�ر������ĵ��Ҳ�����"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "�رմ���(&C)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�����(&E)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuCut4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu numFliphorizontal 
         Caption         =   "ˮƽ��ת"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "����(&C)"
      Begin VB.Menu mnuScreenSnap 
         Caption         =   "ȫ����ͼ(&S)"
      End
      Begin VB.Menu mnuActiveWinSnap 
         Caption         =   "����ڽ�ͼ(&W)"
      End
      Begin VB.Menu mnuCursorSnap 
         Caption         =   "������(&C)"
      End
      Begin VB.Menu mnuAnyWindowCtrlSnap 
         Caption         =   "��ȡ����/�ؼ�(&A)..."
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "����(&T)"
      Begin VB.Menu mnuSetting 
         Caption         =   "����(&S)..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuZoom 
         Caption         =   "����(&Z)"
         Begin VB.Menu mnuZoomIn 
            Caption         =   "�Ŵ�(&I)"
         End
         Begin VB.Menu mnuZoomOut 
            Caption         =   "��С(&O)"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuSourceCode 
         Caption         =   "����Դ��(&S)..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnufrmPictureRight 
      Caption         =   "frmPictureRight"
      Visible         =   0   'False
      Begin VB.Menu mnufrmPicCopy 
         Caption         =   "����"
      End
      Begin VB.Menu mnufrmPicPaste 
         Caption         =   "ճ��"
      End
      Begin VB.Menu mnufrmPicClose 
         Caption         =   "�رմ��ĵ�"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "��ʾ����..."
      End
      Begin VB.Menu mnuCut1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayScreenSnap 
         Caption         =   "ȫ����ͼ"
      End
      Begin VB.Menu mnuTrayActiveWinSnap 
         Caption         =   "����ڽ�ͼ"
      End
      Begin VB.Menu mnuTrayCursorSnap 
         Caption         =   "������"
      End
      Begin VB.Menu mnuTrayWinCtrlSnap 
         Caption         =   "��ȡ����/�ؼ�"
      End
      Begin VB.Menu mnuCut2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTraySetting 
         Caption         =   "����..."
      End
      Begin VB.Menu mnuTrayAbout 
         Caption         =   "����..."
      End
      Begin VB.Menu mnuCut3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "�˳�����"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSoundPlay_Click()
    If chkSoundPlay.Value = 1 Then
        SoundPlayBoo = True
    Else
        SoundPlayBoo = False
    End If
    SoundPlayInt = chkSoundPlay.Value
    
    WritePrivateProfileString "Sound", "SoundPlay", CStr(chkSoundPlay.Value), App.path & "\ScreenSnapConfig.ini"
End Sub

Private Sub cmbZoom_Click()
    On Error Resume Next
    Dim Boo As Boolean
    Boo = ActiveForm.PictureSaved                                               '��¼�Ŵ�ǰ�Ƿ񱣴�
    If frmPicNum = -1 Then Exit Sub
    If ActiveForm.picScreenShot.Picture = 0 Then Exit Sub
    Dim X1 As Long, Y1 As Long, X0 As Integer, Y0 As Integer, OldStretchBltMode As Integer
    Set ActiveForm.picScreenShot.Picture = ActiveForm.PictureData
    X0 = ActiveForm.picScreenShot.Width / Screen.TwipsPerPixelX
    Y0 = ActiveForm.picScreenShot.Height / Screen.TwipsPerPixelY
    X1 = Val(cmbZoom.List(cmbZoom.ListIndex)) * 0.01 * ActiveForm.picScreenShot.Width
    Y1 = Val(cmbZoom.List(cmbZoom.ListIndex)) * 0.01 * ActiveForm.picScreenShot.Height
    ActiveForm.picScreenShot.Width = X1
    ActiveForm.picScreenShot.Height = Y1
    Set Picture1.Picture = ActiveForm.PictureData
    OldStretchBltMode = SetStretchBltMode(ActiveForm.picScreenShot.hDC, COLORONCOLOR) '�����µ�ģʽ
    StretchBlt ActiveForm.picScreenShot.hDC, 0, 0, X1 / Screen.TwipsPerPixelX, Y1 / Screen.TwipsPerPixelY, Picture1.hDC, 0, 0, X0, Y0, vbSrcCopy
    SetStretchBltMode ActiveForm.picScreenShot.hDC, OldStretchBltMode           '�Ļ�ԭ����ģʽ
    'ActiveForm.picScreenShot.PaintPicture DocData(ActiveForm.FormNum).PictureData _
    , 0, 0, X1, Y1
    ActiveForm.cmdTransferVHScroll.Value = True
    ActiveForm.PicZoom = Val(cmbZoom.List(cmbZoom.ListIndex))
    ActiveForm.PictureSaved = Boo                                               '�ָ��Ŵ�ǰ�Ƿ񱣴�����
End Sub

Private Sub cmbZoom_Scroll()
    cmbZoom_Click
End Sub

Private Sub imgActiveWin_Click()
    SnapSub 1
End Sub

Private Sub imgActiveWin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgActiveWin.Picture = imgSideBarPic(3).Picture                         '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgActiveWin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor hHandCur                                                          '����ָ��
End Sub

Private Sub imgActiveWin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgActiveWin.Picture = imgSideBarPic(2).Picture                         '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgAnyCtrlWindow_Click()
    frmAnyWindowCtrl.Show
End Sub

Private Sub imgAnyCtrlWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgAnyCtrlWindow.Picture = imgSideBarPic(7).Picture                        '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgAnyCtrlWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor hHandCur                                                          '����ָ��
End Sub

Private Sub imgAnyCtrlWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgAnyCtrlWindow.Picture = imgSideBarPic(6).Picture                        '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgCursor_Click()
    SnapSub 3
End Sub

Private Sub imgCursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgCursor.Picture = imgSideBarPic(5).Picture                            '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgCursor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor hHandCur                                                          '����ָ��
End Sub

Private Sub imgCursor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgCursor.Picture = imgSideBarPic(4).Picture                            '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgScreenSnap_Click()
    SnapSub 0
End Sub

Private Sub imgScreenSnap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgScreenSnap.Picture = imgSideBarPic(1).Picture                        '�������ʱ�ı�ͼƬ
End Sub

Private Sub imgScreenSnap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor hHandCur                                                          '����ָ��
End Sub

Private Sub imgScreenSnap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgScreenSnap.Picture = imgSideBarPic(0).Picture                        '�������ʱ�ı�ͼƬ
End Sub

Private Sub labMousePos_DblClick()
    InputBox LoadResString(10801), "", labMousePos.Caption
End Sub

Private Sub labPicQuantity_DblClick()
    InputBox LoadResString(10800), "", labPicQuantity.Caption
End Sub

Private Sub listSnapPic_Click()
    If SnapWhenTrayBoo = False Then
        PictureForms.Item(1 + listSnapPic.ListIndex).Caption = PictureForms.Item(1 + listSnapPic.ListIndex).PictureName
        If PictureForms.Item(1 + listSnapPic.ListIndex).Visible = False Then PictureForms.Item(1 + listSnapPic.ListIndex).Show
        PictureForms.Item(1 + listSnapPic.ListIndex).SetFocus
        PictureForms.Item(1 + listSnapPic.ListIndex).FormNum = listSnapPic.ListIndex
    End If
End Sub

Private Sub listSnapPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '����������������������������������������listSnapPic.ToolTipText��ʾ��Ϣ
    Dim LstPosNum As Long
    LstPosNum = SendMessage(listSnapPic.hwnd, LB_ITEMFROMPOINT, 0, _
    ByVal ((CLng(Y / Screen.TwipsPerPixelY) * 65536) + CLng(X / Screen.TwipsPerPixelX)))
    
    If (LstPosNum >= 0) And (LstPosNum <= listSnapPic.ListCount) Then           '������б�հ�����ֵΪ65536����65536С�ڵ�������������ô��ʾ�ı��͵���List(LstPOS)
        listSnapPic.ToolTipText = listSnapPic.List(LstPosNum)
    Else
        listSnapPic.ToolTipText = ""
    End If
    '����������������������������������������
End Sub

Private Sub MDIForm_Initialize()
    APPPrevInstance                                                             '��ֹ����������
    
    InitCommonControls                                                          'XP��ʽ��ʼ��
    
    RtlGetNtVersionNumbers SysMajor, SysMinor, SysBuild                         '��ȡϵͳ�汾
End Sub

Private Sub MDIForm_Load()
    If Dir(App.path & "\GdiPlus.dll") = "" Then
        '�Զ����Ʒ�ʽд�����ɣ�����ǰĿ¼
        Open App.path & "\GdiPlus.dll" For Binary As #1
        Put #1, , LoadResData(101, "CUSTOM")
        Close #1
    End If
    
    frmTray.Show
    
    Select Case Command
    Case "AUTORUN"
        Me.Visible = False
    End Select
    
    LoadLanguages "frmMain"                                                     '��ȡ����
    
    '��������������������������������������ȡini
    Dim retstrini As String
    If Dir(App.path & "\ScreenSnapConfig.ini") = "" Then
        WritePrivateProfileString "Sound", "ChooseSoundPlay", "DAZIJI", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Sound", "SoundPlay", "1", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "DelayTimeFullScreen", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "DelayTimeActiveWindow", "3", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "DelayTimeHotKey", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "DelayTimeCursor", "1", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "DelayTimeWindowCtrl", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Picture", "SaveJpgQuality", "80", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "ActiveWindowSnapMode", "1", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "EndOrMin", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "HotKey", "HotKeyCode", "122", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Save", "AutoSaveSnapFullScreen", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Save", "AutoSaveSnapActiveWindow", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Save", "AutoSaveSnapHotKey", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Save", "AutoSaveSnapCursor", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Save", "AutoSaveSnapWindowCtrl", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "HideWinCaptureFullScreen", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "HideWinCaptureActiveWindow", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "HideWinCaptureCursor", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "HideWinCaptureWindowCtrl", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Picture", "AutoSaveSnapPath", App.path, App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Picture", "AutoSaveSnapFormat", "*.bmp", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "AutoSendToClipBoard", "False", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "IncludeCursor", "False", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "HotKey", "DeclareHotKeyWay", "1", App.path & "\ScreenSnapConfig.ini"
    End If
    
    '��ȡע���ȼ���ʽ
    DeclareHotKeyWayInt = GetPrivateProfileInt("HotKey", "DeclareHotKeyWay", 1, App.path & "\ScreenSnapConfig.ini")
    '��ȡ�ȼ�
    HotKeyCodeInt = GetPrivateProfileInt("HotKey", "HotKeyCode", 122, App.path & "\ScreenSnapConfig.ini")
    '��ȡ�ر�������ʱֱ���˳���������С��������
    EndOrMinBoo = Abs(GetPrivateProfileInt("Config", "EndOrMin", 0, App.path & "\ScreenSnapConfig.ini"))
    '��ȡ����ڽ�ͼ��ʽ
    ActiveWindowSnapMode = GetPrivateProfileInt("Config", "ActiveWindowSnapMode", 1, App.path & "\ScreenSnapConfig.ini")
    '��ȡ�Ƿ��ȡ���
    retstrini = String(255, 0)
    GetPrivateProfileString "Config", "IncludeCursor", "NoData", retstrini, 256, App.path & "\ScreenSnapConfig.ini"
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "NoData" Then
        IncludeCursorBoo = False
        
        WritePrivateProfileString "Config", "IncludeCursor", "False", App.path & "\ScreenSnapConfig.ini"
    Else
        If retstrini = "True" Then
            IncludeCursorBoo = True
        Else
            IncludeCursorBoo = False
        End If
    End If
    '��ȡ�ȼ���ͼ���Ƿ�ֱ�ӽ�ͼƬ���Ƶ�������
    retstrini = String(255, 0)
    GetPrivateProfileString "Config", "AutoSendToClipBoard", "NoData", retstrini, 256, App.path & "\ScreenSnapConfig.ini"
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "NoData" Then
        AutoSendToClipBoardBoo = False
        
        WritePrivateProfileString "Config", "AutoSendToClipBoard", "False", App.path & "\ScreenSnapConfig.ini"
    Else
        If retstrini = "True" Then
            AutoSendToClipBoardBoo = True
        Else
            AutoSendToClipBoardBoo = False
        End If
    End If
    '��ȡ�Զ������ͼ��ʽ
    retstrini = String(255, 0)
    GetPrivateProfileString "Picture", "AutoSaveSnapFormat", "NoData", retstrini, 256, App.path & "\ScreenSnapConfig.ini"
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "NoData" Then
        AutoSaveSnapFormatStr = "*.bmp"
        
        WritePrivateProfileString "Picture", "AutoSaveSnapFormat", "*.bmp", App.path & "\ScreenSnapConfig.ini"
    Else
        AutoSaveSnapFormatStr = retstrini
    End If
    '��ȡ�Զ������ͼĿ¼
    retstrini = String(255, 0)
    GetPrivateProfileString "Picture", "AutoSaveSnapPath", "NoData", retstrini, 256, App.path & "\ScreenSnapConfig.ini"
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "NoData" Then
        AutoSaveSnapPathStr = App.path
        
        WritePrivateProfileString "Picture", "AutoSaveSnapPath", App.path, App.path & "\ScreenSnapConfig.ini"
    Else
        AutoSaveSnapPathStr = retstrini
    End If
    '��ȡȫ����ͼʱ�Ƿ����ش���
    HideWinCaptureInt(0) = GetPrivateProfileInt("Config", "HideWinCaptureFullScreen", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ����ڽ�ͼʱ�Ƿ����ش���
    HideWinCaptureInt(1) = GetPrivateProfileInt("Config", "HideWinCaptureActiveWindow", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ��ȡ���ʱ�Ƿ����ش���
    HideWinCaptureInt(3) = GetPrivateProfileInt("Config", "HideWinCaptureCursor", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ��ȡ���ⴰ��ʱ�Ƿ����ش���
    HideWinCaptureInt(4) = GetPrivateProfileInt("Config", "HideWinCaptureWindowCtrl", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡȫ����ͼ���Ƿ񱣴�
    AutoSaveSnapInt(0) = GetPrivateProfileInt("Save", "AutoSaveSnapFullScreen", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ����ڽ�ͼ���Ƿ񱣴�
    AutoSaveSnapInt(1) = GetPrivateProfileInt("Save", "AutoSaveSnapActiveWindow", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ�ȼ���ͼ���Ƿ񱣴�
    AutoSaveSnapInt(2) = GetPrivateProfileInt("Save", "AutoSaveSnapHotKey", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ��ȡ�����Ƿ񱣴�
    AutoSaveSnapInt(3) = GetPrivateProfileInt("Save", "AutoSaveSnapCursor", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ��ȡ���ⴰ�ں��Ƿ񱣴�
    AutoSaveSnapInt(4) = GetPrivateProfileInt("Save", "AutoSaveSnapWindowCtrl", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ������ʾ��ֵ
    SoundPlayInt = GetPrivateProfileInt("Sound", "SoundPlay", 1, App.path & "\ScreenSnapConfig.ini")
    chkSoundPlay.Value = SoundPlayInt
    '��ȡѡ�����ʾ��
    retstrini = String(255, 0)
    GetPrivateProfileString "Sound", "ChooseSoundPlay", "NoData", retstrini, 256, App.path & "\ScreenSnapConfig.ini"
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "NoData" Then
        ChooseSoundPlayStr = "DAZIJI"
        
        WritePrivateProfileString "Sound", "ChooseSoundPlay", "DAZIJI", App.path & "\ScreenSnapConfig.ini"
    Else
        ChooseSoundPlayStr = retstrini
    End If
    '��ȡȫ����ͼ��ʱֵ
    DelayTimeInt(0) = GetPrivateProfileInt("Config", "DelayTimeFullScreen", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ����ڽ�ͼ��ʱֵ
    DelayTimeInt(1) = GetPrivateProfileInt("Config", "DelayTimeActiveWindow", 3, App.path & "\ScreenSnapConfig.ini")
    '��ȡ�ȼ���ͼ��ʱֵ
    DelayTimeInt(2) = GetPrivateProfileInt("Config", "DelayTimeHotKey", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ��������ʱֵ
    DelayTimeInt(3) = GetPrivateProfileInt("Config", "DelayTimeCursor", 1, App.path & "\ScreenSnapConfig.ini")
    '��ȡ�������ⴰ����ʱֵ
    DelayTimeInt(4) = GetPrivateProfileInt("Config", "DelayTimeWindowCtrl", 0, App.path & "\ScreenSnapConfig.ini")
    '��ȡ����JPGͼƬѹ��Ʒ��ֵ
    SetJpgQuality = GetPrivateProfileInt("Picture", "SaveJpgQuality", 80, App.path & "\ScreenSnapConfig.ini")
    '������������������������������������
    
    frmPicNum = -1                                                              '���ĵ�
    
    hHandCur = LoadCursorA(0&, IDC_HAND)                                        '����ָ��
    
    TrayAddIcon frmMain, App.Title & " - " & LoadResString(10807) & vbNullChar  '��ε������������������½�һ��ͼ��
    
    'explorer����֮��㲥��һ�� windows message ��Ϣ
    MsgTaskbarRestart = RegisterWindowMessage("TaskbarCreated")
    
    With cmbZoom
        .AddItem "5%"
        .AddItem "10%"
        Dim i As Integer
        For i = 25 To 700 Step 25
            .AddItem i & "%"
        Next i
    End With
    
    cmbZoom.Text = "100%"
    
    imgScreenSnap.Picture = imgSideBarPic(0).Picture
    imgActiveWin.Picture = imgSideBarPic(2).Picture
    imgCursor.Picture = imgSideBarPic(4).Picture
    imgAnyCtrlWindow.Picture = imgSideBarPic(6).Picture
    '--------------------�ȼ�
    If DeclareHotKeyWayInt = 1 Then                                             'ϵͳ�ȼ�
        RegHotkeySub
    ElseIf DeclareHotKeyWayInt = 2 Then                                         '���̹���
        RegKeyHook
    End If
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '����---------------------
    If Me.Visible = False Then
        Select Case CLng(X / Screen.TwipsPerPixelX)
        Case WM_LBUTTONUP
            If Me.Enabled = False Then Exit Sub
            SetForegroundWindow Me.hwnd                                         '��������������㲻��ý���ʱ�����˵����Զ���ʧ
            
            ShowWindow Me.hwnd, SW_RESTORE
            
            If SnapWhenTrayLng <> 0 Then CreatPicsAfterTraySub
        Case WM_RBUTTONUP
            If Me.Enabled = False Then Exit Sub
            'If GetActiveWindow = hwnd Then Exit Sub
            SetForegroundWindow Me.hwnd
            
            PopupMenu mnuTray                                                   '�������˵�
        End Select
    End If
    '---------------------
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim isUnloadWindows As Boolean
    Select Case UnloadMode
    Case vbAppWindows                                                           ' 2 ��ǰ Microsoft Windows ���������Ự������
        EndOrMinBoo = True
        isUnloadWindows = True
        mnuTrayShow_Click
    End Select
    
    Dim i As Integer                                                            '����
    CloseFilesModeInt = 1                                                       '��־��Ҫ�˳��������Ա����Msgbox��������ʾ4����ť
    '����---------------------
    If EndOrMinBoo Then
        If frmPicNum > -1 Then listSnapPic.Selected(frmPicNum) = True           '���Ӵ��巢���ر�ʱ��ǰѡ���б�����һ�ȷ�������һ���ĵ����ιر�
        
        '�Ȱ��Ӵ���رգ���Ϊvb6Ĭ�ϵ�˳���Ǵ�0��n��������㷨˳��պ��෴�������Ҫ���ֶ��ر��Ӵ��壬�ٴ���������Ĺر��¼�
        For i = frmPicNum To 0 Step -1
            Unload PictureForms.Item(1 + i)
            If NewMsgBoxInt = -1 Then Cancel = 1: NewMsgBoxInt = 0: Exit For
            If NewMsgBoxInt = 4 Then NewMsgBoxInt = 0: Exit For                 '���Ӵ������Լ�ȫ�ر����ˣ�����Ҫ��ѭ��
            'PictureForms.Remove (1 + i)                                         '���˳��Ҫ������
        Next
        If isUnloadWindows Then TrayRemoveIcon: UnRegHotkeySub: End
    Else
        Cancel = True                                                           'ȡ���˳�
        CloseFilesModeInt = 0
        Me.Visible = False
    End If
    '---------------------
End Sub

Private Sub MDIForm_Resize()
    cmbZoom.Left = Me.ScaleWidth - cmbZoom.Width + picSideBar.Width + picSnapPic.Width
    labMousePos.Left = cmbZoom.Left - labMousePos.Width
    labPicQuantity.Left = labMousePos.Left - labPicQuantity.Width
    chkSoundPlay.Left = labPicQuantity.Left - chkSoundPlay.Width
    
    listSnapPic.Width = picSnapPic.ScaleWidth - listSnapPic.Left
    listSnapPic.Height = Me.ScaleHeight
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    TrayRemoveIcon                                                              '�˳�ʱɾ������ͼ��
    
    '---------------------�ȼ�
    UnRegHotkeySub
    
    End
End Sub

Private Sub mnuAbout_Click()
    frmAbout.ShowForm 1
End Sub

Private Sub mnuActiveWinSnap_Click()
    imgActiveWin_Click
End Sub

Private Sub mnuAnyWindowCtrlSnap_Click()
    imgAnyCtrlWindow_Click
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuCloseAllFilesUnsaved_Click()
    CloseAllFilesUnsaved
End Sub

Private Sub mnuCopy_Click()
    mnufrmPicCopy_Click
End Sub

Private Sub mnuCursorSnap_Click()
    imgCursor_Click
End Sub

Private Sub mnuExit_Click()
    EndOrMinBoo = True
    Unload Me
End Sub

Private Sub mnufrmPicClose_Click()
    Unload Me.ActiveForm
End Sub

Private Sub mnufrmPicCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetData PictureForms.Item(1 + listSnapPic.ListIndex).PictureData
End Sub

Private Sub mnufrmPicPaste_Click()
    If frmPicNum = -1 Or listSnapPic.ListIndex < 0 Then Exit Sub
    If Clipboard.GetFormat(2) Or Clipboard.GetFormat(3) Or Clipboard.GetFormat(8) Then
        If ActiveForm.PictureSaved Then                                         'Ҫ��Clipboard.GetData()֮ǰ
            ActiveForm.PictureSaved = False
            ActiveForm.PictureName = ActiveForm.PictureName & " *"
            ActiveForm.Caption = ActiveForm.PictureName
        End If
        
        Set ActiveForm.picScreenShot.Picture = LoadPicture()
        Set Me.ActiveForm.picScreenShot.Picture = Clipboard.GetData()           '����PictureSaved�ı�
        ActiveForm.picScreenShot.Picture = ActiveForm.picScreenShot.Image
        Set ActiveForm.PictureData = ActiveForm.picScreenShot.Picture
        ActiveForm.cmdTransferVHScroll.Value = True
        
        'listbox�ӡ�*��
        listSnapPic.List(listSnapPic.ListIndex) = ActiveForm.PictureName
        
        ActiveForm.PicZoom = 100
        cmbZoom.Text = "100%"
    End If
End Sub

Public Sub mnuNew_Click()
    frmPicNum = frmPicNum + 1
    PicFilesCount = PicFilesCount + 1
    
    Dim frmPicture As New frmPicture
    frmPicture.PicZoom = 100
    frmPicture.PictureName = LoadResString(10705) & PicFilesCount
    frmPicture.Caption = frmPicture.PictureName
    Set frmPicture.PictureData = frmPicture.picScreenShot.Picture
    PictureForms.Add frmPicture
    frmPicture.Show
    
    listSnapPic.AddItem frmPicture.PictureName
    listSnapPic.Selected(frmPicNum) = True
    
    TrayTip Me, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '��   �Ž�ͼ
    
    cmbZoom.Text = "100%"                                                       'Ҫ��frmPictureCopy(frmPicNum).Show֮��cmb���ܻ�ȡ���Ӵ���
End Sub

Private Sub mnuOpenTheFolder_Click()
    Shell "explorer.exe " & AutoSaveSnapPathStr, vbNormalFocus
End Sub

Private Sub mnuPaste_Click()
    mnufrmPicPaste_Click
End Sub

Private Sub mnuSave_Click()
    If frmPicNum = -1 Then Exit Sub                                             'û��ͼƬ
    Dim Str As String
    Str = GetPicturePath(Me, ActiveForm.FormNum)
    If SavePictures(Str, ActiveForm.FormNum) = 1 Then
        ActiveForm.PictureName = Str
        ActiveForm.Caption = Str
        listSnapPic.List(listSnapPic.ListIndex) = Str
    End If
End Sub

Private Sub mnuScreenSnap_Click()
    imgScreenSnap_Click
End Sub

Private Sub mnuSetting_Click()
    frmSettings.Show 1
End Sub

Private Sub mnuSourceCode_Click()
    ShellExecute Me.hwnd, "open", "https://github.com/SkyD666/VB6_ScreenShot", "", "", 1
End Sub

Private Sub mnuTrayAbout_Click()
    frmAbout.ShowForm 1
End Sub

Private Sub mnuTrayActiveWinSnap_Click()
    imgActiveWin_Click
End Sub

Private Sub mnuTrayCursorSnap_Click()
    imgCursor_Click
End Sub

Private Sub mnuTrayExit_Click()
    EndOrMinBoo = True
    mnuTrayShow_Click
    Unload Me
End Sub

Private Sub mnuTrayScreenSnap_Click()
    imgScreenSnap_Click
End Sub

Private Sub mnuTraySetting_Click()
    frmSettings.Show 1
End Sub

Private Sub mnuTrayShow_Click()
    If Me.Visible = False Then
        If Me.Enabled = False Then Exit Sub
        SetForegroundWindow Me.hwnd                                             '��������������㲻��ý���ʱ�����˵����Զ���ʧ
        
        ShowWindow Me.hwnd, SW_RESTORE
        
        If SnapWhenTrayLng <> 0 Then CreatPicsAfterTraySub
    End If
End Sub

Private Sub mnuTrayWinCtrlSnap_Click()
    imgAnyCtrlWindow_Click
End Sub

Private Sub mnuZoomIn_Click()
    If cmbZoom.ListIndex < frmMain.cmbZoom.ListCount - 1 Then cmbZoom.ListIndex = cmbZoom.ListIndex + 1
End Sub

Private Sub mnuZoomOut_Click()
    If cmbZoom.ListIndex > 0 Then cmbZoom.ListIndex = cmbZoom.ListIndex - 1
End Sub

Private Sub timerHotKey_Timer()
    On Error GoTo Err
    If frmMain.Visible = True Then
        SnapSub 2                                                               '����ǿɼ�
    Else
        '������ɼ�
        DelayTimeSub 2                                                          '��ʱ
        
        SnapWhenTrayLng = SnapWhenTrayLng + 1
        SnapWhenTrayBoo = True
        
        PicFilesCount = PicFilesCount + 1
        frmPicNum = frmPicNum + 1
        
        Dim frmPicture As New frmPicture
        frmPicture.PicZoom = 100
        frmPicture.PictureName = LoadResString(10705) & PicFilesCount
        frmPicture.Caption = frmPicture.PictureName
        Set frmPicture.PictureData = frmPicture.picScreenShot.Picture
        PictureForms.Add frmPicture
        
        If ActiveWindowSnapMode = 0 Then                                        '��ͬ�Ļ���ڽ�ͼ���� �Ƚ�ͼ���ٴ����ĵ���
            Set PictureForms.Item(1 + frmPicNum).PictureData = CaptureActiveWindow() '=ԭʼ����
        ElseIf ActiveWindowSnapMode = 1 Then
            Set PictureForms.Item(1 + frmPicNum).PictureData = CaptureActiveWindowB() '=�·���
        End If
        'PictureSaved(frmPicNum) = False                                      '�ĵ�������������
        PictureForms.Item(1 + frmPicNum).PicZoom = 100
        PictureForms.Item(1 + frmPicNum).PictureName = LoadResString(10705) & PicFilesCount & " *"
        PictureForms.Item(1 + frmPicNum).Caption = PictureForms.Item(1 + frmPicNum).PictureName
        Set PictureForms.Item(1 + frmPicNum).picScreenShot.Picture = PictureForms.Item(1 + frmPicNum).PictureData
        Set PictureForms.Item(1 + frmPicNum).picScreenShot.Picture = PictureForms.Item(1 + frmPicNum).picScreenShot.Image
        '������������������������������������������������� ��CaptureActiveWindow(B)�ڲ���
        Set PictureForms.Item(1 + frmPicNum).PictureData = PictureForms.Item(1 + frmPicNum).picScreenShot.Picture
        
        If AutoSendToClipBoardBoo Then Clipboard.Clear: Clipboard.SetData PictureForms.Item(1 + frmPicNum).PictureData '�ȼ���ͼ��ֱ�ӽ�ͼƬ���Ƶ�������
        
        If AutoSaveSnapInt(2) = 1 Then AutoSaveSnapSub 2, frmPicNum             '�Զ����棿
        
        cmbZoom.Text = "100%"                                                   'Ҫ��frmPictureCopy(frmPicNum).Show֮��cmb���ܻ�ȡ���Ӵ���
        
        If SoundPlayBoo Then SoundPlay                                          '�Ƿ񲥷���ʾ��
        
        TrayTip Me, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '��   �Ž�ͼ
    End If
    
    timerHotKey.Enabled = False
    Exit Sub
Err:
    MsgBox "����timerHotKey_Timer" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
End Sub

Public Sub CloseAllFilesUnsaved()
    If frmPicNum = -1 Then Exit Sub
    If NewMsgBoxInt = 4 Then GoTo pos
    
    '�ٴ�ȷ�ϴ˲����� ȫ���ر��Ҳ�����δ������ĵ�,�Ƿ����?
    If MsgBox(LoadResString(10808), vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
pos:
        Dim i As Integer
        listSnapPic.Selected(frmPicNum) = True                                  '���Ӵ��巢���ر�ʱ��ǰѡ���б�����һ�ȷ�������һ���ĵ����ιر�
        CloseAllFilesUnsavedBoo = True
        For i = frmPicNum To 0 Step -1
            Unload PictureForms.Item(1 + i)
        Next
        listSnapPic.Clear
        frmPicNum = -1
        
        CloseAllFilesUnsavedBoo = False
    End If
End Sub
