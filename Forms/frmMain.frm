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
      Begin VB.CommandButton cmdMainTran 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
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
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
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
Dim ScrWidth As Integer
Dim ScrHeight As Integer, msgboxed As Boolean

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
    Boo = DocData(CLng(ActiveForm.labfrmi.Caption)).frmPictureSaved             '��¼�Ŵ�ǰ�Ƿ񱣴�
    If frmPicNum = -1 Then Exit Sub
    If ActiveForm.picScreenShot.Picture = 0 Then Exit Sub
    Dim X1 As Single, Y1 As Single
    Set ActiveForm.picScreenShot.Picture = DocData(CLng(ActiveForm.labfrmi.Caption)).PictureData
    X1 = Val(cmbZoom.List(cmbZoom.ListIndex)) * 0.01 * ActiveForm.picScreenShot.Width
    Y1 = Val(cmbZoom.List(cmbZoom.ListIndex)) * 0.01 * ActiveForm.picScreenShot.Height
    Me.ActiveForm.picScreenShot.Width = X1
    Me.ActiveForm.picScreenShot.Height = Y1
    Me.ActiveForm.picScreenShot.PaintPicture DocData(CLng(ActiveForm.labfrmi.Caption)).PictureData _
    , 0, 0, X1, Y1
    Me.ActiveForm.cmdTransferVHScroll.Value = True
    DocData(CLng(ActiveForm.labfrmi.Caption)).PicZoom = Val(cmbZoom.List(cmbZoom.ListIndex))
    DocData(CLng(ActiveForm.labfrmi.Caption)).frmPictureSaved = Boo             '�ָ��Ŵ�ǰ�Ƿ񱣴�����
End Sub

Private Sub cmbZoom_Scroll()
    cmbZoom_Click
End Sub

Private Sub cmdMainTran_Click()
    mnuCloseAllFilesUnsaved_Click
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
    On Error Resume Next
    
    If SnapWhenTrayBoo = False Then
        DocData(listSnapPic.ListIndex).frmPictureCopy.Caption = DocData(listSnapPic.ListIndex).frmPictureName
        If DocData(listSnapPic.ListIndex).frmPictureCopy.Visible = False Then DocData(listSnapPic.ListIndex).frmPictureCopy.Show
        DocData(listSnapPic.ListIndex).frmPictureCopy.SetFocus
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
    frmTray.Show
    
    Select Case Command
    Case "AUTORUN"
        Me.Visible = False
    End Select
    
    '---------------------------------------��ȡ����
    LoadLanguages "frmMain"
    '    Select Case GetSystemDefaultLangID
    '    Case &H804                                                                  '��������
    '        LoadLanguages 1
    '    Case &H404                                                                  '��������
    '        LoadLanguages 2
    '    Case &H409                                                                  'Ӣ��
    '        LoadLanguages 3
    '    End Select
    
    '---------------------------------------
    
    '��������������������������������������ȡini
    Dim lngini As Long
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
        'WritePrivateProfileString "Config", "HideWinCaptureHotKey", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "HideWinCaptureCursor", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "HideWinCaptureWindowCtrl", "0", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Picture", "AutoSaveSnapPath", App.path, App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Picture", "AutoSaveSnapFormat", "*.bmp", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "AutoSendToClipBoard", "False", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "Config", "IncludeCursor", "False", App.path & "\ScreenSnapConfig.ini"
        WritePrivateProfileString "HotKey", "DeclareHotKeyWay", "1", App.path & "\ScreenSnapConfig.ini"
    End If
    
    '��ȡע���ȼ���ʽ
    lngini = GetPrivateProfileInt("HotKey", "DeclareHotKeyWay", 1, App.path & "\ScreenSnapConfig.ini")
    DeclareHotKeyWayInt = lngini
    '��ȡ�ȼ�
    lngini = GetPrivateProfileInt("HotKey", "HotKeyCode", 122, App.path & "\ScreenSnapConfig.ini")
    HotKeyCodeInt = lngini
    '��ȡ�ر�������ʱֱ���˳���������С��������
    lngini = GetPrivateProfileInt("Config", "EndOrMin", 0, App.path & "\ScreenSnapConfig.ini")
    EndOrMinBoo = Abs(lngini)
    '��ȡ����ڽ�ͼ��ʽ
    lngini = GetPrivateProfileInt("Config", "ActiveWindowSnapMode", -1, App.path & "\ScreenSnapConfig.ini")
    If lngini = -1 Then
        ActiveWindowSnapMode = 1
        
        lngini = WritePrivateProfileString("Config", "ActiveWindowSnapMode", "1", App.path & "\ScreenSnapConfig.ini")
    Else
        ActiveWindowSnapMode = lngini
    End If
    '��ȡ�Ƿ��ȡ���
    retstrini = String(260, 0)
    lngini = GetPrivateProfileString("Config", "IncludeCursor", "δ�ҵ�����", retstrini, 245, App.path & "\ScreenSnapConfig.ini")
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "δ�ҵ�����" Then
        IncludeCursorBoo = False
        
        lngini = WritePrivateProfileString("Config", "IncludeCursor", "False", App.path & "\ScreenSnapConfig.ini")
    Else
        If retstrini = "True" Then
            IncludeCursorBoo = True
        Else
            IncludeCursorBoo = False
        End If
    End If
    '��ȡ�ȼ���ͼ���Ƿ�ֱ�ӽ�ͼƬ���Ƶ�������
    retstrini = String(260, 0)
    lngini = GetPrivateProfileString("Config", "AutoSendToClipBoard", "δ�ҵ�����", retstrini, 245, App.path & "\ScreenSnapConfig.ini")
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "δ�ҵ�����" Then
        AutoSendToClipBoardBoo = False
        
        lngini = WritePrivateProfileString("Config", "AutoSendToClipBoard", "False", App.path & "\ScreenSnapConfig.ini")
    Else
        If retstrini = "True" Then
            AutoSendToClipBoardBoo = True
        Else
            AutoSendToClipBoardBoo = False
        End If
    End If
    '��ȡ�Զ������ͼ��ʽ
    retstrini = String(260, 0)
    lngini = GetPrivateProfileString("Picture", "AutoSaveSnapFormat", "δ�ҵ�����", retstrini, 245, App.path & "\ScreenSnapConfig.ini")
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "δ�ҵ�����" Then
        AutoSaveSnapFormatStr = "*.bmp"
        
        lngini = WritePrivateProfileString("Picture", "AutoSaveSnapFormat", "*.bmp", App.path & "\ScreenSnapConfig.ini")
    Else
        AutoSaveSnapFormatStr = retstrini
    End If
    '��ȡ�Զ������ͼĿ¼
    retstrini = String(260, 0)
    lngini = GetPrivateProfileString("Picture", "AutoSaveSnapPath", "δ�ҵ�����", retstrini, 245, App.path & "\ScreenSnapConfig.ini")
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "δ�ҵ�����" Then
        AutoSaveSnapPathStr = App.path
        
        lngini = WritePrivateProfileString("Picture", "AutoSaveSnapPath", App.path, App.path & "\ScreenSnapConfig.ini")
    Else
        AutoSaveSnapPathStr = retstrini
    End If
    '��ȡȫ����ͼʱ�Ƿ����ش���
    lngini = GetPrivateProfileInt("Config", "HideWinCaptureFullScreen", 0, App.path & "\ScreenSnapConfig.ini")
    HideWinCaptureInt(0) = lngini
    '��ȡ����ڽ�ͼʱ�Ƿ����ش���
    lngini = GetPrivateProfileInt("Config", "HideWinCaptureActiveWindow", 0, App.path & "\ScreenSnapConfig.ini")
    HideWinCaptureInt(1) = lngini
    '    '��ȡ�ȼ���ͼʱ�Ƿ����ش���
    '    lngini = GetPrivateProfileInt("Config", "HideWinCaptureHotKey", 0, App.path & "\ScreenSnapConfig.ini")
    '    HideWinCaptureInt(2) = lngini
    '��ȡ��ȡ���ʱ�Ƿ����ش���
    lngini = GetPrivateProfileInt("Config", "HideWinCaptureCursor", 0, App.path & "\ScreenSnapConfig.ini")
    HideWinCaptureInt(3) = lngini
    '��ȡ��ȡ���ⴰ��ʱ�Ƿ����ش���
    lngini = GetPrivateProfileInt("Config", "HideWinCaptureWindowCtrl", 0, App.path & "\ScreenSnapConfig.ini")
    HideWinCaptureInt(4) = lngini
    '��ȡȫ����ͼ���Ƿ񱣴�
    lngini = GetPrivateProfileInt("Save", "AutoSaveSnapFullScreen", 0, App.path & "\ScreenSnapConfig.ini")
    AutoSaveSnapInt(0) = lngini
    '��ȡ����ڽ�ͼ���Ƿ񱣴�
    lngini = GetPrivateProfileInt("Save", "AutoSaveSnapActiveWindow", 0, App.path & "\ScreenSnapConfig.ini")
    AutoSaveSnapInt(1) = lngini
    '��ȡ�ȼ���ͼ���Ƿ񱣴�
    lngini = GetPrivateProfileInt("Save", "AutoSaveSnapHotKey", 0, App.path & "\ScreenSnapConfig.ini")
    AutoSaveSnapInt(2) = lngini
    '��ȡ��ȡ�����Ƿ񱣴�
    lngini = GetPrivateProfileInt("Save", "AutoSaveSnapCursor", 0, App.path & "\ScreenSnapConfig.ini")
    AutoSaveSnapInt(3) = lngini
    '��ȡ��ȡ���ⴰ�ں��Ƿ񱣴�
    lngini = GetPrivateProfileInt("Save", "AutoSaveSnapWindowCtrl", 0, App.path & "\ScreenSnapConfig.ini")
    AutoSaveSnapInt(4) = lngini
    '��ȡ������ʾ��ֵ
    lngini = GetPrivateProfileInt("Sound", "SoundPlay", 3, App.path & "\ScreenSnapConfig.ini")
    If lngini = 3 Then
        SoundPlayInt = 1
        chkSoundPlay.Value = SoundPlayInt
        
        lngini = WritePrivateProfileString("Sound", "SoundPlay", "1", App.path & "\ScreenSnapConfig.ini")
    Else
        SoundPlayInt = lngini
        chkSoundPlay.Value = SoundPlayInt
    End If
    '��ȡѡ�����ʾ��
    retstrini = String(260, 0)
    lngini = GetPrivateProfileString("Sound", "ChooseSoundPlay", "δ�ҵ�����", retstrini, 245, App.path & "\ScreenSnapConfig.ini")
    retstrini = Replace(retstrini, Chr(0), "")
    If retstrini = "δ�ҵ�����" Then
        ChooseSoundPlayStr = "DAZIJI"
        
        lngini = WritePrivateProfileString("Sound", "ChooseSoundPlay", "DAZIJI", App.path & "\ScreenSnapConfig.ini")
    Else
        ChooseSoundPlayStr = retstrini
    End If
    '��ȡȫ����ͼ��ʱֵ
    lngini = GetPrivateProfileInt("Config", "DelayTimeFullScreen", 0, App.path & "\ScreenSnapConfig.ini")
    DelayTimeInt(0) = lngini
    '��ȡ����ڽ�ͼ��ʱֵ
    lngini = GetPrivateProfileInt("Config", "DelayTimeActiveWindow", 3, App.path & "\ScreenSnapConfig.ini")
    DelayTimeInt(1) = lngini
    '��ȡ�ȼ���ͼ��ʱֵ
    lngini = GetPrivateProfileInt("Config", "DelayTimeHotKey", 0, App.path & "\ScreenSnapConfig.ini")
    DelayTimeInt(2) = lngini
    '��ȡ��������ʱֵ
    lngini = GetPrivateProfileInt("Config", "DelayTimeCursor", 1, App.path & "\ScreenSnapConfig.ini")
    DelayTimeInt(3) = lngini
    '��ȡ�������ⴰ����ʱֵ
    lngini = GetPrivateProfileInt("Config", "DelayTimeWindowCtrl", 0, App.path & "\ScreenSnapConfig.ini")
    DelayTimeInt(4) = lngini
    '��ȡ����JPGͼƬѹ��Ʒ��ֵ
    lngini = GetPrivateProfileInt("Picture", "SaveJpgQuality", -1, App.path & "\ScreenSnapConfig.ini")
    If lngini = -1 Then
        SetJpgQuality = 80
        
        lngini = WritePrivateProfileString("Picture", "SaveJpgQuality", "80", App.path & "\ScreenSnapConfig.ini")
    Else
        SetJpgQuality = lngini
    End If
    '������������������������������������
    
    frmPicNum = -1                                                              '���ĵ�
    
    hHandCur = LoadCursorA(0&, IDC_HAND)                                        '����ָ��
    
    TrayAddIcon frmMain, App.Title & " - " & LoadResString(10807) & vbNullChar  '��ε������������������½�һ��ͼ��
    
    'explorer����֮��㲥��һ�� windows message ��Ϣ
    MsgTaskbarRestart = RegisterWindowMessage("TaskbarCreated")
    
    With cmbZoom
        .AddItem "5%"
        .AddItem "10%"
        .AddItem "25%"
        .AddItem "50%"
        .AddItem "75%"
        .AddItem "90%"
        .AddItem "100%"
        .AddItem "150%"
        .AddItem "200%"
        .AddItem "250%"
        .AddItem "300%"
        .AddItem "350%"
        .AddItem "400%"
        .AddItem "450%"
        .AddItem "500%"
        .AddItem "550%"
        .AddItem "600%"
        .AddItem "650%"
        .AddItem "700%"
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
            Unload DocData(i).frmPictureCopy
            If NewMsgBoxInt = -1 Then Cancel = 1: NewMsgBoxInt = 0: Exit For
            If NewMsgBoxInt = 4 Then NewMsgBoxInt = 0: Exit For                 '���Ӵ������Լ�ȫ�ر����ˣ�����Ҫ��ѭ��
        Next
        
        '        If NewMsgBoxInt = 4 Then NewMsgBoxInt = 0: Exit Sub                     '��ȫ����ʱͨ���˵������ �ر������ļ������� ���˳�
        '        If NewMsgBoxInt = -1 Then Cancel = True: NewMsgBoxInt = 0: Exit Sub     '��ȡ��ʱֹͣ�˳�
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
    If frmPicNum = -1 Then Exit Sub
    If NewMsgBoxInt = 4 Then GoTo pos
    
    '�ٴ�ȷ�ϴ˲����� ȫ���ر��Ҳ�����δ������ĵ�,�Ƿ����?
    If MsgBox(LoadResString(10808), vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
pos:
        Dim i As Integer
        listSnapPic.Selected(frmPicNum) = True                                  '���Ӵ��巢���ر�ʱ��ǰѡ���б�����һ�ȷ�������һ���ĵ����ιر�
        CloseAllFilesUnsavedBoo = True
        For i = frmPicNum To 0 Step -1
            Unload DocData(i).frmPictureCopy
        Next
        listSnapPic.Clear
        Erase DocData
        
        CloseAllFilesUnsavedBoo = False
    End If
End Sub

Private Sub mnuCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetData Me.ActiveForm.picScreenShot.Picture
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
    Clipboard.SetData DocData(CLng(ActiveForm.labfrmi.Caption)).PictureData
End Sub

Private Sub mnufrmPicPaste_Click()
    Dim SelectedInt As Long
    If frmPicNum = -1 Then Exit Sub
    If Clipboard.GetFormat(2) Or Clipboard.GetFormat(3) Or Clipboard.GetFormat(8) Then
        If DocData(CInt(Me.ActiveForm.labfrmi.Caption)).frmPictureSaved Then    'Ҫ��Clipboard.GetData()֮ǰ
            DocData(CInt(Me.ActiveForm.labfrmi.Caption)).frmPictureSaved = False
            DocData(CInt(Me.ActiveForm.labfrmi.Caption)).frmPictureName = DocData(CInt(Me.ActiveForm.labfrmi.Caption)).frmPictureName & " *"
            Me.ActiveForm.Caption = DocData(CInt(Me.ActiveForm.labfrmi.Caption)).frmPictureName
        End If
        
        Me.ActiveForm.picScreenShot.Picture = LoadPicture()
        Set Me.ActiveForm.picScreenShot.Picture = Clipboard.GetData()           '����frmPictureSaved�ı�
        Me.ActiveForm.picScreenShot.Picture = Me.ActiveForm.picScreenShot.Image
        Set DocData(frmPicNum).PictureData = DocData(frmPicNum).frmPictureCopy.picScreenShot.Picture
        Me.ActiveForm.cmdTransferVHScroll.Value = True
        
        'listbox�ӡ�*��
        listSnapPic.AddItem DocData(CInt(Me.ActiveForm.labfrmi.Caption)).frmPictureName, listSnapPic.ListIndex
        SelectedInt = listSnapPic.ListIndex - 1
        listSnapPic.RemoveItem listSnapPic.ListIndex
        listSnapPic.Selected(SelectedInt) = True
        
        DocData(CInt(Me.ActiveForm.labfrmi.Caption)).PicZoom = 100
        cmbZoom.Text = "100%"
    End If
End Sub

Public Sub mnuNew_Click()
    frmPicNum = frmPicNum + 1
    PicFilesCount = PicFilesCount + 1
    ReDim Preserve DocData(0 To frmPicNum) As DocumentsData
    DocData(frmPicNum).PicZoom = 100
    DocData(frmPicNum).frmPictureCopy.Show
    DocData(frmPicNum).frmPictureName = LoadResString(10705) & PicFilesCount
    DocData(frmPicNum).frmPictureCopy.Caption = DocData(frmPicNum).frmPictureName
    Set DocData(frmPicNum).PictureData = DocData(frmPicNum).frmPictureCopy.picScreenShot.Picture
    
    listSnapPic.AddItem DocData(frmPicNum).frmPictureName
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
    Str = SaveFiles(Me, CLng(Me.ActiveForm.labfrmi.Caption))
    SaveFiles2 Str, CLng(Me.ActiveForm.labfrmi.Caption)
End Sub

Private Sub mnuScreenSnap_Click()
    imgScreenSnap_Click
End Sub

Private Sub mnuSetting_Click()
    frmSettings.Show 1
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
        ReDim Preserve DocData(0 To frmPicNum) As DocumentsData
        If ActiveWindowSnapMode = 0 Then                                        '��ͬ�Ļ���ڽ�ͼ���� �Ƚ�ͼ���ٴ����ĵ���
            Set DocData(frmPicNum).PictureData = CaptureActiveWindow()          '=ԭʼ����
        ElseIf ActiveWindowSnapMode = 1 Then
            Set DocData(frmPicNum).PictureData = CaptureActiveWindowB()         '=�·���
        End If
        'frmPictureSaved(frmPicNum) = False                                      '�ĵ�������������
        DocData(frmPicNum).PicZoom = 100
        DocData(frmPicNum).frmPictureName = LoadResString(10705) & PicFilesCount & " *"
        DocData(frmPicNum).frmPictureCopy.Caption = DocData(frmPicNum).frmPictureName
        DocData(frmPicNum).frmPictureCopy.labfrmi.Caption = frmPicNum
        Set DocData(frmPicNum).frmPictureCopy.picScreenShot.Picture = DocData(frmPicNum).PictureData
        Set DocData(frmPicNum).frmPictureCopy.picScreenShot.Picture = DocData(frmPicNum).frmPictureCopy.picScreenShot.Image
        '������������������������������������������������� ��CaptureActiveWindow(B)�ڲ���
        Set DocData(frmPicNum).PictureData = DocData(frmPicNum).frmPictureCopy.picScreenShot.Picture
        
        If AutoSendToClipBoardBoo Then Clipboard.Clear: Clipboard.SetData DocData(frmPicNum).PictureData '�ȼ���ͼ��ֱ�ӽ�ͼƬ���Ƶ�������
        
        If AutoSaveSnapInt(2) = 1 Then AutoSaveSnapSub 2, frmPicNum             '�Զ����棿
        
        cmbZoom.Text = "100%"                                                   'Ҫ��frmPictureCopy(frmPicNum).Show֮��cmb���ܻ�ȡ���Ӵ���
        
        If SoundPlayBoo Then SoundPlay                                          '�Ƿ񲥷���ʾ��
        
        TrayTip Me, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '��   �Ž�ͼ
    End If
    
    timerHotKey.Enabled = False
    Exit Sub
Err:
    MsgBox "����" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
End Sub
