VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ҵ�Ӧ�ó���"
   ClientHeight    =   5190
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5055
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3582.229
   ScaleMode       =   0  'User
   ScaleWidth      =   4746.906
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.HScrollBar sclVol 
      Height          =   319
      LargeChange     =   16
      Left            =   720
      Max             =   64
      TabIndex        =   6
      Top             =   4695
      Value           =   64
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2835
      Left            =   1080
      ScaleHeight     =   2835
      ScaleWidth      =   3855
      TabIndex        =   5
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   360
   End
   Begin VB.CheckBox chkMusic 
      Caption         =   "&Music Funky Stars 2:20"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   4230
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Top             =   4260
      Width           =   1485
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   4695
      Width           =   1485
   End
   Begin VB.Label labSysVer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ�汾��0.0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vol:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   4695
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4620.134
      Y1              =   2847.147
      Y2              =   2847.147
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�ó������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1065
      TabIndex        =   2
      Top             =   120
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4620.134
      Y1              =   2857.501
      Y2              =   2857.501
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�汾"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3795
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------������Ļ
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As Rect, lprcClip As Rect, ByVal hrgnUpdate As Long, lprcUpdate As Rect) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '��ȡ�����ĵ��ַ������ȣ��õ���chr(0)Ϊ��β���ַ������ֽ���
Dim Canceled As Boolean
Dim Scrolling As Boolean                                                        'Scroll flag
Dim m_View As Boolean

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'------------------

' ע���ؼ��ְ�ȫѡ��...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
    KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
    KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
    ' ע���ؼ��� ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                                                                ' �����Ŀյ��ս��ַ���
Const REG_DWORD = 4                                                             ' 32λ����

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub chkMusic_Click()
    Select Case chkMusic.Value
    Case 1
        uFMOD_PlaySong 1, 0, XM_RESOURCE                                        '��������
        uFMOD_SetVolume sclVol.Value
    Case 0
        uFMOD_PlaySong 0, 0, 0                                                  'ֹͣ����
    End Select
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Timer1.Enabled = False
    Canceled = False
    Scrolling = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sysverstr As String
    LoadLanguages "frmAbout"
    
    Me.Icon = frmMain.Icon
    Me.Caption = LoadResString(11002) & " " & App.Title
    lblVersion.Caption = LoadResString(11003) & ":" & App.Major & "." & App.Minor & "." & App.Revision & "." & "200324"
    lblTitle.Caption = App.Title
    
    Select Case CStr(SysMajor & "." & SysMinor)
    Case "5.1"
        sysverstr = "Windows XP"
    Case "6.0"
        sysverstr = "Windows Vista"
    Case "6.1"
        sysverstr = "Windows 7"
    Case "6.2"
        sysverstr = "Windows 8"
    Case "6.3"
        sysverstr = "Windows 8.1"
    Case "10.0"
        sysverstr = "Windows 10"
    End Select
    
    labSysVer.Caption = LoadResString(11004) & ":" & sysverstr & " (" & SysMajor & "." & SysMinor & ")"
    
    chkMusic_Click                                                              '��������
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ��ͼ��ע����л��ϵͳ��Ϣ�����·��������...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' ��ͼ����ע����л��ϵͳ��Ϣ�����·��...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' ��֪32λ�ļ��汾����Чλ��
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' ���� - �ļ����ܱ��ҵ�...
        Else
            GoTo SysInfoErr
        End If
        ' ���� - ע�����Ӧ��Ŀ���ܱ��ҵ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox LoadResString(11005), vbOKOnly                                       '��ʱϵͳ��Ϣ������
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                                               ' ѭ��������
    Dim rc As Long                                                              ' ���ش���
    Dim hKey As Long                                                            ' �򿪵�ע���ؼ��־��
    Dim hDepth As Long                                                          '
    Dim KeyValType As Long                                                      ' ע���ؼ�����������
    Dim tmpVal As String                                                        ' ע���ؼ���ֵ����ʱ�洢��
    Dim KeyValSize As Long                                                      ' ע���ؼ��Ա����ĳߴ�
    '------------------------------------------------------------
    ' �� {HKEY_LOCAL_MACHINE...} �µ� RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)                ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError                              ' �������...
    
    tmpVal = String$(1024, 0)                                                   ' ��������ռ�
    KeyValSize = 1024                                                           ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
    KeyValType, tmpVal, KeyValSize)                                             ' ���/�����ؼ���ֵ
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError                              ' �������
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then                               ' Win95 ��ӳ�����ս��ַ���...
        tmpVal = Left(tmpVal, KeyValSize - 1)                                   ' Null ���ҵ�,���ַ����з������
    Else                                                                        ' WinNT û�п��ս��ַ���...
        tmpVal = Left(tmpVal, KeyValSize)                                       ' Null û�б��ҵ�, �����ַ���
    End If
    '------------------------------------------------------------
    ' ����ת���Ĺؼ��ֵ�ֵ����...
    '------------------------------------------------------------
    Select Case KeyValType                                                      ' ������������...
    Case REG_SZ                                                                 ' �ַ���ע��ؼ�����������
        KeyVal = tmpVal                                                         ' �����ַ�����ֵ
    Case REG_DWORD                                                              ' ���ֽڵ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                                        ' ��ÿλ����ת��
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))                       ' ����ֵ�ַ��� By Char��
        Next
        KeyVal = Format$("&h" + KeyVal)                                         ' ת�����ֽڵ��ַ�Ϊ�ַ���
    End Select
    
    GetKeyValue = True                                                          ' ���سɹ�
    rc = RegCloseKey(hKey)                                                      ' �ر�ע���ؼ���
    Exit Function                                                               ' �˳�
    
GetKeyError:                                                                    ' �������������...
    KeyVal = ""                                                                 ' ���÷���ֵ�����ַ���
    GetKeyValue = False                                                         ' ����ʧ��
    rc = RegCloseKey(hKey)                                                      ' �ر�ע���ؼ���
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Canceled = False
    Scrolling = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    
    uFMOD_PlaySong 0, 0, 0                                                      'ֹͣ����xm����
End Sub

Private Sub sclVol_Change()
    uFMOD_SetVolume sclVol.Value
End Sub

Private Sub sclVol_Scroll()
    uFMOD_SetVolume sclVol.Value
End Sub

Private Sub Timer1_Timer()
    Dim txt1 As String
    txt1 = "ScreenSnap" & vbCrLf & vbCrLf & lblVersion.Caption & vbCrLf & vbCrLf & _
    "�ر���л" & vbCrLf & vbCrLf & "���������ߵ����ѣ�" & vbCrLf & vbCrLf & _
    "֧�ֱ���" & vbCrLf & vbCrLf & "BMP/PNG/JPG/GIF" & vbCrLf & vbCrLf & "��ʽͼƬ" & _
    vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & Chr(0)
    Timer1.Enabled = False
    Scrolling = True
    Scroll Picture1, txt1
End Sub

Private Sub Scroll(Pic As PictureBox, TxtScroll As String, Optional Alignment As Long = &H1)
    Dim TextLine() As String                                                    'Text lines array
    
    ' Dim Alignment As Long       'Text alignment
    Dim t As Long                                                               'Timer counter (frame delay)
    Dim Index As Long                                                           'Actual line index
    Dim RText As Rect                                                           'Rectangle into each new text line will be drawed
    Dim RClip As Rect                                                           'Rectangle to scroll up
    Dim RUpdate As Rect                                                         'Rectangle to update (not used)
    
    If TxtScroll = "" Then Exit Sub
    TextLine() = Split(TxtScroll, vbCrLf)
    
    With Pic
        .ScaleMode = vbPixels
        .AutoRedraw = True
        'Set rectangles
        SetRect RClip, 0, 1, .ScaleWidth, .ScaleHeight
        SetRect RText, 0, .ScaleHeight, .ScaleWidth, .ScaleHeight + .TextHeight("")
    End With
    
    Dim txt As String                                                           'Text to be drawed
    With Pic
        Do
            'Periodic frames
            If GetTickCount - t > 25 Then                                       'Set your delay here [ms]
                'Reset timer counter
                t = GetTickCount
                'Line ( + spacing ) totaly scrolled ?
                If RText.Bottom < .ScaleHeight Then
                    'Move down Text area out scroll area...
                    OffsetRect RText, 0, .TextHeight("")                        ' + space between lines [Pixels]
                    'Get new line
                    If Alignment = &H1 Then
                        'If alignment = Center, remove spaces
                        txt = Trim(TextLine(Index))
                    Else
                        'Case else, preserve them
                        txt = TextLine(Index)
                    End If
                    'Source line counter...
                    Index = Index + 1
                End If
                'Draw text
                DrawText .hDC, txt, lstrlen(txt), RText, Alignment
                'Move up one pixel Text area
                OffsetRect RText, 0, -1
                'Finaly, scroll up (1 pixel)...
                ScrollDC .hDC, 0, -1, RClip, RClip, 0, RUpdate
                
                '...and draw a bottom line to prevent... (well, don't draw it and see what happens)
                Pic.Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), .BackColor
                '(Refresh doesn't needed: any own PictureBox draw method calls Refresh method)
            End If
            DoEvents
        Loop Until Scrolling = False Or Index > UBound(TextLine)
    End With
    If m_View = False Then
        Unload Me
    Else
        If Scrolling Then Timer1_Timer
    End If
End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(Optional sView As Boolean = False) As Boolean
    m_View = sView
    cmdOK.Visible = sView
    Timer1.Enabled = True
    ' Display the form.
    Show 1
    ShowForm = Canceled
End Function
