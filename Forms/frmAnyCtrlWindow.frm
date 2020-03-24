VERSION 5.00
Begin VB.Form frmAnyWindowCtrl 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.Image imgGetPosition 
         Height          =   480
         Left            =   480
         Picture         =   "frmAnyCtrlWindow.frx":0000
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmAnyWindowCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HwndLng As Long

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE    '让窗口在顶层
End Sub

Private Sub imgGetPosition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgGetPosition.MousePointer = 99
    imgGetPosition.MouseIcon = imgGetPosition.Picture
    Set imgGetPosition.Picture = LoadPicture()
    Timer1.Enabled = True
End Sub

Private Sub imgGetPosition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err1
    Timer1.Enabled = False
    imgGetPosition.Picture = imgGetPosition.MouseIcon
    imgGetPosition.MousePointer = 0
    imgGetPosition.Enabled = False
    
    SnapSub 4
    
    imgGetPosition.Enabled = True
    Exit Sub
Err1:
    MsgBox "错误！imgGetPosition_MouseUp" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub Timer1_Timer()
    Dim a As POINTAPI
    GetCursorPos a
    HwndLng = WindowFromPoint(a.X, a.Y)
    Me.Caption = LoadResString(11800) & HwndLng
End Sub
