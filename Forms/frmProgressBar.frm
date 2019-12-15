VERSION 5.00
Begin VB.Form frmProgressBar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "请稍候……"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4920
   StartUpPosition =   1  '所有者中心
   Begin VB.Shape shpProgressBar 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    LoadLanguages "frmProgressBar"
    
    shpProgressBar.Visible = True
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE    '让窗口在顶层
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
End Sub
