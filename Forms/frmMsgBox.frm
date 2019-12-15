VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5760
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAllNo 
      Caption         =   "全部否"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAllYes 
      Caption         =   "全部是"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "否(&N)"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "是(&Y)"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "有未保存的文件，是否保存？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   3120
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAllNo_Click()
    NewMsgBoxInt = 4
    Unload Me
End Sub

Private Sub cmdAllYes_Click()
    NewMsgBoxInt = 3
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    NewMsgBoxInt = -1
    If CloseFilesModeInt = 1 Then CloseFilesModeInt = 0
    Unload Me
End Sub

Private Sub cmdYes_Click()
    NewMsgBoxInt = 1
    Unload Me
End Sub

Private Sub cmdNo_Click()
    NewMsgBoxInt = 2
    Unload Me
End Sub

Private Sub Form_Load()
    MessageBeep MB_ICONEXCLAMATION                                              '发声音
    
    LoadLanguages "frmMsgBox"
    
    Me.Caption = App.EXEName
    Me.Icon = frmMain.Icon
    
    If CloseFilesModeInt = 0 Then                                               '显示三个按钮
        Me.Width = 5340
        cmdAllNo.Visible = False
        cmdAllYes.Visible = False
        cmdYes.Left = 360
        cmdYes.Width = 1335
        cmdNo.Left = 1920
        cmdNo.Width = 1335
        cmdCancel.Left = 3480
        cmdCancel.Width = 1335
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If NewMsgBoxInt = 0 Then
        NewMsgBoxInt = -1
        If CloseFilesModeInt = 1 Then CloseFilesModeInt = 0
    End If
End Sub
