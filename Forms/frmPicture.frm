VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10185
   Icon            =   "frmPicture.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   10186
   Begin VB.CommandButton cmdTransferVHScroll 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   8175
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox picScreenShot 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3000
         Left            =   1680
         ScaleHeight     =   3000
         ScaleWidth      =   4920
         TabIndex        =   4
         Top             =   1200
         Width           =   4920
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   5940
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3641
      Left            =   9240
      TabIndex        =   0
      Top             =   0
      Width           =   253
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Moving As Boolean, X1 As Long, Y1 As Long
Private horzMax As Long
Private vertMax As Long
Public FormNum As Integer                                                       '代替原来的label
Public PictureData As Picture                                                   '存储图片
Public PictureSaved As Boolean                                                  '图片是否保存
Public PictureName As String                                                    '每个窗体内图片名称
Public PicZoom As Integer                                                       '缩放

Private Sub cmdTransferVHScroll_Click()
    VHScroll
    
    With picScreenShot                                                          '小图片时居中
        If .Width < Picture1.ScaleWidth Or .Height < Picture1.ScaleHeight Then
            .Top = (Picture1.ScaleHeight - .Height) / 2
            .Left = (Picture1.ScaleWidth - .Width) / 2
        Else
            .Top = 0
            .Left = 0
        End If
    End With
End Sub

Private Sub Form_Activate()
    frmMain.labPicQuantity.Caption = Replace(LoadResString(10703), "0", CStr(frmPicNum + 1), , 1)
    frmMain.labPicQuantity.Caption = frmMain.labPicQuantity.Caption & Replace(LoadResString(10704), "0", FormNum + 1, , 1)
    frmMain.cmbZoom.Text = PicZoom & "%"
    If (frmMain.listSnapPic.ListIndex <> FormNum And SnapWhenTrayBoo = False) Then frmMain.listSnapPic.Selected(FormNum) = True
    
    HookMouse Me                                                                '鼠标滚轮HOOK
End Sub

Private Sub Form_Deactivate()
    UnHookMouse Me                                                              '卸载滚轮钩子
End Sub

Private Sub Form_Load()
    PictureSaved = True
    '-----------------------滚动条
    With HScroll1
        .Top = Picture1.Top + Picture1.Height
        .Left = Picture1.Left
        .Width = Picture1.Width
        .Min = 0
        .Max = 100
        .SmallChange = 4
        .LargeChange = 10
    End With
    
    With VScroll1
        .Top = Picture1.Top
        .Left = Picture1.Left + Picture1.Width
        .Height = Picture1.Height
        .Min = 0
        .Max = 100
        .SmallChange = 4
        .LargeChange = 10
    End With
    
    VHScroll
    '---------------------------
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With Picture1
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth - VScroll1.Width
        .Height = Me.ScaleHeight - HScroll1.Height
    End With
    
    With HScroll1
        .Top = Picture1.Top + Picture1.Height
        .Left = Picture1.Left
        .Width = Picture1.Width
    End With
    
    With VScroll1
        .Top = Picture1.Top
        .Left = Picture1.Left + Picture1.Width
        .Height = Picture1.Height
    End With
    
    VHScroll                                                                    '在“With picScreenShot”之前，否则滚动条一滑动，位置就变了，如不居中了
    
    With picScreenShot                                                          '小图片时居中
        If .Width < Picture1.ScaleWidth Or .Height < Picture1.ScaleHeight Then
            .Top = (Picture1.ScaleHeight - .Height) / 2
            .Left = (Picture1.ScaleWidth - .Width) / 2
        Else
            .Top = 0
            .Left = 0
        End If
    End With
End Sub

Private Sub VHScroll()
    horzMax = picScreenShot.ScaleWidth - Picture1.ScaleWidth
    
    With HScroll1
        .Value = 0
        If horzMax < 0 Then
            .Max = 0
            '.Visible = False                                                    ' Optional
        Else
            .Max = 100
            .Visible = True                                                     ' Optional
        End If
    End With
    
    vertMax = picScreenShot.ScaleHeight - Picture1.ScaleHeight
    With VScroll1
        .Value = 0
        If vertMax < 0 Then
            .Max = 0
            '.Visible = False                                                    ' Optional
        Else
            .Max = 100
            .Visible = True                                                     ' Optional
        End If
    End With
End Sub

Private Sub HScroll1_Change()
    If HScroll1.Max > 0 Then picScreenShot.Left = -(HScroll1.Value / HScroll1.Max) * horzMax
End Sub

Private Sub HScroll1_Scroll()
    If HScroll1.Max > 0 Then picScreenShot.Left = -(HScroll1.Value / HScroll1.Max) * horzMax
End Sub

Private Sub picScreenShot_Change()
    VHScroll
    
    With picScreenShot                                                          '调整图片居中或顶格
        If .Width < Picture1.ScaleWidth Or .Height < Picture1.ScaleHeight Then
            .Top = (Picture1.ScaleHeight - .Height) / 2
            .Left = (Picture1.ScaleWidth - .Width) / 2
        Else
            .Top = 0
            .Left = 0
        End If
    End With
    
    PictureSaved = False
End Sub

Private Sub VScroll1_Change()
    If VScroll1.Max > 0 Then picScreenShot.Top = -(VScroll1.Value / VScroll1.Max) * vertMax
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.Max > 0 Then picScreenShot.Top = -(VScroll1.Value / VScroll1.Max) * vertMax
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PictureSaved = False Then
        If CloseAllFilesUnsavedBoo = True Then
            UnloadfrmPic frmMain.listSnapPic.ListIndex
            frmMain.labMousePos.Caption = LoadResString(10702)                  '鼠标位置: X:0  Y:0
            UnHookMouse Me                                                      '卸载滚轮钩子
            frmMain.labPicQuantity.Caption = Replace(LoadResString(10701), "0", CStr(frmPicNum + 1), , 1)
            Exit Sub
        End If
        frmMsgBox.Show 1
        If NewMsgBoxInt = 1 Then
            SavePictures GetPicturePath(frmMain, frmMain.listSnapPic.ListIndex), frmMain.listSnapPic.ListIndex
            UnloadfrmPic frmMain.listSnapPic.ListIndex
            frmMain.labMousePos.Caption = LoadResString(10702)                  '鼠标位置: X:0  Y:0
            UnHookMouse Me                                                      '卸载滚轮钩子
            frmMain.labPicQuantity.Caption = Replace(LoadResString(10701), "0", CStr(frmPicNum + 1), , 1)
        ElseIf NewMsgBoxInt = 4 Then
            frmMain.CloseAllFilesUnsaved                                        '这里先不要NewMsgBoxInt = 0复原状态，标记是从退出中“全部否”的，给MDIForm_QueryUnload中的循环一个信号，防止在这里把子窗体全关闭了后在MDIForm_QueryUnload的循环中出错
        ElseIf NewMsgBoxInt = 3 Then                                            '保存到文件
            SavePictures GetDialog("save", LoadResString(10806), "Snap", frmMain.hwnd), 0, 1 '此处“0”随便改，因为批量保存不会用到此参数
        ElseIf NewMsgBoxInt = 2 Then
            UnloadfrmPic frmMain.listSnapPic.ListIndex
            frmMain.labMousePos.Caption = LoadResString(10702)                  '鼠标位置: X:0  Y:0
            UnHookMouse Me                                                      '卸载滚轮钩子
            frmMain.labPicQuantity.Caption = Replace(LoadResString(10701), "0", CStr(frmPicNum + 1), , 1)
        ElseIf NewMsgBoxInt = -1 Then
            Cancel = 1
            If EndOrMinBoo Then EndOrMinBoo = False
            Exit Sub                                                            '这里先不要NewMsgBoxInt = 0复原状态，标记是从退出中“取消”的，给MDIForm_QueryUnload中的循环一个信号，停止关闭过程，在那里会设置为0
        End If
    Else
        UnloadfrmPic frmMain.listSnapPic.ListIndex
        frmMain.labMousePos.Caption = LoadResString(10702)                      '鼠标位置: X:0  Y:0
        UnHookMouse Me                                                          '卸载滚轮钩子
    End If
End Sub

Private Sub picScreenShot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu frmMain.mnufrmPictureRight
    Else
        '鼠标在按钮上点击下去时执行(此时鼠标一直按着没提起来)
        Moving = True
        X1 = X
        Y1 = Y
    End If
End Sub

Private Sub picScreenShot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Moving Then
        picScreenShot.Left = picScreenShot.Left + X - X1
        picScreenShot.Top = picScreenShot.Top + Y - Y1
    End If
    
    frmMain.labMousePos.Caption = Replace(LoadResString(10702), "X:0", "X:" & X / Screen.TwipsPerPixelX)
    frmMain.labMousePos.Caption = Replace(frmMain.labMousePos.Caption, "Y:0", "Y:" & Y / Screen.TwipsPerPixelY)
End Sub

Private Sub picScreenShot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '鼠标在点击的情况下,放手时产生的事件
    Moving = False
End Sub

Private Sub Picture1_Click()
    With picScreenShot                                                          '小图片时居中
        If .Width < Picture1.ScaleWidth Or .Height < Picture1.ScaleHeight Then
            .Top = (Picture1.ScaleHeight - .Height) / 2
            .Left = (Picture1.ScaleWidth - .Width) / 2
        Else
            .Top = 0
            .Left = 0
        End If
    End With
End Sub
