Attribute VB_Name = "modSnap"
Option Explicit

Public Sub SnapSub(ByVal Index As Integer)                                      '标志截图类型 0,1,2,3,4全屏，活动窗口，热键，光标，任意窗口控件
    'On Error GoTo ErrSnapSub
    Dim frmMainVis As Boolean
    
    If Index = 1 Then
        If MsgBox(LoadResString(10804) & DelayTimeInt(1) & LoadResString(10805), vbInformation + vbOKCancel) <> vbOK Then Exit Sub '提示对话框
    End If
    
    If HideWinCaptureInt(Index) = 1 Then                                        '截图时隐藏窗口
        If frmMain.Visible Then
            frmMainVis = True
            frmMain.Visible = False
            Sleep 1000
        End If
    End If
    
    DelayTimeSub Index                                                          '延时
    
    Dim Pic As Picture                                                          '临时保存图片变量
    '——————————————————————画鼠标1  在窗体显示前（mnuNew_Click），先获取参数，显示后再赋值
    Dim pci As CURSORINFO, iconinf As ICONINFO                                  '两个结构
    If (IncludeCursorBoo = True And Index <> 1 And Index <> 2) Or Index = 3 Then '(Or IncludeCursorBoo = 3 此时为捕获光标，一定要截取光标)
        pci.cbSize = Len(pci)                                                   '初始
        GetCursorInfo pci
        GetIconInfo pci.hCursor, iconinf                                        '为了获取xHotspot
    End If
    '——————————————————————
    '===========================================================================
    Select Case Index                                                           'mnuNew_Click之前一次
    Case 0
        Set Pic = CaptureScreen()                                               '先截图，再创建文档，
    Case 1
        If ActiveWindowSnapMode = 0 Then                                        '不同的活动窗口截图方法 先截图，再创建文档，
            Set Pic = CaptureActiveWindow()                                     '=原始方法
        ElseIf ActiveWindowSnapMode = 1 Then
            Set Pic = CaptureActiveWindowB()                                    '=新方法
        End If
    Case 2
        If ActiveWindowSnapMode = 0 Then                                        '不同的活动窗口截图方法 先截图，再创建文档，
            Set Pic = CaptureActiveWindow()                                     '=原始方法
        ElseIf ActiveWindowSnapMode = 1 Then
            Set Pic = CaptureActiveWindowB()                                    '=新方法
        End If
    Case 3
        '        '————————————————在窗体显示前（mnuNew_Click），先获取参数，显示后再赋值
        '        Dim pci As CURSORINFO, iconinf As ICONINFO                              '两个结构
        '        pci.cbSize = Len(pci)                                                   '初始
        '        GetCursorInfo pci
        '        GetIconInfo pci.hCursor, iconinf                                        '为了获取iconinf信息
    Case 4
        Dim Rect1 As Rect
        Dim HwndLng As Long
        Dim a As POINTAPI
        GetCursorPos a
        HwndLng = WindowFromPoint(a.X, a.Y)
        GetWindowRect HwndLng, Rect1
        
        Set Pic = CaptureWindow(HwndLng, False, 0, 0, Rect1.Right3 - Rect1.Left, Rect1.Bottom - Rect1.Top) '先截图，再创建文档，
    End Select
    '=========================
    If frmMainVis Then frmMain.Visible = True                                   '与上面的 截图时隐藏窗口 对应
    frmMain.mnuNew_Click
    '=========================
    Select Case Index                                                           'mnuNew_Click之后一次
    Case 0
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    Case 1
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    Case 2
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    Case 3
        '————————————————显示后赋值
        frmMain.ActiveForm.picScreenShot.Width = GetSystemMetrics(SM_CXCURSOR) * Screen.TwipsPerPixelX 'GetSystemMetrics  API
        frmMain.ActiveForm.picScreenShot.Height = GetSystemMetrics(SM_CYCURSOR) * Screen.TwipsPerPixelY
        DrawIcon frmMain.ActiveForm.picScreenShot.hDC, 0, 0, pci.hCursor
        DeleteObject iconinf.hbmColor
        DeleteObject iconinf.hbmMask
    Case 4
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    End Select
    '===========================================================================
    
    frmMain.ActiveForm.PictureName = LoadResString(10705) & PicFilesCount & " *"
    frmMain.ActiveForm.Caption = frmMain.ActiveForm.PictureName
    
    '——————————————————————画鼠标2  显示后赋值
    If (IncludeCursorBoo = True And Index <> 1 And Index <> 2) Or IncludeCursorBoo = 3 Then '捕获光标时此bool为false   (Or IncludeCursorBoo = 3 此时为捕获光标，一定要截取光标)
        DrawIcon frmMain.ActiveForm.picScreenShot.hDC, _
        pci.ptScreenPos.X - iconinf.xHotspot, pci.ptScreenPos.Y - iconinf.yHotspot, pci.hCursor ''获取的位置先减去Hotspot得到鼠标左上角坐标
        DeleteObject iconinf.hbmColor
        DeleteObject iconinf.hbmMask
    End If
    '——————————————————————
    Set frmMain.ActiveForm.picScreenShot.Picture = frmMain.ActiveForm.picScreenShot.Image
    Set frmMain.ActiveForm.PictureData = frmMain.ActiveForm.picScreenShot.Picture
    
    If Index = 2 Then If AutoSendToClipBoardBoo Then Clipboard.Clear: Clipboard.SetData frmMain.ActiveForm.PictureData '热键截图后直接将图片复制到剪贴板
    
    'listbox加“*”
    frmMain.listSnapPic.List(frmMain.listSnapPic.ListIndex) = frmMain.ActiveForm.PictureName
    
    If AutoSaveSnapInt(Index) = 1 Then AutoSaveSnapSub Index, frmPicNum
    
    TrayTip frmMain, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '共   张截图
    
    If SoundPlayBoo Then SoundPlay                                              '是否播放提示音
    
    Exit Sub
ErrSnapSub:
    MsgBox "错误！Snap" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, vbCritical + vbOKOnly
End Sub
