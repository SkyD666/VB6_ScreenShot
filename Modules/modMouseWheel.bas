Attribute VB_Name = "modMouseWheel"
Option Explicit

Public frmName As Form

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Global lpPrevWndProcA As Long

Public bMouseFlag As Boolean                                                    '鼠标事件激活标志

Public Sub HookMouse(ByVal frm As Form)
    lpPrevWndProcA = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    Set frmName = frm
End Sub

Public Sub UnHookMouse(ByVal frm As Form)
    SetWindowLong frm.hwnd, GWL_WNDPROC, lpPrevWndProcA
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    Case WM_MOUSEWHEEL                                                          '滚动
        Dim wzDelta, wKeys As Integer
        'wzDelta传递滚轮滚动的快慢，该值小于零表示滚轮向后滚动（朝用户方向），
        '大于零表示滚轮向前滚动（朝显示器方向）
        wzDelta = HIWORD(wParam)
        'wKeys指出是否有CTRL=8、SHIFT=4、鼠标键(左=2、中=16、右=2、附加)按下，允许复合
        wKeys = LOWORD(wParam)
        '--------------------------------------------------
        Dim k As Integer
        '——————————————frmmain滚动条
        If wKeys = 4 Then                                                       '按住Shift键滚动鼠标滚轮实现左右移动
            k = (frmName.HScroll1.Value - Sgn(wzDelta) * frmName.HScroll1.LargeChange)
            If k > frmName.HScroll1.Max Then k = frmName.HScroll1.Max
            If k < frmName.HScroll1.Min Then k = frmName.HScroll1.Min
            frmName.HScroll1.Value = k
        ElseIf wKeys = 8 Then                                                   'CTRL+滚轮缩放图片
            If Sgn(wzDelta) = -1 And frmMain.cmbZoom.ListIndex > 0 Then
                frmMain.cmbZoom.ListIndex = frmMain.cmbZoom.ListIndex - 1
            ElseIf Sgn(wzDelta) = 1 And frmMain.cmbZoom.ListIndex < frmMain.cmbZoom.ListCount - 1 Then
                frmMain.cmbZoom.ListIndex = frmMain.cmbZoom.ListIndex + 1
            End If
        Else
            k = (frmName.VScroll1.Value - Sgn(wzDelta) * frmName.VScroll1.LargeChange)
            If k > frmName.VScroll1.Max Then k = frmName.VScroll1.Max
            If k < frmName.VScroll1.Min Then k = frmName.VScroll1.Min
            frmName.VScroll1.Value = k
        End If
        '--------------------------------------------------
    Case Else
        WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lParam)
    End Select
End Function

Private Function HIWORD(LongIn As Long) As Integer
    HIWORD = (LongIn And &HFFFF0000) \ &H10000                                  '取出32位值的高16位
End Function
Private Function LOWORD(LongIn As Long) As Integer
    LOWORD = LongIn And &HFFFF&                                                 '取出32位值的低16位
End Function
