Attribute VB_Name = "modMouseWheel"
Option Explicit

Public frmName As Form

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Global lpPrevWndProcA As Long

Public bMouseFlag As Boolean                                                    '����¼������־

Public Sub HookMouse(ByVal frm As Form)
    lpPrevWndProcA = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    Set frmName = frm
End Sub

Public Sub UnHookMouse(ByVal frm As Form)
    SetWindowLong frm.hwnd, GWL_WNDPROC, lpPrevWndProcA
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    Case WM_MOUSEWHEEL                                                          '����
        Dim wzDelta, wKeys As Integer
        'wzDelta���ݹ��ֹ����Ŀ�������ֵС�����ʾ���������������û����򣩣�
        '�������ʾ������ǰ����������ʾ������
        wzDelta = HIWORD(wParam)
        'wKeysָ���Ƿ���CTRL=8��SHIFT=4������(��=2����=16����=2������)���£�������
        wKeys = LOWORD(wParam)
        '--------------------------------------------------
        Dim MW_FB As Integer
        MW_FB = (wParam And &HFFFF0000) \ &H10000
        Dim k As Integer
        '        '����������������������������frmeffect������
        '        If DiffereStr = "frmEffectVScroll" Then
        '            If wKeys = 4 Then                                                   '��סShift������������ʵ�������ƶ�
        '                k = (frmEffect.HScroll1.Value - Sgn(MW_FB) * frmEffect.HScroll1.LargeChange)
        '                If k > frmEffect.HScroll1.Max Then k = frmEffect.HScroll1.Max
        '                If k < frmEffect.HScroll1.Min Then k = frmEffect.HScroll1.Min
        '                frmEffect.HScroll1.Value = k
        '            Else
        '                k = (frmEffect.VScroll1.Value - Sgn(MW_FB) * frmEffect.VScroll1.LargeChange)
        '                If k > frmEffect.VScroll1.Max Then k = frmEffect.VScroll1.Max
        '                If k < frmEffect.VScroll1.Min Then k = frmEffect.VScroll1.Min
        '                frmEffect.VScroll1.Value = k
        '            End If
        '        End If
        '����������������������������frmmain������
        If wKeys = 4 Then                                                       '��סShift������������ʵ�������ƶ�
            k = (frmName.HScroll1.Value - Sgn(MW_FB) * frmName.HScroll1.LargeChange)
            If k > frmName.HScroll1.Max Then k = frmName.HScroll1.Max
            If k < frmName.HScroll1.Min Then k = frmName.HScroll1.Min
            frmName.HScroll1.Value = k
        Else
            k = (frmName.VScroll1.Value - Sgn(MW_FB) * frmName.VScroll1.LargeChange)
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
    HIWORD = (LongIn And &HFFFF0000) \ &H10000                                  'ȡ��32λֵ�ĸ�16λ
End Function
Private Function LOWORD(LongIn As Long) As Integer
    LOWORD = LongIn And &HFFFF&                                                 'ȡ��32λֵ�ĵ�16λ
End Function
