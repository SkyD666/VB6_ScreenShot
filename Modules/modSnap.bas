Attribute VB_Name = "modSnap"
Option Explicit

Public Sub SnapSub(ByVal Index As Integer)                                      '��־��ͼ���� 0,1,2,3,4ȫ��������ڣ��ȼ�����꣬���ⴰ�ڿؼ�
    On Error GoTo ErrSnapSub
    Dim frmMainVis As Boolean, SelectedInt As Integer
    
    If Index = 1 Then
        If MsgBox(LoadResString(10804) & DelayTimeInt(1) & LoadResString(10805), vbInformation + vbOKCancel) <> vbOK Then Exit Sub '��ʾ�Ի���
    End If
    
    If HideWinCaptureInt(Index) = 1 Then                                        '��ͼʱ���ش���
        If frmMain.Visible Then
            frmMainVis = True
            frmMain.Visible = False
            Sleep 1000
        End If
    End If
    
    DelayTimeSub Index                                                          '��ʱ
    
    Dim Pic As Picture                                                          '��ʱ����ͼƬ����
    '�������������������������������������������������1  �ڴ�����ʾǰ��mnuNew_Click�����Ȼ�ȡ��������ʾ���ٸ�ֵ
    If (IncludeCursorBoo = True And Index <> 1) Or Index = 3 Then               '(Or IncludeCursorBoo = 3 ��ʱΪ�����꣬һ��Ҫ��ȡ���)
        Dim pci As CURSORINFO, iconinf As ICONINFO                              '�����ṹ
        pci.cbSize = Len(pci)                                                   '��ʼ
        GetCursorInfo pci
        GetIconInfo pci.hCursor, iconinf                                        'Ϊ�˻�ȡxHotspot
    End If
    '��������������������������������������������
    '===========================================================================
    Select Case Index                                                           'mnuNew_Click֮ǰһ��
    Case 0
        Set Pic = CaptureScreen()                                               '�Ƚ�ͼ���ٴ����ĵ���
    Case 1
        If ActiveWindowSnapMode = 0 Then                                        '��ͬ�Ļ���ڽ�ͼ���� �Ƚ�ͼ���ٴ����ĵ���
            Set Pic = CaptureActiveWindow()                                     '=ԭʼ����
        ElseIf ActiveWindowSnapMode = 1 Then
            Set Pic = CaptureActiveWindowB()                                    '=�·���
        End If
    Case 2
        If ActiveWindowSnapMode = 0 Then                                        '��ͬ�Ļ���ڽ�ͼ���� �Ƚ�ͼ���ٴ����ĵ���
            Set Pic = CaptureActiveWindow()                                     '=ԭʼ����
        ElseIf ActiveWindowSnapMode = 1 Then
            Set Pic = CaptureActiveWindowB()                                    '=�·���
        End If
    Case 3
        '        '���������������������������������ڴ�����ʾǰ��mnuNew_Click�����Ȼ�ȡ��������ʾ���ٸ�ֵ
        '        Dim pci As CURSORINFO, iconinf As ICONINFO                              '�����ṹ
        '        pci.cbSize = Len(pci)                                                   '��ʼ
        '        GetCursorInfo pci
        '        GetIconInfo pci.hCursor, iconinf                                        'Ϊ�˻�ȡiconinf��Ϣ
    Case 4
        Dim Rect1 As Rect
        Dim HwndLng As Long
        Dim a As POINTAPI
        GetCursorPos a
        HwndLng = WindowFromPoint(a.X, a.Y)
        GetWindowRect HwndLng, Rect1
        
        Set Pic = CaptureWindow(HwndLng, False, 0, 0, Rect1.Right3 - Rect1.Left, Rect1.Bottom - Rect1.Top) '�Ƚ�ͼ���ٴ����ĵ���
    End Select
    '=========================
    If frmMainVis Then frmMain.Visible = True                                   '������� ��ͼʱ���ش��� ��Ӧ
    frmMain.mnuNew_Click
    '=========================
    Select Case Index                                                           'mnuNew_Click֮��һ��
    Case 0
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    Case 1
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    Case 2
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    Case 3
        '����������������������������������ʾ��ֵ
        frmPictureCopy(frmPicNum).picScreenShot.Width = GetSystemMetrics(SM_CXCURSOR) * Screen.TwipsPerPixelX 'GetSystemMetrics  API
        frmPictureCopy(frmPicNum).picScreenShot.Height = GetSystemMetrics(SM_CYCURSOR) * Screen.TwipsPerPixelY
        DrawIcon frmPictureCopy(frmPicNum).picScreenShot.hDC, 0, 0, pci.hCursor
        DeleteObject iconinf.hbmColor
        DeleteObject iconinf.hbmMask
    Case 4
        Set frmMain.ActiveForm.picScreenShot.Picture = Pic
    End Select
    '===========================================================================
    
    frmPictureName(frmPicNum) = LoadResString(10705) & PicFilesCount & " *"
    frmPictureCopy(frmPicNum).Caption = frmPictureName(frmPicNum)
    
    '�������������������������������������������������2  ��ʾ��ֵ
    If (IncludeCursorBoo = True And Index <> 1) Or IncludeCursorBoo = 3 Then    '������ʱ��boolΪfalse   (Or IncludeCursorBoo = 3 ��ʱΪ�����꣬һ��Ҫ��ȡ���)
        DrawIcon frmPictureCopy(frmPicNum).picScreenShot.hDC, _
        pci.ptScreenPos.X - iconinf.xHotspot, pci.ptScreenPos.Y - iconinf.yHotspot, pci.hCursor ''��ȡ��λ���ȼ�ȥHotspot�õ�������Ͻ�����
        DeleteObject iconinf.hbmColor
        DeleteObject iconinf.hbmMask
    End If
    '��������������������������������������������
    Set frmPictureCopy(frmPicNum).picScreenShot.Picture = frmPictureCopy(frmPicNum).picScreenShot.Image
    Set PictureData(frmPicNum) = frmPictureCopy(frmPicNum).picScreenShot.Picture
    
    If Index = 2 Then If AutoSendToClipBoardBoo Then Clipboard.Clear: Clipboard.SetData PictureData(frmPicNum) '�ȼ���ͼ��ֱ�ӽ�ͼƬ���Ƶ�������
    
    'listbox�ӡ�*��
    frmMain.listSnapPic.AddItem frmPictureName(frmPicNum), frmMain.listSnapPic.ListIndex
    SelectedInt = frmMain.listSnapPic.ListIndex - 1
    frmMain.listSnapPic.RemoveItem frmMain.listSnapPic.ListIndex
    frmMain.listSnapPic.Selected(SelectedInt) = True
    
    If AutoSaveSnapInt(Index) = 1 Then AutoSaveSnapSub Index, frmPicNum
    
    TrayTip frmMain, App.Title & " - " & LoadResString(10809) & frmPicNum + 1 & LoadResString(10810) '��   �Ž�ͼ
    
    If SoundPlayBoo Then SoundPlay                                              '�Ƿ񲥷���ʾ��
    
    Exit Sub
ErrSnapSub:
    MsgBox "����" & vbCrLf & "������룺" & Err.Number & vbCrLf & "����������" & Err.Description, vbCritical + vbOKOnly
End Sub