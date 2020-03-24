Attribute VB_Name = "modCreatPicsAfterTray"
Option Explicit

Public Sub CreatPicsAfterTraySub()
    Dim CPAT As Long
    
    For CPAT = (frmPicNum - SnapWhenTrayLng + 1) To frmPicNum
        frmMain.listSnapPic.AddItem PictureForms.Item(1 + CPAT).PictureName
        'frmMain.listSnapPic.Selected(CPAT) = True
    Next CPAT
    
    If frmPicNum - SnapWhenTrayLng = -1 Then
        frmMain.listSnapPic.Selected(0) = True
    Else
        frmMain.listSnapPic.Selected(frmPicNum - SnapWhenTrayLng) = True
    End If
    SnapWhenTrayLng = 0
    SnapWhenTrayBoo = False
End Sub
