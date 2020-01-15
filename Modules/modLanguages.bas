Attribute VB_Name = "modLanguages"
Option Explicit
Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer     '≈–∂œ”Ô—‘

Public Sub LoadLanguages(ByVal IDStr As String)
    Select Case IDStr
    Case "frmSettings"
        With frmSettings
            .Caption = LoadResString(11200)
            
            .chkSoundPlaySetFrm.Caption = LoadResString(11300)
            .chkAutoRun.Caption = LoadResString(11301)
            .chkHideWinValue.Caption = LoadResString(11302)
            .chkAutoSaveSnapValue.Caption = LoadResString(11310)
            .chkEndOrMin.Caption = LoadResString(11304)
            .optActiveWinSnapMode1.Caption = LoadResString(11305)
            .optActiveWinSnapMode0.Caption = LoadResString(11306)
            .labDelayTime.Caption = LoadResString(11307)
            .Label1.Caption = LoadResString(11308)
            .chkAutoSendToClipBoard.Caption = LoadResString(11309)
            .chkIncludeCursor.Caption = LoadResString(11312)
            
            .HotKeyNow_lab.Caption = LoadResString(11400)
            .txtHotKeyScreenShot.Text = LoadResString(11401)
            .Label7.Caption = LoadResString(11404)
            .optDeclareHotKeyWay1.Caption = LoadResString(11402)
            .optDeclareHotKeyWay2.Caption = LoadResString(11403)
            
            .Label2.Caption = LoadResString(11500)
            .Label4.Caption = LoadResString(11501)
            .Label3.Caption = LoadResString(11502)
            .cmdOpenTheFolder.Caption = LoadResString(11503)
            
            .Label10.Caption = LoadResString(11600)
        End With
        
    Case "frmMsgBox"
        With frmMsgBox
            .Label1.Caption = LoadResString(10900)
            .cmdYes.Caption = LoadResString(10901)
            .cmdNo.Caption = LoadResString(10902)
            .cmdAllYes.Caption = LoadResString(10903)
            .cmdAllNo.Caption = LoadResString(10904)
            .cmdCancel.Caption = LoadResString(10905)
        End With
        
    Case "frmProgressBar"
        With frmProgressBar
            .Caption = LoadResString(11100)
        End With
        
    Case "frmAbout"
        With frmAbout
            .cmdOK.Caption = LoadResString(11000)
            .cmdSysInfo.Caption = LoadResString(11001)
        End With
        
    Case "frmMain"
        With frmMain
            .mnuFile.Caption = LoadResString(10000)
            .mnuNew.Caption = LoadResString(10001)
            .mnuSave.Caption = LoadResString(10002)
            .mnuCloseAllFilesUnsaved.Caption = LoadResString(10003)
            .mnuClose.Caption = LoadResString(10004)
            .mnuExit.Caption = LoadResString(10005)
            .mnuOpenTheFolder.Caption = LoadResString(10006)
            
            .mnuEdit.Caption = LoadResString(10100)
            .mnuCopy.Caption = LoadResString(10101)
            .mnuPaste.Caption = LoadResString(10102)
            
            .mnuCapture.Caption = LoadResString(10200)
            .mnuScreenSnap.Caption = LoadResString(10201)
            .mnuActiveWinSnap.Caption = LoadResString(10202)
            .mnuCursorSnap.Caption = LoadResString(10203)
            .mnuAnyWindowCtrlSnap.Caption = LoadResString(10204)
            
            .mnuTool.Caption = LoadResString(10300)
            .mnuSetting.Caption = LoadResString(10301)
            
            .mnuView.Caption = LoadResString(10403)
            .mnuZoom.Caption = LoadResString(10404)
            .mnuZoomIn.Caption = LoadResString(10405)
            .mnuZoomOut.Caption = LoadResString(10406)
            
            .mnuHelp.Caption = LoadResString(10400)
            .mnuAbout.Caption = LoadResString(10401)
            .mnuSourceCode.Caption = LoadResString(10402)
            
            .mnufrmPicCopy.Caption = LoadResString(10500)
            .mnufrmPicPaste.Caption = LoadResString(10501)
            .mnufrmPicClose.Caption = LoadResString(10502)
            
            .mnuTrayShow.Caption = LoadResString(10600)
            .mnuTrayScreenSnap.Caption = LoadResString(10601)
            .mnuTrayActiveWinSnap.Caption = LoadResString(10602)
            .mnuTrayCursorSnap.Caption = LoadResString(10606)
            .mnuTrayWinCtrlSnap.Caption = LoadResString(10606)
            .mnuTraySetting.Caption = LoadResString(10603)
            .mnuTrayAbout.Caption = LoadResString(10604)
            .mnuTrayExit.Caption = LoadResString(10605)
            
            .chkSoundPlay.Caption = LoadResString(10700)
            .labPicQuantity.Caption = LoadResString(10701)
            .labMousePos.Caption = LoadResString(10702)
            
            .imgScreenSnap.ToolTipText = LoadResString(10802)
            .imgActiveWin.ToolTipText = LoadResString(10803)
            .imgCursor.ToolTipText = LoadResString(10812)
            .imgAnyCtrlWindow.ToolTipText = LoadResString(10813)
        End With
    End Select
End Sub
