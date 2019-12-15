Attribute VB_Name = "modChooseFileDialog"
Option Explicit

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OpenFileName) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As OpenFileName) As Long
Type OpenFileName
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Function GetDialog(ByVal sMethod As String, ByVal sTitle As String, ByVal sFilename As String, ByVal Frmhdc As Long, Optional ByVal Ways As Integer = 0) As String
    On Error GoTo myError
    Dim rtn As Long, pos As Integer
    Dim file As OpenFileName
    file.lStructSize = Len(file)
    file.hwndOwner = Frmhdc
    file.hInstance = App.hInstance
    file.lpstrFile = sFilename & String$(255 - Len(sFilename), 0)
    file.nMaxFile = 255
    file.lpstrFileTitle = String$(255, 0)
    file.nMaxFileTitle = 255
    file.lpstrInitialDir = ""
    If Ways = 0 Then
        file.lpstrFilter = "Bitmap文件(*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "JPEG文件(*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) _
        + "PNG文件(*.png)" + Chr$(0) + "*.png" + Chr$(0) + "GIF文件(*.gif)" + Chr$(0) + "*.gif" + Chr$(0) '这个为图片文件
    ElseIf Ways = 1 Then
        file.lpstrFilter = "所有支持的图片文件" + Chr$(0) + "*.bmp;*.jpg;*.png" + Chr$(0) + "Bitmap文件(*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "JPEG文件(*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) + "PNG文件(*.png)" + Chr$(0) + "*.png" '这个为图片文件
    ElseIf Ways = 2 Then
        file.lpstrFilter = "PNG文件(*.png)" + Chr$(0) + "*.png" + Chr$(0) + "Bitmap文件(*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "JPEG文件(*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) _
        + "GIF文件(*.gif)" + Chr$(0) + "*.gif" + Chr$(0)                        '这个为图片文件
    End If
    file.lpstrTitle = sTitle
    If UCase(sMethod) = "OPEN" Then
        file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
        rtn = GetOpenFileName(file)
    Else
        file.lpstrDefExt = "exe"
        file.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_OVERWRITEPROMPT
        rtn = GetSaveFileName(file)
    End If
    If rtn > 0 Then
        pos = InStr(file.lpstrFile, Chr$(0))
        If pos > 0 Then
            GetDialog = Left$(file.lpstrFile, pos - 1)
        End If
    End If
    Exit Function
myError:
    MsgBox "操作失败！", vbCritical + vbOKOnly
End Function
