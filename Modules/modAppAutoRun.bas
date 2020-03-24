Attribute VB_Name = "modAppAutoRun"
Option Explicit

Public Function AppAutoRun() As Integer                                         '读取是否开机启动  0为不启动，1为开机启动
    Dim Str As String
    Dim w As Object
    Set w = CreateObject("wscript.shell")
    On Error GoTo Err:
    Str = w.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\ScreenSnap")
    If Str = Chr(34) & App.path & "\" & "ScreenSnap" & ".exe" & Chr(34) & " AUTORUN" Then
        AppAutoRun = 1
    Else
        AppAutoRun = 0
    End If
    Exit Function
Err:
    AppAutoRun = 0
End Function
