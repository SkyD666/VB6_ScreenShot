Attribute VB_Name = "modChooseFolderDialog"
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1                                        '浏览文件夹
Private Const BIF_NEWDIALOGSTYLE = &H40                                         '新样式（有新建文件夹按钮，可调整对话框大小）
Private Const BIF_NONEWFOLDERBUTTON = &H200                                     '新样式中，没有新建按钮（只调大小）

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, _
    ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long

Public AutoSavePicFolderPath As String

Public Function GetFolderName(hwnd As Long, Text As String) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim path As String
    With bi
        .hOwner = hwnd
        .pidlRoot = 0&                                                          '根目录，一般不需要改
        .lpszTitle = Text
        .ulFlags = BIF_NEWDIALOGSTYLE                                           '根据需要调整
    End With
    pidl = SHBrowseForFolder(bi)
    path = Space$(512)
    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        GetFolderName = Left(path, InStr(path, Chr(0)) - 1)
    End If
End Function

