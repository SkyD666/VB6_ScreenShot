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

Private Const BIF_RETURNONLYFSDIRS = &H1                                        '����ļ���
Private Const BIF_NEWDIALOGSTYLE = &H40                                         '����ʽ�����½��ļ��а�ť���ɵ����Ի����С��
Private Const BIF_NONEWFOLDERBUTTON = &H200                                     '����ʽ�У�û���½���ť��ֻ����С��

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
        .pidlRoot = 0&                                                          '��Ŀ¼��һ�㲻��Ҫ��
        .lpszTitle = Text
        .ulFlags = BIF_NEWDIALOGSTYLE                                           '������Ҫ����
    End With
    pidl = SHBrowseForFolder(bi)
    path = Space$(512)
    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        GetFolderName = Left(path, InStr(path, Chr(0)) - 1)
    End If
End Function

