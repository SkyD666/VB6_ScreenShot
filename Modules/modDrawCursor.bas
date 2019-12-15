Attribute VB_Name = "modDrawCursor"
Option Explicit
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long '��������� 75 ����ʶ����ͨ�����ò�ͬ�ı�ʶ���Ϳ��Ի�ȡϵͳ�ֱ��ʡ�������ʾ����Ŀ�Ⱥ͸߶ȡ��������Ŀ�Ⱥ͸߶�
Public Const SM_CXSCREEN = 0                                                    'X Size of screen
Public Const SM_CYSCREEN = 1                                                    'Y Size of Screen
Public Const SM_CXCURSOR = 13                                                   'Width of standard cursor
Public Const SM_CYCURSOR = 14                                                   'Height of standard cursor
'������������������������������������������������������ȡ���һ����
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As CURSORINFO) As Boolean
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type ICONINFO
    fIcon As Boolean
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Public Type CURSORINFO
    cbSize As Long
    ' Specifies the size, in bytes, of the structure.
    ' The caller must set this to Marshal.SizeOf(typeof(CURSORINFO)).
    flags As Long
    ' Specifies the cursor state. This parameter can be one of the following values:
    ' 0 The cursor is hidden.
    ' 1 The cursor is showing.
    hCursor As Long
    ' Handle to the cursor.
    ptScreenPos As POINTAPI
End Type
Public IncludeCursorBoo As Boolean
'����������������������������������������������������
