Attribute VB_Name = "modGDIScreenShot"
Option Explicit

'声明API函数
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'StretchBlt() Modes => SetStretchBltMode
Public Const BLACKONWHITE As Long = 1
Public Const WHITEONBLACK As Long = 2
Public Const COLORONCOLOR As Long = 3
Public Const HALFTONE As Long = 4
Public Const MAXSTRETCHBLTMODE As Long = 4

Public Const SRCCOP = &HCC0020
'自定义数据类型
Public Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Public Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Public Type Rect
    Left As Long
    Top As Long
    Right3 As Long
    Bottom As Long
End Type
Public Type PicBmp
    size As Long
Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
'定义常量
Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

'=================================
Public Function CaptureActiveWindowB() As Picture                               '另一种活动窗口截图方法，先全屏截图，再抠图
    Dim hWndActive As Long
    Dim RectActive As Rect
    Set frmMain.Picture1.Picture = LoadPicture()
    
    hWndActive = GetForegroundWindow()
    GetWindowRect hWndActive, RectActive                                        '取得窗口rect
    
    RectActive.Bottom = RectActive.Bottom * Screen.TwipsPerPixelY               '由像素变成twip
    RectActive.Left = RectActive.Left * Screen.TwipsPerPixelX
    RectActive.Right3 = RectActive.Right3 * Screen.TwipsPerPixelX
    RectActive.Top = RectActive.Top * Screen.TwipsPerPixelY
    
    If RectActive.Right3 < 0 Or RectActive.Left > (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX) Or _
        RectActive.Top > (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY) Or RectActive.Bottom < 0 Then
        '只需要4个其中之一不在屏幕范围内，就说明活动窗体全都不在可见范围内
        Set CaptureActiveWindowB = CaptureActiveWindow
        Exit Function                                                           '*********************************
    End If
    
    frmMain.Picture1.Width = (RectActive.Right3 - RectActive.Left)              '需要一个picturebox做载体
    frmMain.Picture1.Height = (RectActive.Bottom - RectActive.Top)
    frmMain.Picture1.PaintPicture CaptureScreen, 0, 0, (RectActive.Right3 - RectActive.Left), (RectActive.Bottom - RectActive.Top), RectActive.Left, RectActive.Top, _
    (RectActive.Right3 - RectActive.Left), (RectActive.Bottom - RectActive.Top)
    '——————————————————————画鼠标
    If IncludeCursorBoo = True Then
        Dim pci As CURSORINFO, iconinf As ICONINFO                              '两个结构
        pci.cbSize = Len(pci)                                                   '初始
        GetCursorInfo pci
        GetIconInfo pci.hCursor, iconinf                                        '为了获取xHotspot
        DrawIcon frmMain.Picture1.hDC, pci.ptScreenPos.X - iconinf.xHotspot - (RectActive.Left / Screen.TwipsPerPixelY), _
        pci.ptScreenPos.Y - iconinf.yHotspot - (RectActive.Top / Screen.TwipsPerPixelY), pci.hCursor '获取的位置先减去Hotspot得到鼠标左上角坐标，再减去活动窗口左上角
    End If
    '——————————————————————
    Set frmMain.Picture1.Picture = frmMain.Picture1.Image
    Set CaptureActiveWindowB = frmMain.Picture1.Picture
End Function

Public Function CaptureActiveWindow() As Picture                                '第一种活动窗口截图方法，直接截图，在Aero下会有bug
    Dim hWndActive As Long
    Dim R As Long
    Dim RectActive As Rect
    hWndActive = GetForegroundWindow()
    R = GetWindowRect(hWndActive, RectActive)
    Set frmMain.Picture1.Picture = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right3 - RectActive.Left, RectActive.Bottom - RectActive.Top)
    '——————————————————————画鼠标
    If IncludeCursorBoo = True Then
        Dim pci As CURSORINFO, iconinf As ICONINFO                              '两个结构
        pci.cbSize = Len(pci)                                                   '初始
        GetCursorInfo pci
        GetIconInfo pci.hCursor, iconinf                                        '为了获取xHotspot
        DrawIcon frmMain.Picture1.hDC, pci.ptScreenPos.X - iconinf.xHotspot - (RectActive.Left / Screen.TwipsPerPixelY), _
        pci.ptScreenPos.Y - iconinf.yHotspot - (RectActive.Top / Screen.TwipsPerPixelY), pci.hCursor '获取的位置先减去Hotspot得到鼠标左上角坐标，再减去活动窗口左上角
    End If
    '——————————————————————
    Set CaptureActiveWindow = frmMain.Picture1.Image
End Function

'抓取窗口任意部分的函数
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    On Error GoTo Err7
    
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim R As Long
    Dim hdcSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long                                                  '定义各种变量
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
    If Client Then
        hdcSrc = GetDC(hWndSrc)                                                 '从客户区获得设备信息
    Else
        hdcSrc = GetWindowDC(hWndSrc)
    End If
    hDCMemory = CreateCompatibleDC(hdcSrc)
    hBmp = CreateCompatibleBitmap(hdcSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    '获得屏幕属性
    RasterCapsScrn = GetDeviceCaps(hdcSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hdcSrc, SIZEPALETTE)                        '屏幕大小
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hdcSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hdcSrc, LeftSrc, TopSrc, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    R = DeleteDC(hDCMemory)
    R = ReleaseDC(hWndSrc, hdcSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    
    Exit Function
Err7:
    MsgBox "错误！CaptureWindow" & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "错误描述：" & Err.Description, vbCritical + vbOKOnly
End Function

'抓取整个屏幕的函数
Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long, ScrWidth As Long, ScrHeight As Long
    ScrWidth = GetSystemMetrics(SM_CXSCREEN)
    ScrHeight = GetSystemMetrics(SM_CYSCREEN)
    '设置桌面窗体对象
    hWndScreen = GetDesktopWindow
    '不要用Screen.Width  ，Screen.Height，在win8/10系统上有bug
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, ScrWidth, ScrHeight)
End Function

'创建一个".bitmap"类型的Picture对象
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long
    Dim Pic As PicBmp
    ' IPicture要求引用"Standard OLE Types."
    Dim IPic As IPicture
    Dim IID_IDispatch As Guid
    With IID_IDispatch                                                          'With语句
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With Pic
        .size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With
    '创建一个图像对象
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    '返回新的图像对象
    Set CreateBitmapPicture = IPic
End Function
