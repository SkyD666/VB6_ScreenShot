Attribute VB_Name = "modGDISavePic"
Option Explicit

Public Const UnitPixel                  As Long = 2
Public Const EncoderQuality             As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Public Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Public Type EncoderParameter
    Guid(0 To 3)        As Long
    NumberOfValues      As Long
Type                As EncoderParameterValueType
    Value               As Long
End Type

Public Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

Public Type ImageCodecInfo
    ClassID(0 To 3)     As Long
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As Long
Public Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, size As Long) As Long
Public Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal size As Long, Encoders As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pclsid As Any) As Long

Public Enum ImageFileFormat
    bmp = 1
    Jpg = 2
    png = 3
    Gif = 4
End Enum

Public SetJpgQuality As Integer
Public AutoSavePicFormatStr As String

Public Function SaveStdPicToFile(StdPic As StdPicture, ByVal filename As String, _
    ByVal FileFormat As String, _
    Optional ByVal JpgQuality As Long = 80, Optional ByVal StdPicLongBoo As Boolean = False, _
    Optional ByVal StdPicLongLng As Long = 0) As Boolean
    
    JpgQuality = SetJpgQuality
    
    Dim CLSID(3)        As Long
    Dim Bitmap          As Long
    Dim Token           As Long
    Dim Gsp             As GdiplusStartupInput
    
    Gsp.GdiplusVersion = 1                                                      'GDI+ 1.0版本
    GdiplusStartup Token, Gsp                                                   '初始化GDI+
    If StdPicLongBoo = False Then
        GdipCreateBitmapFromHBITMAP StdPic.Handle, StdPic.hPal, Bitmap
    Else
        Bitmap = StdPicLongLng
    End If
    If Bitmap <> 0 Then                                                         '说明我们成功的将StdPic对象转换为GDI+的Bitmap对象了
        Select Case FileFormat
        Case "bmp"
            If Not GetEncoderClsID("Image/bmp", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), ByVal 0) = 0)
            End If
        Case "jpg"                                                              'JPG格式可以设置保存的质量
            Dim aEncParams()        As Byte
            Dim uEncParams          As EncoderParameters
            If GetEncoderClsID("Image/jpeg", CLSID) <> -1 Then
                uEncParams.Count = 1                                            ' 设置自定义的编码参数，这里为1个参数
                If JpgQuality < 0 Then
                    JpgQuality = 0
                ElseIf JpgQuality > 100 Then
                    JpgQuality = 100
                End If
                ReDim aEncParams(1 To Len(uEncParams))
                With uEncParams.Parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong                       ' 设置参数值的数据类型为长整型
                    Call CLSIDFromString(StrPtr(EncoderQuality), .Guid(0))      ' 设置参数唯一标志的GUID，这里为编码品质
                    .Value = VarPtr(JpgQuality)                                 ' 设置参数的值：品质等级，最高为100，图像文件大小与品质成正比
                End With
                CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), aEncParams(1)) = 0)
            End If
        Case "png"
            If Not GetEncoderClsID("Image/png", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), ByVal 0) = 0)
            End If
        Case "gif"
            If Not GetEncoderClsID("Image/gif", CLSID) = -1 Then                '如果原始的图像是24位，则这个函数会调用系统的调色板来将图像转换为8位，转换的效果会不尽人意,但也有可能系统不自动转换，保存失败
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), ByVal 0) = 0)
            End If
        End Select
    End If
    GdipDisposeImage Bitmap                                                     '注意释放资源
    GdiplusShutdown Token                                                       '关闭GDI+。
End Function

Public Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim size        As Long
    Dim i           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, size                                          '得到解码器数组的大小
    If size <> 0 Then
        ReDim Info(1 To Num) As ImageCodecInfo                                  '给数组动态分配内存
        ReDim Buffer(1 To size) As Byte
        GdipGetImageEncoders Num, size, Buffer(1)                               '得到数组和字符数据
        CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)                     '复制类头
        For i = 1 To Num                                                        '循环检测所有解码
            If (StrComp(PtrToStrW(Info(i).MimeType), strMimeType, vbTextCompare) = 0) Then '必须把指针转换成可用的字符
                CopyMemory ClassID(0), Info(i).ClassID(0), 16                   '保存类的ID
                GetEncoderClsID = i                                             '返回成功的索引值
                Exit For
            End If
        Next
    End If
End Function

Public Function PtrToStrW(ByVal lpsz As Long) As String
    Dim Out         As String
    Dim Length      As Long
    Length = lstrlenW(lpsz)
    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If
End Function

