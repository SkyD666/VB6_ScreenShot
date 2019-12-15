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
    
    Gsp.GdiplusVersion = 1                                                      'GDI+ 1.0�汾
    GdiplusStartup Token, Gsp                                                   '��ʼ��GDI+
    If StdPicLongBoo = False Then
        GdipCreateBitmapFromHBITMAP StdPic.Handle, StdPic.hPal, Bitmap
    Else
        Bitmap = StdPicLongLng
    End If
    If Bitmap <> 0 Then                                                         '˵�����ǳɹ��Ľ�StdPic����ת��ΪGDI+��Bitmap������
        Select Case FileFormat
        Case "bmp"
            If Not GetEncoderClsID("Image/bmp", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), ByVal 0) = 0)
            End If
        Case "jpg"                                                              'JPG��ʽ�������ñ��������
            Dim aEncParams()        As Byte
            Dim uEncParams          As EncoderParameters
            If GetEncoderClsID("Image/jpeg", CLSID) <> -1 Then
                uEncParams.Count = 1                                            ' �����Զ���ı������������Ϊ1������
                If JpgQuality < 0 Then
                    JpgQuality = 0
                ElseIf JpgQuality > 100 Then
                    JpgQuality = 100
                End If
                ReDim aEncParams(1 To Len(uEncParams))
                With uEncParams.Parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong                       ' ���ò���ֵ����������Ϊ������
                    Call CLSIDFromString(StrPtr(EncoderQuality), .Guid(0))      ' ���ò���Ψһ��־��GUID������Ϊ����Ʒ��
                    .Value = VarPtr(JpgQuality)                                 ' ���ò�����ֵ��Ʒ�ʵȼ������Ϊ100��ͼ���ļ���С��Ʒ�ʳ�����
                End With
                CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), aEncParams(1)) = 0)
            End If
        Case "png"
            If Not GetEncoderClsID("Image/png", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), ByVal 0) = 0)
            End If
        Case "gif"
            If Not GetEncoderClsID("Image/gif", CLSID) = -1 Then                '���ԭʼ��ͼ����24λ����������������ϵͳ�ĵ�ɫ������ͼ��ת��Ϊ8λ��ת����Ч���᲻������,��Ҳ�п���ϵͳ���Զ�ת��������ʧ��
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(filename), CLSID(0), ByVal 0) = 0)
            End If
        End Select
    End If
    GdipDisposeImage Bitmap                                                     'ע���ͷ���Դ
    GdiplusShutdown Token                                                       '�ر�GDI+��
End Function

Public Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim size        As Long
    Dim i           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, size                                          '�õ�����������Ĵ�С
    If size <> 0 Then
        ReDim Info(1 To Num) As ImageCodecInfo                                  '�����鶯̬�����ڴ�
        ReDim Buffer(1 To size) As Byte
        GdipGetImageEncoders Num, size, Buffer(1)                               '�õ�������ַ�����
        CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)                     '������ͷ
        For i = 1 To Num                                                        'ѭ��������н���
            If (StrComp(PtrToStrW(Info(i).MimeType), strMimeType, vbTextCompare) = 0) Then '�����ָ��ת���ɿ��õ��ַ�
                CopyMemory ClassID(0), Info(i).ClassID(0), 16                   '�������ID
                GetEncoderClsID = i                                             '���سɹ�������ֵ
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

