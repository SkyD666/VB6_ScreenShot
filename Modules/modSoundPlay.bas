Attribute VB_Name = "modSoundPlay"
Option Explicit

'��������������������������������������������������ͼ�󲥷���ʾ��
'API �����õ��ĳ���
Private Const SND_SYNC = &H0                                                    'ͬ��
Private Const SND_ASYNC = &H1                                                   '�첽
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8                                                    'ѭ������
 
'API����
Private Declare Function sndPlaySoundFromMemory Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
'������������������������������������������������

Public Sub SoundPlay()
    Dim myMusic() As Byte
    'Res �ļ���ȡʱҲ�����ִ�Сд
    myMusic = LoadResData(ChooseSoundPlayStr, "SOUND")
    sndPlaySoundFromMemory myMusic(0), SND_MEMORY Or SND_ASYNC
End Sub
