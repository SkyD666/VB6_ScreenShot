Attribute VB_Name = "modSoundPlay"
Option Explicit

'————————————————————————截图后播放提示音
'API 函数用到的常数
Private Const SND_SYNC = &H0                                                    '同步
Private Const SND_ASYNC = &H1                                                   '异步
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8                                                    '循环播放
 
'API函数
Private Declare Function sndPlaySoundFromMemory Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
'————————————————————————

Public Sub SoundPlay()
    Dim myMusic() As Byte
    'Res 文件读取时也不区分大小写
    myMusic = LoadResData(ChooseSoundPlayStr, "SOUND")
    sndPlaySoundFromMemory myMusic(0), SND_MEMORY Or SND_ASYNC
End Sub
