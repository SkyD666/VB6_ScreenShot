Attribute VB_Name = "uF_vb"
' uFMOD header file
' Target OS: Windows
' Compiler:  Visual Basic 6
' Driver:    WINMM

' WARNING!!!
' Do not rename / modify this file!
' Read the README file before trying to use uFMOD with VB6!

Public Const XM_RESOURCE As Long = 0
Public Const XM_MEMORY As Long = 1
Public Const XM_FILE As Long = 2
Public Const XM_NOLOOP As Long = 8
Public Const XM_SUSPENDED As Long = 16

' The uFMOD_PlaySong function starts playing an XM song.
' --------------
' Parameters:
'   lpXM
'      Specifies the song to play. If this parameter is NULL,
'      any currently playing song is stopped. In such a case, function
'      does not return a meaningful value.
'      fdwSong parameter determines whether this value is interpreted
'      as a filename, as a resource identifier or a pointer to an image
'      of the song in memory.
'   param
'      Handle to the executable file that contains the resource to be
'      loaded or size of the image of the song in memory. This parameter
'      is ignored unless XM_RESOURCE or XM_MEMORY is specified in fdwSong.
'   fdwSong
'      Flags for playing the song. The following values are defined:
'      Value        Meaning
'      XM_FILE      lpXM points to filename.
'                   param is ignored.
'      XM_MEMORY    lpXM points to an image of a song in memory.
'                   param is the image size. Once, uFMOD_PlaySong
'                   returns, it's safe to free/discard the memory buffer.
'      XM_RESOURCE  lpXM Specifies the name of the resource.
'                   param identifies the module whose executable file
'                   contains the resource.
'                   The resource type must be RT_RCDATA.
'      XM_NOLOOP    An XM track plays repeatedly by default. Specify
'                   this flag to play it only once.
'      XM_SUSPENDED The XM track is loaded in a suspended state,
'                   and will not play until the uFMOD_Resume function
'                   is called. This is useful for preloading a song
'                   or testing an XM track for validity.
' Return Values:
'    Returns a pointer to HWAVEOUT on success or NULL otherwise.
Function uFMOD_PlaySong(ByVal lpXM As Long, ByVal param As Long, ByVal fdwSong As Long) As Long
    MsgBox "Check the readme.txt for details", vbCritical, "uFMOD not found"
    uFMOD_PlaySong = 0
End Function

' The uFMOD_Pause function pauses the currently playing song, if any.
Sub uFMOD_Pause()
End Sub

' The uFMOD_Resume function resumes the currently paused song, if any.
Sub uFMOD_Resume()
End Sub

' The uFMOD_GetStats function returns the current RMS volume coefficients
' in L and R channels.
' --------------
' Return Values:
'    low-order word : RMS volume in R channel
'    hi-order  word : RMS volume in L channel
Function uFMOD_GetStats() As Integer
    uFMOD_GetStats = 0
End Function

' The uFMOD_GetTime function returns the time in milliseconds since the
' song was started. This is useful for synchronizing purposes.
' --------------
' Return Values:
'    Returns the time in milliseconds since the song was started.
Function uFMOD_GetTime() As Long
    uFMOD_GetTime = 0
End Function

' The uFMOD_GetTitle function returns the current track's title, if any.
' --------------
' Return Values:
'    Returns the track's title in ASCIIZ format.
Function uFMOD_GetTitle() As Long
    uFMOD_GetTitle = 0
End Function

' The uFMOD_SetVolume function sets the global volume.
' --------------
' 0:  muting
' 64: maximum volume
' NOTE: Any value above 64 maps to maximum volume too.
' The volume scale is linear. Maximum volume is set by default.
Sub uFMOD_SetVolume(ByVal vol As Long)
End Sub
