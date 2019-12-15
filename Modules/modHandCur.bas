Attribute VB_Name = "modHandCur"
Option Explicit

Public Const IDC_HAND As Long = 32649&

Public Declare Function LoadCursorA Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public hHandCur As Long
