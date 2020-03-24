Attribute VB_Name = "modAsciiToName"
Option Explicit

Public Function AsciiToName(ByVal code As Integer) As String
    Select Case code
    Case 8
        AsciiToName = "Backspace"
    Case 9
        AsciiToName = "Tab"
    Case 13
        AsciiToName = "Enter"
    Case 19
        AsciiToName = "Pause"
    Case 20
        AsciiToName = "CapsLock"
    Case 27
        AsciiToName = "Esc"
    Case 32
        AsciiToName = "�ո�"
    Case 33
        AsciiToName = "PageUp"
    Case 34
        AsciiToName = "PageDown"
    Case 35
        AsciiToName = "End"
    Case 36
        AsciiToName = "Home"
    Case 37
        AsciiToName = "�����(��)"
    Case 38
        AsciiToName = "�����(��)"
    Case 39
        AsciiToName = "�����(��)"
    Case 40
        AsciiToName = "�����(��)"
    Case 44
        AsciiToName = "PrtSc"
    Case 45
        AsciiToName = "Insert"
    Case 46
        AsciiToName = "Delete"
    Case 93
        AsciiToName = "�˵���(appskey)"
    Case 106
        AsciiToName = "*"
    Case 107
        AsciiToName = "+"
    Case 110
        AsciiToName = "."
    Case 144
        AsciiToName = "NumLK"
    Case 145
        AsciiToName = "ScrLK"
    Case 160
        AsciiToName = "��Shift"
    Case 161
        AsciiToName = "��Shift"
    Case 162
        AsciiToName = "��Ctrl"
    Case 163
        AsciiToName = "��Ctrl"
    Case 164
        AsciiToName = "��Alt"
    Case 165
        AsciiToName = "��Alt"
    Case 189
        AsciiToName = "-"
    Case 109
        AsciiToName = "-"
    Case 187
        AsciiToName = "="
    Case 192
        AsciiToName = "`"
    Case 219
        AsciiToName = "["
    Case 221
        AsciiToName = "]"
    Case 186
        AsciiToName = ";"
    Case 222
        AsciiToName = "'"
    Case 220
        AsciiToName = "\"
    Case 188
        AsciiToName = ","
    Case 190
        AsciiToName = "."
    Case 191
        AsciiToName = "/"
    Case 111
        AsciiToName = "/"
    Case 193
        AsciiToName = "\"
    Case 112
        AsciiToName = "F1"
    Case 113
        AsciiToName = "F2"
    Case 114
        AsciiToName = "F3"
    Case 115
        AsciiToName = "F4"
    Case 116
        AsciiToName = "F5"
    Case 117
        AsciiToName = "F6"
    Case 118
        AsciiToName = "F7"
    Case 119
        AsciiToName = "F8"
    Case 120
        AsciiToName = "F9"
    Case 121
        AsciiToName = "F10"
    Case 122
        AsciiToName = "F11"
    Case 123
        AsciiToName = "F12"
    Case 97
        AsciiToName = "С����1"
    Case 98
        AsciiToName = "С����2"
    Case 99
        AsciiToName = "С����3"
    Case 100
        AsciiToName = "С����4"
    Case 101
        AsciiToName = "С����5"
    Case 102
        AsciiToName = "С����6"
    Case 103
        AsciiToName = "С����7"
    Case 104
        AsciiToName = "С����8"
    Case 105
        AsciiToName = "С����9"
    Case Else
        AsciiToName = Chr(code)
    End Select
End Function
