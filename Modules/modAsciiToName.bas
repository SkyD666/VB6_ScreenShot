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
        AsciiToName = "空格"
    Case 33
        AsciiToName = "PageUp"
    Case 34
        AsciiToName = "PageDown"
    Case 35
        AsciiToName = "End"
    Case 36
        AsciiToName = "Home"
    Case 37
        AsciiToName = "方向键(←)"
    Case 38
        AsciiToName = "方向键(↑)"
    Case 39
        AsciiToName = "方向键(→)"
    Case 40
        AsciiToName = "方向键(↓)"
    Case 44
        AsciiToName = "PrtSc"
    Case 45
        AsciiToName = "Insert"
    Case 46
        AsciiToName = "Delete"
    Case 93
        AsciiToName = "菜单键(appskey)"
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
        AsciiToName = "左Shift"
    Case 161
        AsciiToName = "右Shift"
    Case 162
        AsciiToName = "左Ctrl"
    Case 163
        AsciiToName = "右Ctrl"
    Case 164
        AsciiToName = "左Alt"
    Case 165
        AsciiToName = "右Alt"
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
        AsciiToName = "小键盘1"
    Case 98
        AsciiToName = "小键盘2"
    Case 99
        AsciiToName = "小键盘3"
    Case 100
        AsciiToName = "小键盘4"
    Case 101
        AsciiToName = "小键盘5"
    Case 102
        AsciiToName = "小键盘6"
    Case 103
        AsciiToName = "小键盘7"
    Case 104
        AsciiToName = "小键盘8"
    Case 105
        AsciiToName = "小键盘9"
    Case Else
        AsciiToName = Chr(code)
    End Select
End Function
