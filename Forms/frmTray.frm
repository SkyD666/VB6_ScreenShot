VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   0  'None
   Caption         =   "ScreenSnap Õ–≈Ã≥Ã–Ú"
   ClientHeight    =   450
   ClientLeft      =   -30
   ClientTop       =   -30
   ClientWidth     =   945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Top = -Me.Width - 100
    Me.Left = -Me.Height - 100
End Sub
