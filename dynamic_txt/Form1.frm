VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'©2001 by Alexander Anikin
'e-mail: aka@i.com.ua
'for more my code samples
'visit my personal web site:
'http://www.i.com.ua/~aka
Dim a As String
Dim b As Integer
Private Declare Sub Sleep _
Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Start()
Do
For b = 16 To 64 Step 4
Sleep 1
Cls
ForeColor = RGB(80 + b, 0, 0)
Font.Size = b
CurrentX = (ScaleWidth - TextWidth(a)) \ 2
CurrentY = (ScaleHeight - TextHeight(a)) \ 2
Print a
DoEvents
Next
For b = 64 To 16 Step -4
Sleep 1
Cls
ForeColor = RGB(155 + b, 0, 0)
Font.Size = b
CurrentX = (ScaleWidth - TextWidth(a)) \ 2
CurrentY = (ScaleHeight - TextHeight(a)) \ 2
Print a
DoEvents
Next
DoEvents
Loop
End Sub

Private Sub Form_Activate()
Call Start

End Sub

Private Sub Form_Load()
Caption = "for more code samles - http://www.i.com.ua/~aka"
BackColor = vbBlack
WindowState = vbNormal
AutoRedraw = True
Width = 6075
Height = 4725
Font.Name = "arial"

a = "cooLLook"

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
