VERSION 5.00
Begin VB.Form frmNim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ZX-NIM - a VB port from the old Spectrum Version"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      Height          =   3300
      Left            =   240
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   0
      Top             =   240
      Width           =   3360
   End
End
Attribute VB_Name = "frmNim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
'this is the fastest way to implement the keyboard input
    modZXBasic.AddKey KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmNim = Nothing
End 'bad coding practice but must be done because of a possible pending INPUTs or INKEY$ statement
End Sub

Private Sub picView_KeyPress(KeyAscii As Integer)
    modZXBasic.AddKey KeyAscii
End Sub
