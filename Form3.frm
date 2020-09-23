VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   Caption         =   "Launchpad Doc Viewer"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTF 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form3.frx":0442
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RTF_KeyPress(KeyAscii As Integer)
If KeyAscii <> 0 Then
KeyAscii = 0
End If




End Sub
