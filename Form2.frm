VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog3 
      Left            =   240
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton CmdSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      ToolTipText     =   "Set List to automatically load."
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtlist 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4455
   End
   Begin VB.CheckBox chklist 
      Caption         =   "Automatically load and launch a list when Launchpad starts."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.CheckBox chkpass 
      Caption         =   "Prompt me for a password before launching any programs."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Default List:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chklist_Click()
Form2.txtlist.Text = ""

End Sub

Private Sub chkpass_Click()
Form2.txtpass.Text = ""

End Sub

Private Sub CmdCancel_Click()
Unload Form2

End Sub

Private Sub CmdOk_Click()
Call OptionReset

End Sub

Private Sub CmdSet_Click()
Call SetList
End Sub

