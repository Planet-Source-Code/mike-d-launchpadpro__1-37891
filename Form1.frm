VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launchpad Pro"
   ClientHeight    =   6105
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Controls:"
      Height          =   3135
      Left            =   4920
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
      Begin VB.CommandButton CmdHelp 
         Caption         =   "Help"
         Height          =   495
         Left            =   2040
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton CmdOptions 
         Caption         =   "Options"
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear List"
         Height          =   495
         Left            =   2040
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save List"
         Height          =   495
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "Open List"
         Height          =   495
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdLaunchAll 
         Caption         =   "Launch All"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton CmdLaunch 
         Caption         =   "Launch"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton CmdDeselect 
         Caption         =   "Deselect Item"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "Remove Item"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add Item"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Welcome:"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4695
      Begin VB.Label Label1 
         Caption         =   "Welcome to Launchpad Pro. A Free Open Source Program Launch Utility."
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1535
      ButtonWidth     =   1508
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open a Saved List"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save Current List"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Launch"
            Key             =   "Launch"
            Object.ToolTipText     =   "Launch Selected Item"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Launch All"
            Key             =   "LaunchAll"
            Object.ToolTipText     =   "Launch All Items in List"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "Add"
            Object.ToolTipText     =   "Add Item to List"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "Remove"
            Object.ToolTipText     =   "Remove Selected Item"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Deselect"
            Key             =   "Deselect"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "Clear"
            Object.ToolTipText     =   "Clear Current List"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "Options"
            Object.ToolTipText     =   "Open Options Control Panel"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Open Help File"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":091A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":128A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":181E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2986
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":350A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog2 
      Left            =   4920
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   5400
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "6:31 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/10/02"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Launch Items:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "&Controls"
      Begin VB.Menu mnuControlsAdd 
         Caption         =   "Add Item"
      End
      Begin VB.Menu mnuControlsRemove 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnuControlsDeselect 
         Caption         =   "Deselect"
      End
      Begin VB.Menu mnuControlsLaunch 
         Caption         =   "Launch"
      End
      Begin VB.Menu mnuControlsLaunchAll 
         Caption         =   "Launch All"
      End
      Begin VB.Menu mnuControlsOpenList 
         Caption         =   "Open List"
      End
      Begin VB.Menu mnuControlsSaveList 
         Caption         =   "Save List"
      End
      Begin VB.Menu mnuControlsClear 
         Caption         =   "Clear List"
      End
      Begin VB.Menu mnuControlsOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuHelpReadMe 
         Caption         =   "Read Me"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
Call Add
End Sub

Private Sub CmdClear_Click()
Call Clear
End Sub

Private Sub CmdDeselect_Click()
Call Deselect
End Sub

Private Sub CmdHelp_Click()
Call Help
End Sub

Private Sub CmdLaunch_Click()
Call CLaunch
End Sub

Private Sub CmdLaunchAll_Click()
Call CLaunchAll
End Sub

Private Sub CmdOpen_Click()
Call OpenList
End Sub

Private Sub CmdOptions_Click()
Call Options
End Sub

Private Sub CmdRemove_Click()
Call Remove
End Sub

Private Sub CmdSave_Click()
Call SaveList
End Sub

Private Sub Form_Load()
Call StartConfig

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim AForm As Form
For Each AForm In Forms
Unload AForm
Next

End Sub

Private Sub mnuControlsAdd_Click()
Call Add
End Sub

Private Sub mnuControlsClear_Click()
Call Clear
End Sub

Private Sub mnuControlsDeselect_Click()
Call Deselect
End Sub

Private Sub mnuControlsLaunch_Click()
Call CLaunch
End Sub

Private Sub mnuControlsLaunchAll_Click()
Call CLaunchAll
End Sub

Private Sub mnuControlsOpenList_Click()
Call OpenList
End Sub

Private Sub mnuControlsOptions_Click()
Call Options
End Sub

Private Sub mnuControlsRemove_Click()
Call Remove
End Sub

Private Sub mnuControlsSaveList_Click()
Call SaveList
End Sub

Private Sub mnuFileExit_Click()
Unload Form1

End Sub

Private Sub mnuHelpHelp_Click()
Call Help
End Sub

Private Sub mnuHelpReadMe_Click()
Call ReadMe
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Open"
Call OpenList
Case "Save"
Call SaveList
Case "Add"
Call Add
Case "Remove"
Call Remove
Case "Deselect"
Call Deselect
Case "Clear"
Call Clear
Case "Help"
Call Help
Case "Launch"
Call CLaunch
Case "LaunchAll"
Call CLaunchAll
Case "Options"
Call Options
End Select

End Sub
