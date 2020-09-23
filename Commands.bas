Attribute VB_Name = "Commands"
Public Sub SetList()
On Error GoTo Error1
Form2.Dialog3.DialogTitle = "Choose a default list"
Form2.Dialog3.Filter = "TextFiles (*.txt)|*.txt"
Form2.Dialog3.DefaultExt = "txt"
Form2.Dialog3.ShowOpen
gstrset = Form2.Dialog3.FileName
Form2.txtlist.Text = gstrset

Error1:
Exit Sub


End Sub


Public Sub OptionReset()
On Error GoTo Error2

If Form2.chkpass.Value = 1 And Form2.txtpass.Text = "" Then
MsgBox "You haven't chosen a password!"
Exit Sub
End If
If Form2.chklist.Value = 1 And Form2.txtlist.Text = "" Then
MsgBox "Please choose you default list!"
Exit Sub
End If
Call OptionGoReset



Error2:
Exit Sub

End Sub


Public Sub OptionGoReset()
On Error GoTo Error3

If Form2.chkpass.Value = 1 Then
gblnnewneedpass = True
Else
gblnnewneedpass = False
End If
If Form2.chklist.Value = 1 Then
gblnnewloadlist = True
Else
gblnnewloadlist = False
End If
gstrnewpass = Form2.txtpass.Text
gstrnewlist = Form2.txtlist.Text
SaveSetting "Launchpad", "Startup", "Needpass", gblnnewneedpass
SaveSetting "Launchpad", "Startup", "Password", gstrnewpass
SaveSetting "Launchpad", "Startup", "DefaultList", gblnnewloadlist
SaveSetting "Launchpad", "Startup", "Listname", gstrnewlist
MsgBox "You must restart Launchpad for changes to take effect."
Unload Form2



Error3:
Exit Sub

End Sub


Public Sub CLaunch()
On Error GoTo Error4

If gblnpass = False Then
Call Launch
Exit Sub
Else
gstrlogin = InputBox("Enter your password.", "Password?")
End If
If gstrlogin = gstrpass Then
Call Launch
Else
MsgBox "Wrong Password!"
Exit Sub
End If



Error4:
Exit Sub

End Sub


Public Sub CLaunchAll()
On Error GoTo Error5
If gblnpass = False Then
Call LaunchAll
Exit Sub
Else
gstrlogin = InputBox("Enter your password.", "Password?")
End If
If gstrlogin = gstrpass Then
Call LaunchAll
Exit Sub
Else
MsgBox "Wrong Password!"
Exit Sub
End If

Error5:
Exit Sub

End Sub



Public Sub Help()
On Error GoTo Error6
Form3.Show
Form3.RTF.LoadFile "C:\Program Files\LaunchpadPro\help.rtf"
Error6:
Exit Sub

End Sub



Public Sub ReadMe()
On Error GoTo Error7
Form3.Show
Form3.RTF.LoadFile "C:\Program Files\LaunchpadPro\readme.txt"
Error7:
Exit Sub

End Sub



Public Sub Clear()
Form1.List1.Clear

End Sub


Public Sub Deselect()
Form1.List1.Text = -1

End Sub



Public Sub LaunchAll()
On Error GoTo Error8
For gintindex = 0 To Form1.List1.ListCount - 1
Shell Form1.List1.List(gintindex)
Next
Error8:
Exit Sub
End Sub


Public Sub Remove()
On Error GoTo Error9
If Form1.List1.Text = "" Then
MsgBox "Select item to remove."
Exit Sub
End If
gstrremove = Form1.List1.Text
Form1.List1.RemoveItem Form1.List1.ListIndex
Error9:
Exit Sub

End Sub


Public Sub Add()
On Error GoTo Error10
Form1.Dialog2.DialogTitle = "Choose Item to Add"
Form1.Dialog2.Filter = "Executables (*.exe*.bat*.com*.pif)|*.exe;*.com;*.pif;*.bat"
Form1.Dialog2.ShowOpen
If Form1.Dialog2.FileName = "" Then
Exit Sub
End If
gstradd = Form1.Dialog2.FileName
Form1.List1.AddItem gstradd
Error10:
Exit Sub

End Sub



Public Sub Launch()
On Error GoTo Error11
If Form1.List1.Text = "" Then
MsgBox "Select Item to Launch"
Exit Sub
End If
gstrlaunch = Form1.List1.Text
Shell gstrlaunch

Error11:
Exit Sub

End Sub



Public Sub AdvancedMode()
On Error GoTo Error12
If gblnloadlist = False Then
Exit Sub
Else
Open gstrlist For Input As #1
Do While Not EOF(1)
Input #1, gstrpath
Form1.List1.AddItem gstrpath
Loop
Close #1
End If
Error12:
Exit Sub

End Sub



Public Sub OpenList()
On Error GoTo Error13
Form1.Dialog1.DialogTitle = "Open a List"
Form1.Dialog1.Filter = "Textfiles (*.txt)|*.txt"
Form1.Dialog1.ShowOpen
If Form1.Dialog1.FileName = "" Then
Exit Sub
End If
gstropen = Form1.Dialog1.FileName
Open gstropen For Input As #1
Form1.List1.Clear
Do While Not EOF(1)
Input #1, gstrpath
Form1.List1.AddItem gstrpath
Loop
Close #1
Error13:
Exit Sub

End Sub



Public Sub SaveList()
On Error GoTo Error14
Form1.Dialog1.DialogTitle = "Save Current List"
Form1.Dialog1.Filter = "Textfiles (*.txt)|*.txt"
Form1.Dialog1.DefaultExt = "txt"
Form1.Dialog1.ShowSave
If Form1.Dialog1.FileName = "" Then
MsgBox "Select a file name"
Exit Sub
Else
gstrsave = Form1.Dialog1.FileName
End If
Open gstrsave For Output As #1
For gintindex = 0 To Form1.List1.ListCount - 1
gstrpath = Form1.List1.List(gintindex)
Write #1, gstrpath
Next
Close #1
Error14:
Exit Sub

End Sub



Public Sub StartConfig()
On Error GoTo Error15
gstrpass = GetSetting("Launchpad", "Startup", "Password", "")
gstrlist = GetSetting("Launchpad", "Startup", "ListName", "")
gblnpass = GetSetting("Launchpad", "Startup", "Needpass", "False")
gblnloadlist = GetSetting("Launchpad", "Startup", "DefaultList", "False")
Call Security

Error15:
Exit Sub

End Sub




Public Sub Options()
On Error GoTo Error16
If gblnpass = False Then
Call OptionLauncher
Exit Sub
Else
gstrlogin = InputBox("Enter your password.", "Password?")
End If
If gstrlogin = gstrpass Then
Call OptionLauncher
Exit Sub
Else
MsgBox "Wrong Password!"
Exit Sub
End If
Error16:
Exit Sub

End Sub



Public Sub OptionLauncher()
On Error GoTo Error17
If gblnpass = True Then
Form2.chkpass.Value = 1
Else
Form2.chkpass.Value = 0
End If
If gblnloadlist = True Then
Form2.chklist.Value = 1
Else
Form2.chklist.Value = 0
End If
Form2.Show
Form2.txtpass.Text = gstrpass
Form2.txtlist.Text = gstrlist
Form2.Picture1.Picture = Form1.ImageList1.ListImages(8).Picture
Error17:
Exit Sub


End Sub


Public Sub Security()
On Error GoTo Error18
If gblnpass = False Then
Call AdvancedMode
Exit Sub
ElseIf gblnpass = True Then
gstrlogin = InputBox("Enter your password.", "Password?")
End If
If gstrlogin = gstrpass Then
Call AdvancedMode
Exit Sub
ElseIf gstrlogin <> gstrpass Then
MsgBox "Wrong Password!"
Unload Form1
End If

Error18:
Exit Sub

End Sub



