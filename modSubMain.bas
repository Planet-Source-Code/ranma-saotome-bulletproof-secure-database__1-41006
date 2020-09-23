Attribute VB_Name = "modSubMain"
Option Explicit


Sub Main()
'no double starts please
If App.PrevInstance = True Then Exit Sub
'works under nt4.0,win2k and xp (not tested under win9x)
On Error Resume Next
App.TaskVisible = False

frmLogon.Show
End Sub
