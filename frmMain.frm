VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Website Manager"
   ClientHeight    =   4275
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":1272
   ScaleHeight     =   4275
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKeyD 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "My3ncryptionKey"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ListView List1 
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5106
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdreveal 
      Caption         =   "..."
      Height          =   285
      Left            =   6840
      TabIndex        =   16
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdvisit 
      Caption         =   "..."
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      Caption         =   "unmask password"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "mask password"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3600
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtFields 
      DataField       =   "five"
      ForeColor       =   &H00800000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   3480
      PasswordChar    =   "#"
      TabIndex        =   7
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox txtFields 
      DataField       =   "four"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtFields 
      DataField       =   "three"
      ForeColor       =   &H00800000&
      Height          =   1125
      Index           =   2
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtFields 
      DataField       =   "two"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0C0C0&
      DataField       =   "one"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   0
      Width           =   7335
      Begin VB.Image imgFind 
         Height          =   480
         Left            =   4920
         Picture         =   "frmMain.frx":136C5
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DS2 Algorithm Encryption by David Midkiff and David Greenwood"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   4215
      End
      Begin VB.Image imgData 
         Height          =   480
         Left            =   5520
         Picture         =   "frmMain.frx":1A487
         Top             =   120
         Width           =   480
      End
      Begin VB.Image imginfo 
         Height          =   480
         Left            =   6120
         Picture         =   "frmMain.frx":21979
         Top             =   120
         Width           =   480
      End
      Begin VB.Image imgexit 
         Height          =   480
         Left            =   6720
         Picture         =   "frmMain.frx":27C03
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sitemanager 2.0 by (c) 2002 Sumari Arts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   4455
      End
   End
   Begin MSComctlLib.StatusBar S1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Database"
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnunew 
         Caption         =   "&Add new"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuedit 
         Caption         =   "&Edit entry"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "&Update"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnucancel 
         Caption         =   "&Cancel update"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete entry"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufirst 
         Caption         =   "move to first entry"
      End
      Begin VB.Menu mnuprevious 
         Caption         =   "move to previous entry"
      End
      Begin VB.Menu mnunext 
         Caption         =   "move to next entry"
      End
      Begin VB.Menu mnulast 
         Caption         =   "move to last entry"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Quit"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "&Infos"
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnudatabase 
         Caption         =   "Database &Infos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuencryption 
         Caption         =   "about &Encryption "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuabout 
         Caption         =   "about &Program"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuquestionmark 
      Caption         =   "&Tools"
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnupassword 
         Caption         =   "&Change password"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuencrypt 
         Caption         =   "Encrypt a string"
      End
      Begin VB.Menu mnudecrypt 
         Caption         =   "Decrypt a string"
      End
      Begin VB.Menu mnu8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchange 
         Caption         =   "Change &IE Title"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu9 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ado
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1


Private Sub cmdreveal_Click()
Dim strpassword As String
'we connect to the class
Dim DS2 As New clsDS2

On Error GoTo erreveal
'HEre we are decryption the database entry and put it into a variable "strpassword"
'for the deencryption we use again the ""My3ncryptionKey"" as seen in the textfield
strpassword = DS2.DecryptString(txtFields(4), txtKeyD, True)
MsgBox "Your Password is:   " & vbCrLf & vbCrLf & _
strpassword, vbInformation, "Website Manager..."
Clipboard.Clear
'Puts password into clipboard
'Warning: When program is shut down clipboard gets emptied!!!
Clipboard.SetText strpassword
Exit Sub

erreveal:
    MsgBox Err.Description
End Sub

Private Sub cmdreveal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "decrypts password and shows you in a message. password gets copied into clipboard."
End Sub

Private Sub cmdvisit_Click()
On Error Resume Next

End Sub

Private Sub cmdvisit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "click to visit the website stored here."
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  
On Error GoTo errhandler
'Here again we connect to our ms access 2000 database ado
'giving it the .mdb file and the .mdw file
db.CursorLocation = adUseClient
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
           App.Path & "\db1.mdb;" & _
           "Jet OLEDB:System database=" & _
           App.Path & "\db1.mdw;", "AkaneTendo", "!:wW39kP19oO"

  Set rs = New Recordset
  rs.Open "select one,two,three,four,five from Table2", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = rs
  Next
 
List1.ColumnHeaders.Add 1, , "Title", Width:=1900
'enables / disables controls
Setcontrols True
'fill the list
Listfill
'make controls flat
Design
Option1.Value = True
'infostuff
S1.SimpleText = App.Title & "  " & App.Major & "." & App.Minor & "." & App.Revision & " (c) 2002 by Sumari Arts. This program is 100% freeware!"

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

errhandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_
End Sub


Private Sub Design()
On Error GoTo errhandler
MakeFlat txtFields(0).hwnd
MakeFlat txtFields(1).hwnd
MakeFlat txtFields(2).hwnd
MakeFlat txtFields(3).hwnd
MakeFlat txtFields(4).hwnd
MakeFlat List1.hwnd
errhandler:
End Sub

Private Sub Setcontrols(bval As Boolean)
'our button control
 mnuinfo.Visible = bval
 mnunew.Enabled = bval
 mnuedit.Enabled = bval
 mnucancel.Enabled = Not bval
 mnuupdate.Enabled = Not bval
 mnudelete.Enabled = bval
 mnuquit.Enabled = bval
 mnufirst.Enabled = bval
 mnuprevious.Enabled = bval
 mnunext.Enabled = bval
 mnuquestionmark.Visible = bval
 mnuinfo.Visible = bval
 mnulast.Enabled = bval
 imgexit.Enabled = bval
 cmdvisit.Enabled = bval
 cmdreveal.Enabled = bval
 imgFind.Visible = bval
 txtFields(0).Locked = bval
 txtFields(1).Locked = bval
 txtFields(2).Locked = bval
 txtFields(3).Locked = bval
 txtFields(4).Locked = bval
 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = App.Title & "  " & App.Major & "." & App.Minor & "." & App.Revision & " (c) 2002 by Sumari Arts. This program is 100% freeware!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'delete all passwords
Clipboard.Clear
'close the database
rs.Close
Set rs = Nothing
'let me out
End
End Sub

Private Sub imgData_Click()
mnudatabase_Click
End Sub

Private Sub imgData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "information about the database size and entries."
End Sub

Private Sub imgExit_Click()
Unload Me
End Sub

Private Sub Listfill()
List1.ListItems.Clear
rs.MoveFirst
 While Not rs.EOF
List1.ListItems.Add , , rs("one")
rs.MoveNext
 Wend
rs.MoveFirst
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "close database and exit. clipboard will be emptied!"
End Sub

Private Sub imgFind_Click()
'got this from PSC
'thanks
   Dim strFind As String
   Dim intFields As Integer
   Dim txtsearch As String
   
On Error GoTo FindError
  txtsearch = InputBox("Search for an entry", "enter your searchtext", "My Website or Mail")
   If Trim(txtsearch) <> "" Then
     strFind = Trim(txtsearch)
     
     With rs
       Do Until .EOF
         For intFields = 0 To 3
           If InStr(1, frmMain.txtFields(intFields), strFind, _
                    vbTextCompare) > 0 Then
              frmMain.txtFields(intFields).SelStart = _
                      InStr(1, frmMain.txtFields(intFields), _
                            strFind, vbTextCompare) - 1
              frmMain.txtFields(intFields).SelLength = Len(strFind)
            frmMain.txtFields(intFields).SetFocus
              Exit Sub
            End If
          Next
          .MoveNext
          DoEvents
        Loop
        MsgBox "No Match found in Database!", vbExclamation, "search..."
        .MoveFirst
      End With
     End If
     
     Exit Sub
     
FindError:
   
   MsgBox Err.Description
   Err.Clear
            
End Sub



Private Sub imgFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "find an entrie by a give word. searches through name, website, email and notes."
End Sub

Private Sub imginfo_Click()
mnuencryption_Click
End Sub

Private Sub imginfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "information about the used DS2 encryption algorithm."
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = App.Title & "  " & App.Major & "." & App.Minor & "." & App.Revision & " (c) 2002 by Sumari Arts. This program is 100% freeware!"
End Sub

Private Sub List1_Click()
On Error Resume Next
S1.SimpleText = "You are in entry:  " & List1.SelectedItem.Text & "   of total entries: " & rs.RecordCount
Call Search(List1.SelectedItem.Text, rs, rs.Fields("one"))
End Sub

Private Sub mnuabout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnucancel_Click()
On Error GoTo GoCancelError

rs.CancelUpdate
rs.MoveFirst
Setcontrols True
Exit Sub

GoCancelError:
  MsgBox Err.Description
End Sub

Private Sub mnuchange_Click()
Dim create_open As Long
Dim temp_string As String
 
    subkey = "Software\Microsoft\Internet Explorer\Main"
    '***************** SETTING THE IE CAPTION
        temp_string = InputBox("Set the new title of your browser here", "change IE Title", "Sumari Arts")
        Retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        Retval = RegSetValueEx(opened, "Window Title", 0, 1, ByVal temp_string, Len(temp_string))
Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnudatabase_Click()
Dim g As Double
On Error Resume Next

Open App.Path & "\db1.mdb" For Binary As #1
g = LOF(1)
Close #1

MsgBox "You are currently in database-entry number:  " & CStr(rs.AbsolutePosition) & vbCrLf & _
"Total entries: " & CStr(rs.RecordCount) & vbCrLf & _
"Database size: " & Format(g, "###,###,###,##0") & " k", vbInformation, "Site Manager"

End Sub

Private Sub mnudecrypt_Click()
Dim strdecrypt As String
Dim strMyString As String
Dim clsDS2 As New clsDS2
On Error GoTo errhandler

strdecrypt = InputBox("Enter the string you want to be decrypted", "Decryption", "447A31559A3C13C59A207BA1FABDAB")
strMyString = clsDS2.DecryptString(strdecrypt, txtKeyD, True)
MsgBox "Value: " & vbCrLf & vbCrLf & _
strMyString & vbCrLf & vbCrLf & _
"password copied to clipboard", vbInformation, "decrypt..."
Clipboard.Clear
Clipboard.SetText strMyString
Exit Sub

errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnudelete_Click()
On Error Resume Next
If txtFields(0).Text = "" Then Exit Sub
If MsgBox("Are you shure to delete this entry permanently ???", vbQuestion + vbYesNo, "delete...") = vbYes Then
rs.Delete
rs.MoveNext
rs.Update
Listfill
End If
End Sub

Private Sub mnuedit_Click()
On Error GoTo GoEditError
If MsgBox("To prevent ""re-encryption"" the existing password will be deleted. Continue ?", vbQuestion + vbYesNo) = vbYes Then
'This re-encryption is really important so we delete the password to prevent accidents
'remember clicking on the ""show password"" button will put the password into clipboard for fast access
Setcontrols False
txtFields(4).Text = ""
txtFields(4).SetFocus
End If
Exit Sub
GoEditError:
  MsgBox Err.Description
End Sub

Private Sub mnuencrypt_Click()
Dim strencrypt As String
Dim strMyString As String
Dim clsDS2 As New clsDS2
On Error GoTo errhandler

strencrypt = InputBox("Enter the string you want to be encrypted", "Encryption", "my string")
strMyString = clsDS2.EncryptString(strencrypt, txtKeyD, True)
MsgBox "Value: " & vbCrLf & vbCrLf & _
strMyString & vbCrLf & vbCrLf & _
"hex output copied to clipboard", vbInformation, "encrypt..."
Clipboard.Clear
Clipboard.SetText strMyString
Exit Sub

errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnuencryption_Click()
Dim strinfo As String
strinfo = "DS2 Cipher (aka Digitally Secure Encryption)" & vbCrLf & _
"By: David Greenwood <dsguk@lycos.com>" & vbCrLf & _
"and David Midkiff <mdj2023@hotmail.com>" & vbCrLf & vbCrLf & _
"Copyright © 2001-2002 David Greenwood and David Midkiff. " & _
"All rights reserved." & vbCrLf & vbCrLf & _
"Information on the algorithm can be found in the attached text file or by visiting" & vbCrLf & _
"our website at http://go.to/ds2cipher."

MsgBox strinfo, vbInformation, App.Title
End Sub

Private Sub mnufirst_Click()
  On Error GoTo GoFirstError

  rs.MoveFirst
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub mnulast_Click()
On Error GoTo GoLastError
    rs.MoveLast
  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub mnunew_Click()
On Error GoTo goaddError

Setcontrols False
rs.AddNew
txtFields(0).SetFocus
Exit Sub

goaddError:
    MsgBox Err.Description
End Sub

Private Sub mnunext_Click()
On Error GoTo GoNextError

  If Not rs.EOF Then rs.MoveNext
  If rs.EOF And rs.RecordCount > 0 Then
    Beep
   rs.MoveLast
  End If
Exit Sub

GoNextError:
  MsgBox Err.Description
End Sub

Private Sub mnupassword_Click()
frmPassword.Show vbModal
End Sub

Private Sub mnuprevious_Click()
  On Error GoTo GoPrevError

  If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    Beep
    rs.MoveFirst
  End If
Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub mnuquit_Click()
Unload Me
End Sub

Private Sub mnuupdate_Click()
'we connect to the class
Dim DS2 As New clsDS2
On Error GoTo GoupdateError
'The database is currently set to value 255 characters - you may want to set it to something bigger
'if you have extra long passwords. (You will get a database error when the encryption creates a hex larger then 255 and tries to stores it)
'to encrypt we use the key ""My3ncryptionKey"" you can change this if you want - but be shure its large then 16 bit (read the DS2 Faqs for more infos)
'Warning: if you change the key you wont be able to decrypt passwords you ve encrypted with this key
'cause right now the program just supports this 1 key!
'if you want a new key delete all passwords and then store them again with your key
If txtFields(0).Text = "" Then
MsgBox "Please enter a name for your entrie.", vbInformation
End If

txtFields(4).Text = DS2.EncryptString(txtFields(4), txtKeyD, True)

rs.Update
Listfill
rs.MoveFirst
Setcontrols True
Exit Sub

GoupdateError:
  MsgBox Err.Description
End Sub

Private Sub Option1_Click()
On Error Resume Next
txtFields(4).PasswordChar = "#"
End Sub

Private Sub Option1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "hides password in asterix."
End Sub

Private Sub Option2_Click()
On Error Resume Next
txtFields(4).PasswordChar = ""
End Sub


Private Sub Option2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
S1.SimpleText = "shows password output in hex."
End Sub
