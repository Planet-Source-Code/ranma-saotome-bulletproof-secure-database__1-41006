VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change password"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassword.frx":1272
   ScaleHeight     =   2250
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   320
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   320
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtLogon 
      DataField       =   "two"
      ForeColor       =   &H00800000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "#"
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "one"
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Administrator"
      Top             =   840
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   0
      Width           =   4305
      Begin VB.Image imgExit 
         Height          =   480
         Left            =   3720
         Picture         =   "frmPassword.frx":5306
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "change username and password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
On Error Resume Next
Setcontrols True
rs.CancelUpdate
End Sub

Private Sub cmdEdit_Click()
On Error GoTo erredit
If MsgBox("To prevent ""re-encryption"" the existing password will be deleted. Continue ?", vbQuestion + vbYesNo) = vbYes Then
Setcontrols False
txtLogon.Text = ""
txtUser.SetFocus
End If
Exit Sub
erredit:
    MsgBox Err.Description
End Sub


Private Sub cmdUpdate_Click()
'we connect to the class
Dim DS2 As New clsDS2
On Error Resume Next
If txtUser.Text = "" Then Exit Sub
If txtLogon.Text = "" Then
'please dont allow passwords like nothing at all - they should have at least 8 characters or else we dont need this app. anyway
'Note: If some1 is able to crack the ms access database he could easily delete the password field to the start
'the program and enters no password. you can avoid this by not allowing no password!
MsgBox "Username and or password field cant be empty! ", vbInformation, "security..."
txtLogon.SetFocus
Exit Sub
End If

txtLogon.Text = DS2.EncryptString(txtLogon, frmMain.txtKeyD, True)
rs.Update
Setcontrols True
MsgBox "Username and/or password successfully changed.", vbInformation
End Sub

Private Sub Command1_Click()
Dim strpassword As String
Dim DS2 As New clsDS2

On Error GoTo erreveal

strpassword = DS2.DecryptString(txtLogon, frmMain.txtKeyD, True)
MsgBox "Your Password is:   " & vbCrLf & vbCrLf & _
strpassword, vbInformation, "Website Manager..."
Exit Sub

erreveal:
    MsgBox Err.Description
End Sub


Private Sub Form_Load()

  Dim db As Connection
  Set db = New Connection
  
On Error GoTo errhandler
  
db.CursorLocation = adUseClient
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
           App.Path & "\db1.mdb;" & _
           "Jet OLEDB:System database=" & _
           App.Path & "\db1.mdw;", "AkaneTendo", "!:wW39kP19oO"

  Set rs = New Recordset
  rs.Open "select one,two from Table1", db, adOpenStatic, adLockOptimistic

Set txtUser.DataSource = rs
Set txtLogon.DataSource = rs

MakeFlat txtUser.hwnd
MakeFlat txtLogon.hwnd

rs.MoveFirst
Setcontrols True

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

errhandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_
End Sub

Private Sub Setcontrols(bval As Boolean)
cmdEdit.Visible = bval
cmdCancel.Visible = Not bval
cmdUpdate.Enabled = Not bval
txtLogon.Locked = bval
txtUser.Locked = bval
imgExit.Enabled = bval
 End Sub
 


Private Sub imgExit_Click()
On Error Resume Next
rs.Close
Set rs = Nothing
Unload Me
End Sub
