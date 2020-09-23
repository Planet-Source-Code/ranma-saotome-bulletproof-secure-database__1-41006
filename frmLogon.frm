VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogon.frx":1272
   ScaleHeight     =   2190
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdummy 
      ForeColor       =   &H00800000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "#"
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4110
      TabIndex        =   5
      Top             =   0
      Width           =   4110
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password required"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmLogon.frx":5306
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Enter"
      Height          =   320
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtLogon 
      DataField       =   "two"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "one"
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Administrator"
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim blogon As Boolean
Dim strsecret As String

Private Sub cmdCancel_Click()
blogon = False
Unload Me
End
End Sub

Private Sub cmdOK_Click()
'we connect to the class
Dim DS2 As New clsDS2
On Error GoTo errhandler

'I currently dont allow empty passwords at all --> Security
If txtdummy.Text = "" Then Exit Sub

'OK here we go:
'first we take the encrypted password and decrypt it and store it in a variable "strsecret"
'the we take the password given by the user inside the txtdummy and compare it with the variable
'if they match = blogon = true access granted
'else error
strsecret = DS2.DecryptString(txtLogon, frmMain.txtKeyD, True)
If txtdummy.Text = strsecret Then
blogon = True
frmMain.Show
Unload Me
Else
blogon = False
MsgBox "The password you entered is wrong.", vbInformation, "Logon error"
txtdummy.SetFocus
txtdummy.SelStart = 0
txtdummy.SelLength = Len(txtdummy.Text)
End If
Exit Sub

errhandler:
MsgBox Err.Description
End Sub

Private Sub Form_Load()

  Dim db As Connection
  Set db = New Connection
  
On Error GoTo errhandler
'Here we are connecting to our ms access 2000 database ADO
db.CursorLocation = adUseClient
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
           App.Path & "\db1.mdb;" & _
           "Jet OLEDB:System database=" & _
           App.Path & "\db1.mdw;", "AkaneTendo", "!:wW39kP19oO"

  Set rs = New Recordset
  rs.Open "select one,two from Table1", db, adOpenStatic, adLockReadOnly

Set txtUser.DataSource = rs
Set txtLogon.DataSource = rs

'design
MakeFlat txtUser.hwnd
MakeFlat txtLogon.hwnd
MakeFlat txtdummy.hwnd
txtUser.Locked = True


rs.Close
Set rs = Nothing

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

errhandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_
End Sub
Private Sub txtdummy_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
KeyAscii = 0
cmdOK_Click
End If
End Sub
