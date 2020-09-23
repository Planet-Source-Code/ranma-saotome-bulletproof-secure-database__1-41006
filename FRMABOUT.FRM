VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sumari Arts"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":1272
   ScaleHeight     =   2625
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Accept"
      Height          =   320
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4440
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Website Manager and Password Keeper"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label lblabout 
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      Height          =   1215
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":5306
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim strinfo As String
strinfo = "Database for favorite Websites and " & vbCrLf & _
"Email Addresses and passwords. " & vbCrLf & _
"This program is freeware!" & vbCrLf & _
"Written in VB6 (c) 2002 Sumari Arts" & vbCrLf & vbCrLf & _
"my email address: aeonflux@tokyo.com"
lblabout.Caption = strinfo
Me.Caption = App.Title & "   Version: " & App.Major & "." & App.Minor & " (Build: " & App.Revision & " )"
End Sub
