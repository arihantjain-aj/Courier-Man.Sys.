VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   9090
   ClientLeft      =   5550
   ClientTop       =   1665
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "High Tower Text"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogBg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   8955
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CheckBox chkShowPass 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   6960
         TabIndex        =   8
         Top             =   5880
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   4
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton cmdLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   3
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         IMEMode         =   3  'DISABLE
         Left            =   4680
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   1
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label lblTagLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Your Trusted Courier Partners"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Label lblPass 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Password  :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Label lblLog 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "User Name :"
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   5
         Top             =   4800
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset

Private Sub chkShowPass_Click()
If chkShowPass.Value = 1 Then
txtPassword.PasswordChar = ""     'Shows Password as Normal Text
ElseIf chkShowPass.Value = 0 Then
txtPassword.PasswordChar = "*"    'Show Password as Star (*)
End If
End Sub

Private Sub cmdExit_Click()
    adorecord.Close
    End
End Sub

Private Sub cmdLog_Click()
  'Checking from Database Login Info. is Correct or Not
  If UCase(adorecord.Fields(0)) = UCase(txtUserName.Text) And UCase(adorecord.Fields(1)) = UCase(txtPassword.Text) Then
        MsgBox "Login Success Welcome"
        adorecord.Close
        Me.Hide
        frmMenu.Show
  Else
        MsgBox "Invalid ...try again"
        txtUserName.SetFocus
        txtUserName.Text = ""
        txtPassword.Text = ""
  End If
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordset
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "login", adocn, adOpenDynamic, adLockOptimistic, adCmdTable
End Sub



