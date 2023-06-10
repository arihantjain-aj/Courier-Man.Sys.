VERSION 5.00
Begin VB.Form frmWelscr 
   Caption         =   "Welcome"
   ClientHeight    =   9015
   ClientLeft      =   5610
   ClientTop       =   1620
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "High Tower Text"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picWelBg 
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
      Picture         =   "frmWelscr.frx":0000
      ScaleHeight     =   8955
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   615
         Left            =   4740
         TabIndex        =   2
         Top             =   6360
         Width           =   1695
      End
      Begin VB.CommandButton cmdNxt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next"
         Height          =   615
         Left            =   2580
         TabIndex        =   1
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label lblTagLine 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   5
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Label lblWel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to "
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   3060
         TabIndex        =   4
         Top             =   4680
         Width           =   2895
      End
      Begin VB.Label lblFlashco 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Flash Courier Services"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   1140
         TabIndex        =   3
         Top             =   5400
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmWelscr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  End             'Close the program
End Sub

Private Sub cmdNxt_Click()
  Me.Hide
  frmLogin.Show
End Sub
