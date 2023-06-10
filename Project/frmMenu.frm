VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
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
   ScaleHeight     =   8520
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMenuBg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.OptionButton optSource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Source"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6360
         TabIndex        =   6
         Top             =   2640
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optDest 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6360
         TabIndex        =   5
         Top             =   3720
         Width           =   3015
      End
      Begin VB.CommandButton cmdNxt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   6720
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   495
         Left            =   7080
         TabIndex        =   3
         Top             =   6720
         Width           =   1695
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report"
         Height          =   555
         Left            =   6360
         TabIndex        =   2
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   4740
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNxt_Click()
    If optSource.Value = True Then
        frmSource.Show
    ElseIf optDest.Value = True Then
        frmDest.Show
    ElseIf optReport.Value = True Then
        frmReport.Show
    End If
    Me.Hide
End Sub

Private Sub Form_Activate()
    optSource.Value = True  'On Form Activation option Source is selected by Default
End Sub




