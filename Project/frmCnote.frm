VERSION 5.00
Begin VB.Form frmCnote 
   Caption         =   "CNote"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   14310
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCnoteBg 
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      Picture         =   "frmCnote.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.OptionButton optYes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   4
         Top             =   4200
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   3
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Back"
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
         Left            =   10800
         TabIndex        =   2
         Top             =   6960
         Width           =   1815
      End
      Begin VB.CommandButton cmdnxt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next"
         Default         =   -1  'True
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
         Left            =   7320
         TabIndex        =   1
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label lblCnote 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creation Of Consignment Note"
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
         Left            =   2595
         TabIndex        =   6
         Top             =   240
         Width           =   9120
      End
      Begin VB.Label lblQuest 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Is There Any Consignor Code  ?"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6600
         TabIndex        =   5
         Top             =   3000
         Width           =   6870
      End
   End
End
Attribute VB_Name = "frmCnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Me.Hide
    frmSource.Show
End Sub

Private Sub cmdNxt_Click()
  Me.Hide
    If optYes.Value = True Then
        frmCCode.Show
    Else
        frmIrccode.Show
    End If
End Sub

Private Sub Form_Activate()
    optYes.Value = True
End Sub


