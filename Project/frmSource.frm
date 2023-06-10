VERSION 5.00
Begin VB.Form frmSource 
   Caption         =   "Source"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
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
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSourceBg 
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
      Picture         =   "frmSource.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.OptionButton optCnote 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignment Note"
         Height          =   855
         Left            =   5880
         TabIndex        =   5
         Top             =   2760
         Value           =   -1  'True
         Width           =   5175
      End
      Begin VB.OptionButton optManifest 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Manifest"
         Height          =   615
         Left            =   5880
         TabIndex        =   3
         Top             =   4200
         Width           =   3855
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
         Left            =   7080
         TabIndex        =   2
         Top             =   6720
         Width           =   1695
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
         TabIndex        =   1
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label lblSource 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Courier System As Source"
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
         Left            =   1838
         TabIndex        =   4
         Top             =   240
         Width           =   7635
      End
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Me.Hide
    frmMenu.Show
End Sub

Private Sub cmdNxt_Click()
    If optCnote.Value = True Then
        frmCnote.Show
    ElseIf optManifest.Value = True Then
        frmManifest.Show
    End If
    Me.Hide
End Sub

Private Sub Form_Activate()
    optCnote.Value = True   'On Form Activation option Consignment Note is selected by Default
    cmdnxt.SetFocus         'Set Focus on Command Next
End Sub
