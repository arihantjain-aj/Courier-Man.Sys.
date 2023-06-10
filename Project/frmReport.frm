VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   8535
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
   ScaleHeight     =   8535
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picReportBg 
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
      Picture         =   "frmReport.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdNxt 
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
         Left            =   4200
         MaskColor       =   &H80000018&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6720
         Width           =   1335
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
         Left            =   7320
         TabIndex        =   5
         Top             =   6720
         Width           =   1815
      End
      Begin VB.OptionButton optInvoice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice"
         Height          =   495
         Left            =   6600
         TabIndex        =   3
         Top             =   2160
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDestwise 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination Wise"
         Height          =   495
         Left            =   6600
         TabIndex        =   2
         Top             =   3480
         Width           =   3615
      End
      Begin VB.OptionButton optSummary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Summary"
         Height          =   495
         Left            =   6600
         TabIndex        =   1
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label lblReport 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report"
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
         Left            =   4687
         TabIndex        =   4
         Top             =   240
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Me.Hide
    frmMenu.Show
End Sub

Private Sub cmdNxt_Click()
 Me.Hide
    If optInvoice.Value = True Then
        frmInvoice.Show
    ElseIf optDestwise.Value = True Then
        FrmDestwise.Show
    Else
        drtSummary.Show
    End If
End Sub

Private Sub Form_Activate()
    optInvoice.Value = True 'On Form Activation option Invoice is selected by Default
End Sub
