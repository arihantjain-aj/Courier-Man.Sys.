VERSION 5.00
Begin VB.Form frmCCode 
   Caption         =   "Select Consignor Code"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCCodeBg 
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
      Picture         =   "frmCCode.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   0
      Width           =   14295
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
         TabIndex        =   3
         Top             =   6960
         Width           =   1815
      End
      Begin VB.ComboBox cmbCcode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8880
         TabIndex        =   2
         Top             =   4200
         Width           =   2295
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
         Left            =   7320
         TabIndex        =   1
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label lblSelCCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select The Consignor Code "
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
         Index           =   0
         Left            =   7080
         TabIndex        =   4
         Top             =   3000
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmCCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim i As Integer
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset

Private Sub cmdBack_Click()
    Me.Hide
    frmSource.Show
End Sub

Private Sub cmdNxt_Click()
    adorecord.MoveFirst
    
'Sending Data to Form frmCne
Do While Not adorecord.EOF
   If RTrim(adorecord.Fields(0)) = RTrim(UCase(cmbCcode.List(cmbCcode.ListIndex))) Then
        frmCne.txtCcode = adorecord.Fields(0)
        frmCne.txtName = adorecord.Fields(1)
        frmCne.txtAdd = adorecord.Fields(2)
        frmCne.txtCity = adorecord.Fields(3)
        adorecord.MoveLast
        adorecord.MoveNext
    Else
        adorecord.MoveNext
    End If
Loop

    Me.Hide
    frmCne.Show
    frmCne.txtConame.SetFocus
End Sub

Private Sub Form_Activate()
    adorecord.Close
    adorecord.Open
'Checking if Data is Available or not
'If not
    If adorecord.EOF = True Then
        MsgBox " There is no data in master table"
        Me.Hide
        frmIrccode.Show
    Else
    i = 0
    
'If Available ,Retriving data from Database into Combo box (cmbcode)
    Do While Not adorecord.EOF
        cmbCcode.List(i) = adorecord.Fields(0)
        adorecord.MoveNext
        i = i + 1
    Loop
    cmbCcode.ListIndex = 0
    cmbCcode.SetFocus
    adorecord.MoveFirst
    End If
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordset
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "select *from consignor_info order by consignor_code", adocn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub cmbCcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            cmdNxt.SetFocus
        Case Else
            KeyAscii = 0
    End Select
End Sub
