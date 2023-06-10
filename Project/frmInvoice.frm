VERSION 5.00
Begin VB.Form frmInvoice 
   Caption         =   "Invoice"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInvoiBg 
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
      Picture         =   "frmInvoice.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   0
      Width           =   14295
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
         ItemData        =   "frmInvoice.frx":11E23
         Left            =   9120
         List            =   "frmInvoice.frx":11E25
         TabIndex        =   3
         Top             =   4440
         Width           =   2175
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
         Height          =   510
         Left            =   7320
         TabIndex        =   2
         Top             =   6960
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
         Height          =   510
         Left            =   10800
         TabIndex        =   1
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Label lblSelCCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select the consignor code :"
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
         Left            =   7440
         TabIndex        =   5
         Top             =   3000
         Width           =   5490
      End
      Begin VB.Label lblInvoice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice"
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
         Height          =   735
         Left            =   6053
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset
Dim ADORECTEMP As New ADODB.Recordset
Dim adorecamt As New ADODB.Recordset
Dim tamt As Double
Dim cname As String
Dim cadd As String
Dim cdest As String
Dim rdate As String
Dim wtkg As Integer
Dim wtgm As Integer
Dim totwt As Integer
Dim totamount As Double
Dim chk1 As Boolean
Dim SEQ As Integer
Dim flag As Boolean
Dim flag1 As Boolean
Dim i As Integer
Dim sdate As String

Private Sub cmdBack_Click()
    Me.Hide
    frmReport.Show
End Sub

Private Sub cmdNxt_Click()
    adorecord.MoveFirst
    adorecamt.MoveFirst
    'Checking from Database Con. Code is Correct or Not, If its Correct Storing Data in Variables
    Do While Not adorecord.EOF
        If RTrim(adorecord.Fields(0)) = UCase(RTrim(cmbCcode.Text)) Then
            cname = adorecord.Fields(1)
            cadd = adorecord.Fields(2)
            adorecord.MoveLast
            adorecord.MoveNext
        Else
            adorecord.MoveNext
        End If
    Loop
     'Checking from Database Con. Code is Correct or Not, If its Correct Storing Data in Variables
     Do While Not adorecamt.EOF
         If RTrim(adorecamt.Fields(0)) = UCase(RTrim(cmbCcode.Text)) Then
            cdest = adorecamt.Fields(3)
            rdate = adorecamt.Fields(4)
            wtkg = adorecamt.Fields(6)
            wtgm = adorecamt.Fields(7)
            totamount = adorecamt.Fields(8)
            adorecamt.MoveLast
            adorecamt.MoveNext
        Else
            adorecamt.MoveNext
        End If
    Loop
    
    'Logic for Sending data to Report (drtInvoice)
    drtInvoice.Sections(2).Controls(3).Caption = cname
    drtInvoice.Sections(2).Controls(4).Caption = cadd
    drtInvoice.Sections(2).Controls(12).Caption = UCase(RTrim(cmbCcode.Text))
    drtInvoice.Sections(2).Controls(13).Caption = cdest
    drtInvoice.Sections(2).Controls(15).Caption = rdate
    totwt = (wtkg * 1000) + wtgm
    drtInvoice.Sections(2).Controls(16).Caption = totwt
    drtInvoice.Sections(2).Controls(17).Caption = totamount
    drtInvoice.Sections(2).Controls(18).Caption = UCase(RTrim(cmbCcode.Text))
    Me.Hide
    drtInvoice.Show
End Sub

Private Sub Form_Activate()
        adorecord.MoveFirst
        adorecamt.MoveFirst
        adorecord.Close
        adorecord.Open
        i = 0
        'Logic for Retriving data from Database into Combo box (cmbCcode)
        Do While Not adorecord.EOF
            cmbCcode.List(i) = adorecord.Fields(0)
            adorecord.MoveNext
            i = i + 1
        Loop
        cmbCcode.ListIndex = 0
        adorecord.MoveFirst
        flag = False
        flag1 = False
        cmbCcode.SetFocus
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordset
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "select *from consignor_info", adocn, adOpenDynamic, adLockOptimistic, adCmdText
    adorecamt.Open "select *from consignment_details ", adocn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub
