VERSION 5.00
Begin VB.Form FrmDestwise 
   Caption         =   "Destination Wise Report"
   ClientHeight    =   8625
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
   ScaleHeight     =   8625
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDestwiseBg 
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
      Picture         =   "FrmDestwise.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdnxt 
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
         Left            =   3180
         TabIndex        =   3
         Top             =   6120
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
         Left            =   6180
         TabIndex        =   2
         Top             =   6120
         Width           =   1935
      End
      Begin VB.ComboBox cmbDest 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4620
         TabIndex        =   1
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label lblDestwise 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination Wise Report"
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
         Left            =   1740
         TabIndex        =   5
         Top             =   240
         Width           =   7815
      End
      Begin VB.Label lblEntDes 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select the Destination"
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
         Left            =   3300
         TabIndex        =   4
         Top             =   3000
         Width           =   4695
      End
   End
End
Attribute VB_Name = "FrmDestwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecamt As New ADODB.Recordset
Dim totamount As Double
Dim sdate As String
Dim flag As Boolean
Dim flag1 As Boolean
Dim cname As String
Dim ccode As Integer
Dim rdate As String
Dim wtkg As Integer
Dim wtgm As Integer
Dim totwt As Integer
Dim totamt As Double
Dim i As Integer

Private Sub cmdBack_Click()
    Me.Hide
    frmReport.Show
End Sub

Private Sub cmdNxt_Click()
   If cmbDest.Text = "" Then
   MsgBox "Select The Destination"
   End If
   adorecamt.MoveFirst
   
      'Checking from Database location is Correct or Not, If its Correct Storing Data in Variables
      Do While Not adorecamt.EOF
        If RTrim(adorecamt.Fields(3)) = RTrim(cmbDest.Text) Then
            cname = adorecamt.Fields(5)
            ccode = adorecamt.Fields(0)
            rdate = adorecamt.Fields(4)
            wtkg = adorecamt.Fields(6)
            wtgm = adorecamt.Fields(7)
            totamt = adorecamt.Fields(8)
            adorecamt.MoveLast
            adorecamt.MoveNext
        Else
            adorecamt.MoveNext
        End If
        Loop
     
    'Logic for Sending data to Report (drtDestwise)
    drtDestwise.Sections(2).Controls(9).Caption = cname
    drtDestwise.Sections(2).Controls(2).Caption = RTrim(cmbDest.Text)
    drtDestwise.Sections(2).Controls(11).Caption = rdate
    totwt = (wtkg * 1000) + wtgm
    drtDestwise.Sections(2).Controls(12).Caption = totwt
    drtDestwise.Sections(2).Controls(13).Caption = totamt
    drtDestwise.Sections(2).Controls(10).Caption = ccode
   
    Me.Hide
    drtDestwise.Show

End Sub

Private Sub Form_Activate()
    i = 0
    'Logic for Retriving data from Database into Combo box (cmbDest)
    Do While Not adorecamt.EOF
        cmbDest.List(i) = adorecamt.Fields(3)
        adorecamt.MoveNext
        i = i + 1
    Loop
    cmbDest.ListIndex = 0
    cmbDest.SetFocus
    flag = False
    flag1 = False
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordset
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecamt.Open "select *from consignment_details ", adocn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub
