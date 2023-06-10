VERSION 5.00
Begin VB.Form frmIrccode 
   Caption         =   "Consignor Code"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIrCodeBg 
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
      Height          =   8535
      Left            =   0
      Picture         =   "frmIrcode.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.TextBox txtCity 
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
         Height          =   495
         Left            =   10440
         MaxLength       =   15
         TabIndex        =   6
         Top             =   5400
         Width           =   1935
      End
      Begin VB.TextBox txtCname 
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
         Height          =   705
         Left            =   10440
         MaxLength       =   35
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3240
         Width           =   2685
      End
      Begin VB.TextBox txtCadd 
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
         Height          =   705
         Left            =   10440
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   4320
         Width           =   2685
      End
      Begin VB.TextBox txtCcode 
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
         Left            =   10440
         MaxLength       =   5
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
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
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
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
         Width           =   1575
      End
      Begin VB.Label lblCrConCode 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignor Entry Form"
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
         Left            =   3870
         TabIndex        =   11
         Top             =   240
         Width           =   6555
      End
      Begin VB.Label lblCity 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "City :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6600
         TabIndex        =   10
         Top             =   5400
         Width           =   1005
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Address :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   6600
         TabIndex        =   9
         Top             =   4440
         Width           =   2625
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Name :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   6600
         TabIndex        =   8
         Top             =   3360
         Width           =   2265
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignor Code :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   6600
         TabIndex        =   7
         Top             =   2280
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmIrccode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset
Dim adorecas As New ADODB.Recordset
Dim flag As Boolean
Dim tcode As String
Dim tl As String
Dim tr As Integer
Private Sub cmdAdd_Click()
   If txtCname.Text = "" Then
                MsgBox "Empty field"
                txtCname.SetFocus
   ElseIf txtCadd.Text = "" Then
                MsgBox "Empty field"
                txtCadd.SetFocus
   ElseIf txtCity.Text = "" Then
                MsgBox "Empty field"
                txtCity.SetFocus
   Else
   If MsgBox("Do you want to add?", vbYesNo) = vbYes Then
   
        'Sending Data to Database
        adorecord.AddNew
        adorecord.Fields(0) = RTrim(txtCcode.Text)
        adorecord.Fields(1) = RTrim(txtCname.Text)
        adorecord.Fields(2) = RTrim(txtCadd.Text)
        adorecord.Fields(3) = RTrim(txtCity.Text)
        adorecord.Update
        
        'Sending Data to Form frmCne
        frmCne.txtCcode.Text = RTrim(txtCcode.Text)
        frmCne.txtName.Text = RTrim(txtCname.Text)
        frmCne.txtAdd.Text = RTrim(txtCadd.Text)
        frmCne.txtCity.Text = RTrim(txtCity.Text)
        Me.Hide
        frmCne.Show
        frmCne.txtConame.SetFocus
    Else
        txtCcode.Text = ""
        txtCname.Text = ""
        txtCadd.Text = ""
        txtCity.Text = "KOTA"
        txtCcode.Text = tcode
        txtCname.SetFocus
    End If
End If
End Sub

Private Sub cmdCancel_Click()
    txtCname.SetFocus
    Me.Hide
    frmSource.Show
End Sub

Private Sub Form_Activate()
     adorecas.Close
     adorecas.Open
     txtCcode.Text = ""
     txtCname.Text = ""
     txtCadd.Text = ""
     txtCity.Text = "KOTA"
     txtCcode.Enabled = False
     txtCname.SetFocus
     tcode = ""
     
    'logic for Consignor Code
    If adorecas.EOF = True Then
        tcode = "1111"
    Else
        adorecas.MoveLast
        tcode = adorecas.Fields(0)
        tl = Left(tcode, 1)
        tr = Right(tcode, Len(tcode) - 1)
        If tr = 9999 Then
            tr = 1
            tl = Asc(tl) + 1
            tl = Chr(tl)
        End If
        tr = tr + 1
        tcode = tl & tr
    End If
     txtCcode = tcode
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordsets
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "Consignor_info", adocn, adOpenDynamic, adLockOptimistic, adCmdTable
    adorecas.Open "select *from Consignor_info order by consignor_code", adocn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          If txtCity.Text = "" Then
                MsgBox "Empty field"
                txtCity.SetFocus
          Else
                cmdAdd.SetFocus
          End If
    ElseIf Not ((KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or KeyAscii = vbKeyBack Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
    End If
End Sub

Private Sub txtCadd_GotFocus()
    If txtCname.Text = "" Then
        MsgBox "Empty field"
        txtCname.SetFocus
    End If
End Sub

Private Sub txtCity_GotFocus()
    If txtCname.Text = "" Then
        MsgBox "Empty field"
        txtCname.SetFocus
    ElseIf txtCadd.Text = "" Then
        MsgBox "Empty field"
        txtCadd.SetFocus
    End If
End Sub

Private Sub txtCname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCadd.SetFocus
    ElseIf Not ((KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = vbKeyBack Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
        KeyAscii = 0
    End If
End Sub

