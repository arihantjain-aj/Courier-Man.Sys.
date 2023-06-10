VERSION 5.00
Begin VB.Form frmDestconsignment 
   Caption         =   "Destination Consignment (Entry Form)"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11280
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
   ScaleHeight     =   8640
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEntFrmBg 
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
      Picture         =   "frmDestconsignment.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         Height          =   510
         Left            =   4673
         TabIndex        =   14
         Top             =   6840
         Width           =   1935
      End
      Begin VB.TextBox txtName 
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
         Height          =   975
         Left            =   5880
         MaxLength       =   35
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   975
         Left            =   5880
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtCno 
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
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtPiece 
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
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   5
         Top             =   5040
         Width           =   2055
      End
      Begin VB.TextBox txtOrigin 
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
         Height          =   450
         Left            =   5880
         MaxLength       =   15
         TabIndex        =   4
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Back"
         Height          =   510
         Left            =   7013
         TabIndex        =   3
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
         Height          =   510
         Left            =   2333
         TabIndex        =   2
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Con. Name"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   2880
         Width           =   2955
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Con. Address"
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
         Left            =   1680
         TabIndex        =   12
         Top             =   3960
         Width           =   3345
      End
      Begin VB.Label lblConNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Con. Code"
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
         Left            =   1680
         TabIndex        =   11
         Top             =   1920
         Width           =   2790
      End
      Begin VB.Label lblPeices 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter No. of Peices"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   5040
         Width           =   3240
      End
      Begin VB.Label lblOrigin 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Origin"
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
         Left            =   1680
         TabIndex        =   9
         Top             =   5880
         Width           =   2205
      End
      Begin VB.Label lblEntForm 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignee Entry Form"
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
         Left            =   1973
         TabIndex        =   1
         Top             =   240
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmDestconsignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset
Dim flag As Boolean

Private Sub cmdAdd_Click()
 If MsgBox("Do you want to Save?", vbYesNo) = vbYes Then
        
          Do While Not adorecord.EOF
                If RTrim(adorecord.Fields(0)) = RTrim(txtCno.Text) Then
                MsgBox "Duplicate data....try agian"
                txtCno.SetFocus
                txtCno.Text = ""
                GoTo Reset
                Else
                    adorecord.MoveNext
                End If
            Loop
            
        'Logic for Sending Data to Database
        adorecord.AddNew
        adorecord.Fields(0) = RTrim(txtCno.Text)
        adorecord.Fields(1) = Date
        adorecord.Fields(2) = RTrim(txtPiece.Text)
        adorecord.Fields(3) = RTrim(txtOrigin.Text)
        adorecord.Fields(4) = RTrim(txtName.Text)
        adorecord.Fields(5) = RTrim(txtAddress.Text)
        adorecord.Update
        Me.Hide
        frmDest.Show
 End If
Reset:
        txtCno.Text = ""
        txtName.Text = ""
        txtAddress.Text = ""
        txtPiece.Text = ""
      
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    frmDest.Show
End Sub

Private Sub cmdAdd_GotFocus()
    If txtOrigin.Text = "" Then
        MsgBox "Empty field"
        txtOrigin.SetFocus
    End If
End Sub

Private Sub cmdUpdate_Click()
If MsgBox("Do you want to Update?", vbYesNo) = vbYes Then
          
          Do While Not adorecord.EOF
                If RTrim(adorecord.Fields(0)) = RTrim(txtCno.Text) Then
                MsgBox "Update Success"
                adorecord.Delete
                GoTo Update
                Else
                    adorecord.MoveNext
                End If
            Loop
Update:
        'Logic for Sending Data to Database
        adorecord.AddNew
        adorecord.Fields(0) = RTrim(txtCno.Text)
        adorecord.Fields(1) = Date
        adorecord.Fields(2) = RTrim(txtPiece.Text)
        adorecord.Fields(3) = RTrim(txtOrigin.Text)
        adorecord.Fields(4) = RTrim(txtName.Text)
        adorecord.Fields(5) = RTrim(txtAddress.Text)
        adorecord.Update
End If
        Me.Hide
        frmDest.Show
End Sub

Private Sub Form_Activate()
        txtCno.Text = ""
        txtName.Text = ""
        txtAddress.Text = ""
        txtPiece.Text = ""
        txtOrigin.Text = "Kota"
        txtCno.SetFocus
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordset
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "dest_consignment_details", adocn, adOpenDynamic, adLockOptimistic, adCmdTable
End Sub

Private Sub txtAddress_GotFocus()
    If txtCno.Text = "" Then
        MsgBox " Empty Field"
        txtCno.SetFocus
    ElseIf txtName.Text = "" Then
        MsgBox " Empty Field"
        txtName.SetFocus
    End If
End Sub

 Private Sub txtCno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtName.SetFocus
    ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
            KeyAscii = 0
    End If
End Sub

Private Sub txtName_GotFocus()
    If txtCno.Text = "" Then
        MsgBox "Empty Field"
        txtCno.SetFocus
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtAddress.SetFocus
     ElseIf Not ((KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = vbKeyBack Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
     End If
End Sub

Private Sub txtPiece_GotFocus()
    If txtCno.Text = "" Then
        MsgBox " Empty Field"
        txtCno.SetFocus
    ElseIf txtName.Text = "" Then
        MsgBox " Empty Field"
        txtName.SetFocus
    ElseIf txtAddress.Text = "" Then
        MsgBox " Empty Field"
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtPiece_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtPiece.SetFocus
     ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
     End If
End Sub
