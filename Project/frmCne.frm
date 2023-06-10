VERSION 5.00
Begin VB.Form frmCne 
   Caption         =   "Cne"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   Picture         =   "frmCne.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   16725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCneBg 
      Height          =   8775
      Left            =   0
      Picture         =   "frmCne.frx":11E23
      ScaleHeight     =   8715
      ScaleWidth      =   16635
      TabIndex        =   0
      Top             =   0
      Width           =   16695
      Begin VB.Frame fraDocnon 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   11160
         TabIndex        =   38
         Top             =   3480
         Width           =   5055
         Begin VB.CheckBox ChkNdoc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NonDocument"
            BeginProperty Font 
               Name            =   "High Tower Text"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   40
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox ChkDoc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Document"
            BeginProperty Font 
               Name            =   "High Tower Text"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   35
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   1290
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   8040
         Width           =   1215
      End
      Begin VB.TextBox txtAdd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtCity 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtWk 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8160
         MaxLength       =   2
         TabIndex        =   16
         ToolTipText     =   "Maximum Weight is 30 Kg "
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         TabIndex        =   15
         Top             =   8040
         Width           =   1575
      End
      Begin VB.TextBox txtWgms 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8160
         MaxLength       =   3
         TabIndex        =   14
         Top             =   5880
         Width           =   855
      End
      Begin VB.TextBox txtNpa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   14280
         MaxLength       =   2
         TabIndex        =   13
         Top             =   6600
         Width           =   735
      End
      Begin VB.Frame frmIns 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   11160
         TabIndex        =   10
         Top             =   4200
         Width           =   2535
         Begin VB.OptionButton optNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "High Tower Text"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   12
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optYes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "High Tower Text"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtCcode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtSchg 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14280
         MaxLength       =   4
         TabIndex        =   8
         Top             =   5880
         Width           =   735
      End
      Begin VB.TextBox txtConame 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14280
         MaxLength       =   35
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbDest 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         ItemData        =   "frmCne.frx":23C46
         Left            =   14280
         List            =   "frmCne.frx":23C53
         TabIndex        =   6
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cmbTm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         ItemData        =   "frmCne.frx":23C72
         Left            =   14280
         List            =   "frmCne.frx":23C7F
         TabIndex        =   5
         Top             =   5160
         Width           =   2055
      End
      Begin VB.TextBox txtDest 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   14280
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1275
         Width           =   2175
      End
      Begin VB.TextBox txtChg 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   3
         Top             =   6600
         Width           =   855
      End
      Begin VB.TextBox txtTax 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   2
         Top             =   7320
         Width           =   855
      End
      Begin VB.TextBox txtTot 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14280
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   7320
         Width           =   1455
      End
      Begin VB.Label lblCCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignor Code :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   36
         Top             =   480
         Width           =   2640
      End
      Begin VB.Label lblCName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignor Name:"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   35
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lblCAdd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignor Add. :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   34
         Top             =   2040
         Width           =   2625
      End
      Begin VB.Label lblSelBrnch 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Main Branch"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10560
         TabIndex        =   33
         Top             =   2040
         Width           =   3000
      End
      Begin VB.Label lblCCity 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consignor City :"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   32
         Top             =   2760
         Width           =   2595
      End
      Begin VB.Label lblConType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Type Of Consignment"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6240
         TabIndex        =   31
         Top             =   3600
         Width           =   4350
      End
      Begin VB.Label lblWeiKg 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weight{Kg}"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6120
         TabIndex        =   30
         Top             =   5145
         Width           =   1950
      End
      Begin VB.Label lblWeiGr 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weight{g} "
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6120
         TabIndex        =   29
         Top             =   5880
         Width           =   1800
      End
      Begin VB.Label lblNoPac 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter The No Of Packets"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9480
         TabIndex        =   28
         Top             =   6720
         Width           =   4005
      End
      Begin VB.Label lblIns 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Is There Any Insurance?"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6240
         TabIndex        =   27
         Top             =   4320
         Width           =   3765
      End
      Begin VB.Label lblConame 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter Consignee Name:"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10560
         TabIndex        =   26
         Top             =   480
         Width           =   3570
      End
      Begin VB.Label lblSpeChar 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter The Special Charge "
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9480
         TabIndex        =   25
         Top             =   5880
         Width           =   4035
      End
      Begin VB.Label lblTransMeth 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Transportation Method"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9480
         TabIndex        =   24
         Top             =   5115
         Width           =   4530
      End
      Begin VB.Label lblEntDes 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter The Destination"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10560
         TabIndex        =   23
         Top             =   1320
         Width           =   3390
      End
      Begin VB.Label lblCharges 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Charges"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6120
         TabIndex        =   22
         Top             =   6600
         Width           =   1260
      End
      Begin VB.Label lblGst 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gst"
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6120
         TabIndex        =   21
         Top             =   7320
         Width           =   555
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Rs."
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   12360
         TabIndex        =   20
         Top             =   7440
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmCne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset
Dim adorecr1 As New ADODB.Recordset
Dim adorecr2 As New ADODB.Recordset
Dim str As String
Dim flag As Boolean
Dim wt As Integer
Dim i As Integer

Private Sub chkDoc_GotFocus()
    If cmbDest.Text = "" Then
        MsgBox "Cant' be left empty field"
        cmbDest.SetFocus
    End If
End Sub

Private Sub ChkNdoc_Click()
    If txtConame.Text = "" Then
        MsgBox "Cant' be left empty field"
        txtConame.SetFocus
    End If
End Sub

Private Sub ChkNdoc_GotFocus()
    If cmbDest.Text = "" Then
         MsgBox "Cant' be left empty field"
         cmbDest.SetFocus
    End If
End Sub

Private Sub cmbDest_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
        Case 13
            ChkDoc.SetFocus
        Case 103, 71
            KeyAscii = 0
            cmbDest.Text = cmbDest.List(0)
        Case 109, 77
            KeyAscii = 0
            cmbDest.Text = cmbDest.List(1)
        Case 111, 79
            KeyAscii = 0
            cmbDest.Text = cmbDest.List(2)
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmbTm_GotFocus()
    If txtWgms.Text = "" Then
        MsgBox "Can't left empty field"
        txtWgms.SetFocus
    End If
End Sub

Private Sub cmbTm_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            txtSchg.SetFocus
        Case 65, 97
            KeyAscii = 0
            cmbTm.Text = cmbTm.List(0)
         Case 67, 99
            KeyAscii = 0
            cmbTm.Text = cmbTm.List(1)
        Case 115, 83
            KeyAscii = 0
            cmbTm.Text = cmbTm.List(2)
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    frmSource.Show
End Sub

Private Sub cmdSave_Click()
    If txtNpa.Text = "" Then
        MsgBox "Can't be left empty field"
        txtNpa.SetFocus
    Else
        If MsgBox("Do you want to save?", vbYesNo) = vbYes Then
        
'Checking from Database Con. Code is Already Exists or not
'If Yes Updating Database
    Do While Not adorecord.EOF
        If RTrim(adorecord.Fields(0)) = RTrim(txtCcode.Text) Then
            adorecord.Delete
            GoTo Update
        Else
            adorecord.MoveNext
        End If
        Loop
Update:

'If No Adding New
'Sending Data to Database
        adorecord.AddNew
        adorecord.Fields(0) = UCase(RTrim(txtCcode.Text))
          If ChkDoc.Value = Checked And ChkNdoc.Value = Unchecked Then
            adorecord.Fields(2) = "D"
            drtPcne.Sections(1).Controls(1).Caption = "DOC."
          ElseIf ChkNdoc.Value = Checked And ChkDoc.Value = Unchecked Then
            adorecord.Fields(2) = "ND"
            drtPcne.Sections(1).Controls(1).Caption = "NON-DOC."
          Else
            adorecord.Fields(2) = "D,ND"
            drtPcne.Sections(1).Controls(1).Caption = "DOC."
            drtPcne.Sections(1).Controls(2).Caption = "& NON-DOC."
          End If
        adorecord.Fields(1) = RTrim(txtCcode.Text) - 1110
        adorecord.Fields(3) = RTrim(txtDest.Text)
        adorecord.Fields(4) = Date
        adorecord.Fields(5) = UCase(txtConame.Text)
        adorecord.Fields(6) = txtWk.Text
        adorecord.Fields(7) = txtWgms.Text
        adorecord.Fields(8) = txtTot.Text
        adorecord.Update
        MsgBox "Data Updated Successfully"
        Me.Hide
        
'Sending data to Report (drtPcne)
'org,dest,date
        drtPcne.Sections(1).Controls(3).Caption = UCase("Kota")
        drtPcne.Sections(1).Controls(4).Caption = UCase(txtDest)
        drtPcne.Sections(1).Controls(5).Caption = Date
'consignor
        drtPcne.Sections(1).Controls(6).Caption = UCase(txtName)
        drtPcne.Sections(1).Controls(7).Caption = UCase(txtAdd)
'consignee
        drtPcne.Sections(1).Controls(8).Caption = UCase(txtConame)
        drtPcne.Sections(1).Controls(9).Caption = UCase(txtDest)
'kgs,gms,pkgs,charge,spchg,stax,total,date
        If txtWk >= 0 Then
           drtPcne.Sections(1).Controls(10).Caption = txtWk
        End If
        If txtWgms >= 0 Then
           drtPcne.Sections(1).Controls(11).Caption = txtWgms
        End If
        drtPcne.Sections(1).Controls(12).Caption = txtNpa
        drtPcne.Sections(1).Controls(13).Caption = txtChg
        drtPcne.Sections(1).Controls(14).Caption = txtSchg
        drtPcne.Sections(1).Controls(15).Caption = txtTax
        drtPcne.Sections(1).Controls(16).Caption = txtTot
        drtPcne.Show
'Reset
    Else
        txtConame.Text = ""
        txtWk.Text = ""
        txtWgms.Text = ""
        txtNpa.Text = ""
        txtSchg.Text = ""
        ChkDoc.Value = Unchecked
        ChkNdoc.Value = Unchecked
        optNo = True
        optYes = False
        cmbDest.ListIndex = 0
        cmbTm.ListIndex = 0
        txtDest.Text = ""
        txtChg = ""
        txtTax = ""
        txtTot = ""
        txtConame.SetFocus
    End If
    End If
End Sub

Private Sub Form_Activate()
        txtConame.Text = ""
        txtWk.Text = ""
        txtWgms.Text = ""
        txtNpa.Text = ""
        txtSchg.Text = ""
        optNo = True
        optYes = False
        ChkDoc.Value = Unchecked
        ChkNdoc.Value = Unchecked
        cmbDest.ListIndex = 0
        cmbTm.ListIndex = 0
        txtDest.Text = ""
        txtChg = ""
        txtTax = ""
        txtTot = ""
        txtCcode.Enabled = False
        txtName.Enabled = False
        txtCity.Enabled = False
        txtAdd.Enabled = False
        txtTot.Enabled = False
        txtTax.Enabled = False
        txtChg.Enabled = False
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordsets
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "consignment_details", adocn, adOpenDynamic, adLockOptimistic, adCmdTable
    adorecr1.Open "select * from rate_table ", adocn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub optNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtWk.SetFocus
    End If
End Sub

Private Sub txtConame_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDest.SetFocus
    ElseIf Not ((KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = vbKeyBack Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub cmbDest_GotFocus()
    If txtConame.Text = "" Then
        MsgBox "Can't be left empty field"
        txtConame.SetFocus
    ElseIf txtDest.Text = "" Then
        MsgBox "Can't be left empty field"
        txtDest.SetFocus
    End If
End Sub

Private Sub txtDest_GotFocus()
    If txtConame.Text = "" Then
        MsgBox "Can't be left empty field"
        txtConame.SetFocus
    End If
End Sub

Private Sub txtDest_KeyPress(KeyAscii As Integer)
          If KeyAscii = 13 Then
            cmbDest.SetFocus
          ElseIf Not ((KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or KeyAscii = vbKeyBack Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
          End If
End Sub

Private Sub txtNpa_GotFocus()
   adorecord.Close
   adorecr1.Close
   adorecord.Open
   adorecr1.Open
   If txtSchg.Text = "" Then
        MsgBox "Can't be left empty field"
        txtSchg.SetFocus
        
'logic for amount
        Dim twt As Integer
        Dim amt As Integer
        Dim tamt As Integer
        Dim tm As String
        Dim c As Integer
        Dim charge As Double
        Dim totamt As Double
        
    ElseIf txtWk.Text > 30 Then
        MsgBox "Maximum Limit is 30 kg "
        txtWk.SetFocus
    ElseIf txtWgms.Text > 900 Then
        MsgBox "Maximum Limit is 900 gms "
        txtWgms.SetFocus
    Else
        amt = 0
        tamt = 0
        c = 0
        wt = (txtWk * 1000) + txtWgms
        
'Transportation Method
        If cmbTm.Text = "Airways" Then
            tm = "A"
        ElseIf cmbTm.Text = "Roadways" Then
            tm = "RO"
        Else
            tm = "RA"
        End If
   Do While Not adorecr1.EOF
    If adorecr1.Fields(0) = tm Then
        
            If cmbDest.Text = "Rajasthan" Then
                 amt = adorecr1.Fields(2)
            ElseIf cmbDest.Text = "OTHERS" Then
                 amt = adorecr1.Fields(4)
            Else
                 amt = adorecr1.Fields(3)
            End If
            
                If (txtWgms.Text <> 0) Then
                    c = txtWk.Text + 1
                Else
                    c = txtWk.Text
                End If
                amt = amt * c
                adorecr1.MoveNext
    Else
         adorecr1.MoveNext
    End If
    Loop
'Logic for Calculating Tax,Charges,TotalAmt
        charge = (amt * 0.05) + txtSchg.Text
        totamt = charge + amt
        txtTax.Text = (amt * 0.05)
        txtChg.Text = amt
        txtTot.Text = totamt
    End If
End Sub

Private Sub txtNpa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSchg_GotFocus()
Dim flag As Boolean
Dim wt As Integer
    If txtWgms.Text = "" Then
        MsgBox "Can't be left empty field"
        txtWgms.SetFocus
    
    End If
End Sub

Private Sub txtSchg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNpa.SetFocus
    ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWgms_GotFocus()
    If txtWk.Text = "" Then
        MsgBox "Can't be left empty field"
        txtWk.SetFocus
    End If
End Sub

Private Sub txtWgms_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        cmbTm.SetFocus
     ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtWk_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtWgms.SetFocus
     ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
     End If
End Sub

Private Sub txtYes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtWk.SetFocus
    ElseIf (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Then
        KeyAscii = 0
    End If
End Sub
