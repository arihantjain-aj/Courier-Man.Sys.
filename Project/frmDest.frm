VERSION 5.00
Begin VB.Form frmDest 
   Caption         =   "Destination"
   ClientHeight    =   8580
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
   ScaleHeight     =   8580
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDestBg 
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
      Picture         =   "frmDest.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cmbCcode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "High Tower Text"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4920
         Width           =   2295
      End
      Begin VB.OptionButton optDrs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Delivery Runsheet"
         Height          =   615
         Left            =   5880
         TabIndex        =   4
         Top             =   4200
         Width           =   5055
      End
      Begin VB.OptionButton optEntry 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entry  Form"
         Height          =   495
         Left            =   5880
         TabIndex        =   3
         Top             =   2760
         Value           =   -1  'True
         Width           =   2775
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
         Left            =   4080
         TabIndex        =   2
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
         Left            =   7080
         TabIndex        =   1
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Label lblDest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Courier System As Destination"
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
         Left            =   1110
         TabIndex        =   5
         Top             =   240
         Width           =   9105
      End
   End
End
Attribute VB_Name = "frmDest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorec As New ADODB.Recordset
Dim adorecord As New ADODB.Recordset
Dim ccode As Integer
Dim ddate As String
Dim dpcs As Integer
Dim origin As String
Dim conName As String
Dim conAdd As String
Dim i As Integer

Private Sub cmdBack_Click()
    Me.Hide
    frmMenu.Show
End Sub

Private Sub cmdNxt_Click()
adorecord.Close
adorecord.Open
 If optEntry.Value = True Then
        frmDestconsignment.Show
 Else
     'Checking from Database Con. Code is Correct or Not, If its Correct Storing Data in Variables
     Do While Not adorecord.EOF
     If RTrim(adorecord.Fields(0)) = RTrim(cmbCcode.List(cmbCcode.ListIndex)) Then
           ccode = RTrim(adorecord.Fields(0))
           ddate = adorecord.Fields(1)
           dpcs = adorecord.Fields(2)
           origin = adorecord.Fields(3)
           conName = adorecord.Fields(4)
           conAdd = adorecord.Fields(5)
           adorecord.MoveLast
           adorecord.MoveNext
      Else
           adorecord.MoveNext
      End If
    Loop
    'Logic for Sending data to Report (drtDrs)
    drtDrs.Sections(2).Controls(40).Caption = ccode
    drtDrs.Sections(2).Controls(41).Caption = ddate
    drtDrs.Sections(2).Controls(42).Caption = dpcs
    drtDrs.Sections(2).Controls(43).Caption = origin
    drtDrs.Sections(2).Controls(44).Caption = conName
    drtDrs.Sections(2).Controls(45).Caption = conAdd
    drtDrs.Show
    
    End If
    Me.Hide
End Sub

Private Sub Form_Activate()
    optEntry.Value = True
    adorec.Close
    adorec.Open
    cmbCcode.Locked = True
    adorecord.Close
    adorecord.Open
    'Checking if there is data in Database
    If adorecord.EOF = True Then
        MsgBox " There is no data in master table"
        Me.Hide
        frmDestconsignment.Show
    Else
    i = 0
    'Logic for Retriving data from Database into Combo box (cmbCcode)
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
    adorecord.Open "select *from dest_consignment_details", adocn, adOpenDynamic, adLockOptimistic, adCmdText
    adorec.Open "temp", adocn, adOpenDynamic, adLockPessimistic, adCmdTable
End Sub

Private Sub optDrs_Click()
     If optEntry.Value = True Then
     cmbCcode.Locked = True
     Else
     cmbCcode.Locked = False
     End If
End Sub
