VERSION 5.00
Begin VB.Form frmManifest 
   Caption         =   "Manifest"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11385
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
   Picture         =   "frmManifest.frx":0000
   ScaleHeight     =   8085
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picmanifestBg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      Picture         =   "frmManifest.frx":A8F3
      ScaleHeight     =   7995
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox cmbMdest 
         BackColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   4485
         TabIndex        =   3
         Top             =   4560
         Width           =   2415
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
         Height          =   510
         Left            =   3315
         TabIndex        =   2
         Top             =   6120
         Width           =   1755
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
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
         Left            =   6315
         TabIndex        =   1
         Top             =   6120
         Width           =   1755
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manifest"
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
         Left            =   4365
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblSelDes 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select The Destination"
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
         Left            =   3240
         TabIndex        =   4
         Top             =   3000
         Width           =   4890
      End
   End
End
Attribute VB_Name = "frmManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaraing Variables and its Data types
Dim adocn As New ADODB.Connection
Dim adorecord As New ADODB.Recordset
Dim adorecwt As New ADODB.Recordset
Dim i As Integer
Dim ccode As Integer
Dim twt As Double
Dim count1 As Integer
Dim flag As Boolean

Private Sub cmdBack_Click()
    Me.Hide
    frmSource.Show
End Sub

Private Sub cmdNxt_Click()
        adorecwt.Close
        adorecwt.Open
        twt = 0
        count1 = 0
        
            'Checking from Database Destination is Correct or Not (cmbMdest)
            Do While Not adorecwt.EOF
                If RTrim(adorecwt.Fields(3)) = RTrim(cmbMdest.Text) Then
                    twt = twt + adorecwt.Fields(6) + adorecwt.Fields(7) * 0.001
                    count1 = count1 + 1
                    ccode = RTrim(adorecwt.Fields(0))
                End If
                adorecwt.MoveNext
            Loop
            
            'Sending data to Report (drtManifest)
            drtManifest.Sections(2).Controls(19).Caption = ccode
            drtManifest.Sections(2).Controls(5).Caption = cmbMdest.Text
            drtManifest.Sections(2).Controls(20).Caption = count1
            drtManifest.Sections(5).Controls(4).Caption = twt
            drtManifest.Sections(5).Controls(2).Caption = count1
            cmbMdest.RemoveItem cmbMdest.ListIndex
            If cmbMdest.ListCount > 0 Then
               cmbMdest.ListIndex = 0
            End If
            Me.Hide
            drtManifest.Show
        
End Sub

Private Sub Form_Activate()
    adorecord.Close
    adorecwt.Close
    adorecord.Open
    adorecwt.Open
    i = 0
    
    'Retriving data from Database into Combo box (cmbMdest)
    Do While Not adorecord.EOF
        flag = True
        cmbMdest.List(i) = adorecord.Fields(0)
        adorecord.MoveNext
        i = i + 1
    Loop
    
    'If no Data is Available in Database
    If flag = False Then
        MsgBox " No Appropriate Data Is Available"
        Me.Hide
        frmMenu.Show
    Else
        cmbMdest.ListIndex = 0
    End If
    
    twt = 0
    count1 = 0
    flag = False
End Sub

Private Sub Form_Load()
    'Connecting and Opening ADODB Connection & Recordsets
    adocn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COURIER_DB.MDB;Persist Security Info=False"
    adorecord.Open "select distinct destination from consignment_details ", adocn, adOpenDynamic, adLockOptimistic, adCmdText
    adorecwt.Open "select *from consignment_details ", adocn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub


