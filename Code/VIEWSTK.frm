VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VIEWSTK 
   BackColor       =   &H8000000E&
   Caption         =   "VIEW STOCK"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14775
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "VIEWSTK.frx":0000
      Left            =   3840
      List            =   "VIEWSTK.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "SELECT SEARCH CRITERIA"
      Top             =   6960
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   6960
      TabIndex        =   2
      Top             =   6720
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   240
      MaskColor       =   &H8000000E&
      Picture         =   "VIEWSTK.frx":0032
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   2520
      TabIndex        =   0
      Top             =   2040
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      BackColorBkg    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "STOCK DETAILS"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   6
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "SELECT SEARCH CRITERIA"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "                VIEW STOCK"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   15135
   End
End
Attribute VB_Name = "VIEWSTK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset





Private Sub Command1_Click()
If usern = "user" Then
USERFEAT.Show
Unload Me
ElseIf usern = "admin" Then
admnfeat.Show
Unload Me
End If
End Sub
Private Sub Form_Load()
'main
GridHead
fill
End Sub
Private Sub GridHead()
Grid1.Cols = 10
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "medid"
Grid1.TextMatrix(0, 1) = "medname"
Grid1.TextMatrix(0, 2) = "bno"
Grid1.TextMatrix(0, 4) = "cat"
Grid1.TextMatrix(0, 3) = "mfgname"
Grid1.TextMatrix(0, 5) = "qty"
Grid1.TextMatrix(0, 6) = "exp"
Grid1.TextMatrix(0, 7) = "mfgdate"
Grid1.TextMatrix(0, 8) = "buyrs"
Grid1.TextMatrix(0, 9) = "sellrs"
End Sub
Private Sub fill()
rs.Open "select * from addmed", conn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Grid1.Row = 0
    Do While Not rs.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
            Grid1.Rows = Grid1.Rows + 1
        End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rs!medid) Then Grid1.TextMatrix(Grid1.Row, 0) = rs!medid
        If Not IsNull(rs!medname) Then Grid1.TextMatrix(Grid1.Row, 1) = rs!medname
        If Not IsNull(rs!bno) Then Grid1.TextMatrix(Grid1.Row, 2) = rs!bno
        If Not IsNull(rs!cat) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!cat
        If Not IsNull(rs!mfgname) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!mfgname
        If Not IsNull(rs!QTY) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!QTY
        If Not IsNull(rs!Exp) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!Exp
        If Not IsNull(rs!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!mfgdate
        If Not IsNull(rs!buyrs) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!buyrs
        If Not IsNull(rs!sellrs) Then Grid1.TextMatrix(Grid1.Row, 9) = rs!sellrs
    rs.MoveNext
    Loop
End If
rs.Close
End Sub


Private Sub Text1_Change()
If Combo1.Text = "" Then
MsgBox "Invalid search criteria"
Else
fill1
End If
End Sub
Private Sub fill1()
Grid1.clear
GridHead
If Combo1.Text = "medid" Then
rs.Open "select * from addmed where medid like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
ElseIf Combo1.Text = "medname" Then
rs.Open "select * from addmed where medname like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
ElseIf Combo1.Text = "cat" Then
rs.Open "select * from addmed where cat like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
ElseIf Combo1.Text = "mfgname" Then
rs.Open "select * from addmed where mfgname like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
ElseIf Combo1.Text = "exp" Then
rs.Open "select * from addmed where exp like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
End If
If Not rs.EOF Then
    Grid1.Row = 0
    Do While Not rs.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
            Grid1.Rows = Grid1.Rows + 1
        End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rs!medid) Then Grid1.TextMatrix(Grid1.Row, 0) = rs!medid
        If Not IsNull(rs!medname) Then Grid1.TextMatrix(Grid1.Row, 1) = rs!medname
        If Not IsNull(rs!bno) Then Grid1.TextMatrix(Grid1.Row, 2) = rs!bno
        If Not IsNull(rs!cat) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!cat
        If Not IsNull(rs!mfgname) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!mfgname
        If Not IsNull(rs!QTY) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!QTY
        If Not IsNull(rs!Exp) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!Exp
        If Not IsNull(rs!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!mfgdate
        If Not IsNull(rs!buyrs) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!buyrs
        If Not IsNull(rs!sellrs) Then Grid1.TextMatrix(Grid1.Row, 9) = rs!sellrs
    rs.MoveNext
    Loop
End If
rs.Close

End Sub
