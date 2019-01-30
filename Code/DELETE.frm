VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DELETEMED 
   BackColor       =   &H8000000E&
   Caption         =   "DELETE MEDICINE"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15075
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   11520
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   1200
      Picture         =   "DELETE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "DELETE.frx":3813
      Left            =   2520
      List            =   "DELETE.frx":381D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "SELECT SEARCH CRITERIA"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF00FF&
      Height          =   1305
      Left            =   9480
      MaskColor       =   &H00FF00FF&
      Picture         =   "DELETE.frx":3831
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4095
      Left            =   2520
      TabIndex        =   2
      Top             =   3600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   5640
      TabIndex        =   1
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      Caption         =   "     REMOVE MEDICINES"
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
      Top             =   120
      Width           =   15375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "SEARCH CRITERIA"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "DELETEMED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs4 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

Private Sub Command1_Click()
admnfeat.Show
Unload Me
End Sub

Private Sub Command2_Click()
If text2.Text = "" Then
MsgBox "NO DATA"
Else
rs4.Open "select * from addmed where medid='" & Grid1.TextMatrix(Grid1.Row, 0) & "'", conn, adOpenDynamic, adLockOptimistic
rs4.Delete
rs4.Update
rs4.Close
MsgBox "succesfully DELETED"
Grid1.clear
fill1
text1.Text = clear
End If
text2.Text = ""
End Sub

Private Sub GridHead()
Grid1.Cols = 9
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "MEDID"
Grid1.TextMatrix(0, 1) = "MEDNAME"
Grid1.TextMatrix(0, 2) = "BNO"
Grid1.TextMatrix(0, 4) = "CAT"
Grid1.TextMatrix(0, 3) = "MFGNAME"
Grid1.TextMatrix(0, 5) = "EXP"
Grid1.TextMatrix(0, 6) = "MFGDATE"
Grid1.TextMatrix(0, 7) = "BUYRS"
Grid1.TextMatrix(0, 8) = "SELLRS"
End Sub

Private Sub fill()
Grid1.clear
GridHead
If Combo1.Text = "medid" Then
rs.Open "select * from addmed where medid like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
ElseIf Combo1.Text = "medname" Then
rs.Open "select * from addmed where medname like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
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
        If Not IsNull(rs!Exp) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!Exp
        If Not IsNull(rs!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!mfgdate
        If Not IsNull(rs!buyrs) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!buyrs
        If Not IsNull(rs!sellrs) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!sellrs
    rs.MoveNext
    Loop
End If
rs.Close
End Sub

Private Sub Form_Load()
'main
GridHead
fill1
End Sub



Private Sub Grid1_Click()
text2.Text = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub Text1_Change()
fill
End Sub
Private Sub fill1()
rs.Open "select * from addmed", conn, adOpenDynamic, adLockOptimistic
Grid1.clear
GridHead
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
        If Not IsNull(rs!Exp) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!Exp
        If Not IsNull(rs!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!mfgdate
        If Not IsNull(rs!buyrs) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!buyrs
         If Not IsNull(rs!sellrs) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!sellrs
    rs.MoveNext
    Loop
End If
rs.Close

End Sub
