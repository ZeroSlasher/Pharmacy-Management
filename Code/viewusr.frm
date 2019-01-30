VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form viewusr 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   3120
   ClientTop       =   465
   ClientWidth     =   14835
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Search by Staffname"
      Top             =   2280
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "viewusr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4935
      Left            =   12120
      TabIndex        =   1
      Top             =   3240
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
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
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4935
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      TextStyleFixed  =   1
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "                VIEW USERS"
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
      Left            =   -240
      TabIndex        =   5
      Top             =   1200
      Width           =   15135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "USERS OF THE SYSTEM"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   10575
   End
End
Attribute VB_Name = "viewusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
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
gridhead2
fill
fill2
End Sub
Private Sub GridHead()
Grid1.clear
Grid1.Cols = 10
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "stfname"
Grid1.TextMatrix(0, 1) = "empid"
Grid1.TextMatrix(0, 2) = "gender"
Grid1.TextMatrix(0, 4) = "dob"
Grid1.TextMatrix(0, 3) = "blood"
Grid1.TextMatrix(0, 5) = "address"
Grid1.TextMatrix(0, 6) = "cno"
Grid1.TextMatrix(0, 7) = "email"
Grid1.TextMatrix(0, 8) = "joindate"
Grid1.TextMatrix(0, 9) = "salary"
End Sub
Private Sub gridhead2()
Grid2.clear
Grid2.Cols = 5
Grid2.Rows = 3
Grid2.TextMatrix(0, 0) = "username"
Grid2.TextMatrix(0, 1) = "type"
End Sub
Private Sub fill2()
RS1.Open "select * from login", conn, adOpenDynamic, adLockOptimistic
If Not RS1.EOF Then
    Grid2.Row = 0
    Do While Not RS1.EOF
        If Grid2.Row = Grid2.Rows - 1 Then
            Grid2.Rows = Grid2.Rows + 1
        End If
        Grid2.Row = Grid2.Row + 1
        If Not IsNull(RS1!UserName) Then Grid2.TextMatrix(Grid2.Row, 0) = RS1!UserName
        If Not IsNull(RS1!Type) Then Grid2.TextMatrix(Grid2.Row, 1) = RS1!Type
        RS1.MoveNext
        Loop
        End If
End Sub
Private Sub fill()
rs.Open "select * from adduser", conn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Grid1.Row = 0
    Do While Not rs.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
            Grid1.Rows = Grid1.Rows + 1
        End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rs!stfname) Then Grid1.TextMatrix(Grid1.Row, 0) = rs!stfname
        If Not IsNull(rs!empid) Then Grid1.TextMatrix(Grid1.Row, 1) = rs!empid
        If Not IsNull(rs!gender) Then Grid1.TextMatrix(Grid1.Row, 2) = rs!gender
        If Not IsNull(rs!dob) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!dob
        If Not IsNull(rs!blood) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!blood
        If Not IsNull(rs!address) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!address
        If Not IsNull(rs!cno) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!cno
        If Not IsNull(rs!email) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!email
        If Not IsNull(rs!joindate) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!joindate
        If Not IsNull(rs!salary) Then Grid1.TextMatrix(Grid1.Row, 9) = rs!salary
    rs.MoveNext
    Loop
End If
End Sub


Private Sub Text1_Change()
GridHead
rs3.Open "select * from adduser where stfname like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
If Not rs3.EOF Then
    Grid1.Row = 0
    Do While Not rs3.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
            Grid1.Rows = Grid1.Rows + 1
        End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rs3!stfname) Then Grid1.TextMatrix(Grid1.Row, 0) = rs3!stfname
        If Not IsNull(rs3!empid) Then Grid1.TextMatrix(Grid1.Row, 1) = rs3!empid
        If Not IsNull(rs3!gender) Then Grid1.TextMatrix(Grid1.Row, 2) = rs3!gender
        If Not IsNull(rs3!dob) Then Grid1.TextMatrix(Grid1.Row, 4) = rs3!dob
        If Not IsNull(rs3!blood) Then Grid1.TextMatrix(Grid1.Row, 3) = rs3!blood
        If Not IsNull(rs3!address) Then Grid1.TextMatrix(Grid1.Row, 5) = rs3!address
        If Not IsNull(rs3!cno) Then Grid1.TextMatrix(Grid1.Row, 6) = rs3!cno
        If Not IsNull(rs3!email) Then Grid1.TextMatrix(Grid1.Row, 7) = rs3!email
         If Not IsNull(rs3!joindate) Then Grid1.TextMatrix(Grid1.Row, 8) = rs3!joindate
     If Not IsNull(rs3!salary) Then Grid1.TextMatrix(Grid1.Row, 9) = rs3!salary
    rs3.MoveNext
    Loop
End If
rs3.Close

End Sub

Private Sub Text2_Change()
Grid1.clear
Grid1.Cols = 11
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "stfname"
Grid1.TextMatrix(0, 1) = "empid"
Grid1.TextMatrix(0, 2) = "gender"
Grid1.TextMatrix(0, 4) = "dob"
Grid1.TextMatrix(0, 3) = "blood"
Grid1.TextMatrix(0, 5) = "address"
Grid1.TextMatrix(0, 6) = "cno"
Grid1.TextMatrix(0, 7) = "email"
Grid1.TextMatrix(0, 8) = "joindate"
Grid1.TextMatrix(0, 9) = "salary"

rs2.Open "select * from adduser where medname like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
If Not rs2.EOF Then
    Grid1.Row = 0
    Do While Not rs2.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
            Grid1.Rows = Grid1.Rows + 1
        End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rs2!medid) Then Grid1.TextMatrix(Grid1.Row, 0) = rs!medid
        If Not IsNull(rs2!medname) Then Grid1.TextMatrix(Grid1.Row, 1) = rs!medname
        If Not IsNull(rs2!bno) Then Grid1.TextMatrix(Grid1.Row, 2) = rs!bno
        If Not IsNull(rs2!cat) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!cat
        If Not IsNull(rs2!mfgname) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!mfgname
        If Not IsNull(rs2!Exp) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!Exp
        If Not IsNull(rs2!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!mfgdate
        If Not IsNull(rs2!sellrs) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!sellrs
    rs2.MoveNext
    Loop
End If
rs2.Close
End Sub

