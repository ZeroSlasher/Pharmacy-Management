VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DELETEUSR 
   BackColor       =   &H8000000E&
   Caption         =   "DELETE USERS FROM THE SYSTEM"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "Myriad Hebrew"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   12240
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   360
      Picture         =   "DELETEUSR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   1920
      TabIndex        =   4
      Top             =   3600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16711680
      BackColorBkg    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8400
      Picture         =   "DELETEUSR.frx":3813
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "DELETE USER INFO"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "      REMOVE USERS"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   15015
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "SEARCH BY STAFF NAME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
   End
End
Attribute VB_Name = "DELETEUSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs4 As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
admnfeat.Show
Unload Me
End Sub

Private Sub Command2_Click()
If text1.Text = "" Then
MsgBox "NO DATA"
Else
rs4.Open "select * from adduser where stfname='" & Grid1.TextMatrix(Grid1.Row, 0) & "'", conn, adOpenDynamic, adLockOptimistic
rs4.Delete
rs4.Update
rs4.Close
MsgBox "succesfully DELETED"
fill1
text1.Text = clear
End If
text1.Text = ""
End Sub
Private Sub GridHead()
Grid1.Cols = 10
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "stfname"
Grid1.TextMatrix(0, 1) = "empid"
Grid1.TextMatrix(0, 2) = "gender"
Grid1.TextMatrix(0, 3) = "dob"
Grid1.TextMatrix(0, 4) = "blood"
Grid1.TextMatrix(0, 5) = "address"
Grid1.TextMatrix(0, 6) = "cno"
Grid1.TextMatrix(0, 7) = "email"
Grid1.TextMatrix(0, 8) = "joindate"
Grid1.TextMatrix(0, 9) = "salary"
End Sub

Private Sub fill()
Grid1.clear
GridHead
rs.Open "select * from adduser where stfname like '" & text1.Text & "%'", conn, adOpenDynamic, adLockOptimistic
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
        If Not IsNull(rs!dob) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!dob
        If Not IsNull(rs!blood) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!blood
        If Not IsNull(rs!address) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!address
        If Not IsNull(rs!cno) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!cno
        If Not IsNull(rs!email) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!email
        If Not IsNull(rs!joindate) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!joindate
        If Not IsNull(rs!salary) Then Grid1.TextMatrix(Grid1.Row, 9) = rs!salary
    rs.MoveNext
    Loop
End If
rs.Close

End Sub





Private Sub Form_Load()
'main
GridHead
fill
fill1
End Sub


Private Sub Grid1_Click()
text1.Text = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub Text1_Change()
fill
End Sub
Private Sub fill1()
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
        If Not IsNull(rs!dob) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!dob
        If Not IsNull(rs!blood) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!blood
        If Not IsNull(rs!address) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!address
        If Not IsNull(rs!cno) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!cno
        If Not IsNull(rs!email) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!email
        If Not IsNull(rs!joindate) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!joindate
        If Not IsNull(rs!salary) Then Grid1.TextMatrix(Grid1.Row, 9) = rs!salary
    rs.MoveNext
    Loop
End If
rs.Close

End Sub

