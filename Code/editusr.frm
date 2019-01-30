VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form editusr 
   BackColor       =   &H8000000E&
   Caption         =   "EDIT USERS OF THE SYSTEM"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   6000
      TabIndex        =   27
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   2760
      TabIndex        =   24
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   3120
      Width           =   2400
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   8160
      Picture         =   "editusr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C000&
      Height          =   1095
      Left            =   6120
      Picture         =   "editusr.frx":2C41
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   7920
      Width           =   2400
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   6720
      Width           =   2400
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   6120
      Width           =   2400
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   5520
      Width           =   2400
   End
   Begin VB.TextBox Text9 
      Height          =   735
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   9840
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   4920
      Width           =   2400
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   2400
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   3720
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "editusr.frx":65A3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8880
      TabIndex        =   0
      Text            =   "Select staff name"
      Top             =   2040
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5895
      Left            =   5160
      TabIndex        =   28
      Top             =   2400
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   10398
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      BackColor       =   12632319
      ForeColorSel    =   -2147483630
      BackColorBkg    =   16777215
      GridColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "EDIT USERS OF THE SYSTEM"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   26
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "      EDIT USER INFO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   25
      Top             =   1080
      Width           =   15375
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "SALARY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   7920
      Width           =   2400
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "JOINDATE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   22
      Top             =   7320
      Width           =   2400
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "EMAIL"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   6720
      Width           =   2400
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "CNO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   6120
      Width           =   2400
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   19
      Top             =   5520
      Width           =   2400
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "BLOOD GROUP"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   4920
      Width           =   2400
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   4320
      Width           =   2400
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   3720
      Width           =   2400
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "EMPLOYEE ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   3120
      Width           =   2400
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "STAFF NAME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   2400
   End
End
Attribute VB_Name = "editusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset

Private Sub Combo1_click()
rs2.Open "select * from adduser where stfname = '" & Combo1.Text & "'", conn, adOpenDynamic, adLockOptimistic
text1.Text = rs2!stfname
text2.Text = rs2!empid
Text3.Text = rs2!gender
text4.Text = rs2!dob
Text5.Text = rs2!blood
text6.Text = rs2!address
text7.Text = rs2!cno
Text8.Text = rs2!email
Text11.Text = rs2!joindate
Text10.Text = rs2!salary
rs2.Close
End Sub

Private Sub Command1_Click()
admnfeat.Show
Unload Me
End Sub




Private Sub Command4_Click()
rs3.Open "select * from adduser where stfname='" & text1.Text & "'", conn, adOpenDynamic, adLockOptimistic
rs3!stfname = text1.Text
rs3!empid = text2.Text
rs3!gender = Text3.Text
rs3!dob = text4.Text
rs3!blood = Text5.Text
rs3!address = text6.Text
rs3!cno = text7.Text
rs3!email = Text8.Text
rs3!joindate = Text11.Text
rs3!salary = Text10.Text
rs3.Update
fill
MsgBox "succesfully saved"

text1.Text = ""
text2.Text = ""
Text3.Text = ""
text4.Text = ""
Text5.Text = ""
Combo1.Text = ""
text6.Text = ""
text7.Text = ""
Text8.Text = ""
Text11.Text = ""
Text10.Text = ""
rs3.Close

End Sub

Private Sub Command5_Click()
text1.Text = clear
text2.Text = clear
Text3.Text = clear
text4.Text = clear
Text5.Text = clear
text6.Text = clear
text7.Text = clear
Text8.Text = clear
Text11.Text = clear
Text10.Text = clear
Combo1.Text = clear
End Sub


Private Sub Form_Load()
'main
fill
text1.Enabled = False
RS1.Open "select * from adduser", conn, adOpenDynamic, adLockOptimistic
If Not RS1.EOF Then
While Not RS1.EOF

Combo1.AddItem RS1!stfname
RS1.MoveNext
Wend
End If
End Sub


Private Sub GridHead()
Grid1.clear
Grid1.Cols = 10
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "STFNAME"
Grid1.TextMatrix(0, 1) = "EMPID"
Grid1.TextMatrix(0, 2) = "GENDER"
Grid1.TextMatrix(0, 4) = "DOB"
Grid1.TextMatrix(0, 3) = "BLOOD"
Grid1.TextMatrix(0, 5) = "ADDRESS"
Grid1.TextMatrix(0, 6) = "CNO"
Grid1.TextMatrix(0, 7) = "EMAIL"
Grid1.TextMatrix(0, 8) = "JOINDATE"
Grid1.TextMatrix(0, 9) = "SALARY"
End Sub

Private Sub fill()
GridHead
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



Private Sub Text12_Change()
Grid1.clear
Grid1.Cols = 11
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "stfnme"
Grid1.TextMatrix(0, 1) = "empid"
Grid1.TextMatrix(0, 2) = "gender"
Grid1.TextMatrix(0, 3) = "dob"
Grid1.TextMatrix(0, 4) = "blood"
Grid1.TextMatrix(0, 5) = "address"
Grid1.TextMatrix(0, 6) = "cno"
Grid1.TextMatrix(0, 7) = "email"
Grid1.TextMatrix(0, 8) = "joindate"
Grid1.TextMatrix(0, 9) = "salary"
rs.Open "select * from adduser where stfname like '" & Text12.Text & "%'", conn, adOpenDynamic, adLockOptimistic
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
