VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EDITSTK 
   BackColor       =   &H8000000E&
   Caption         =   "MODIFY MEDICINE INFORMATION"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   13755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   2160
      TabIndex        =   26
      Top             =   4200
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   525
      ItemData        =   "EDITSTK.frx":0000
      Left            =   2160
      List            =   "EDITSTK.frx":000D
      TabIndex        =   25
      Text            =   "Combo2"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Height          =   525
      Left            =   6240
      TabIndex        =   24
      ToolTipText     =   "SEARCH BY MEDNAME"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   23
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   22
      Top             =   1320
      Width           =   2300
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42676
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42676
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "EDITSTK.frx":0029
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "EDITSTK.frx":3C6E
      Left            =   9120
      List            =   "EDITSTK.frx":3C70
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "SELECT MEDICINE NAME"
      Top             =   840
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5895
      Left            =   4440
      TabIndex        =   16
      Top             =   1320
      Width           =   13000
      _ExtentX        =   22939
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
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   2300
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3720
      Width           =   2300
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   6120
      Width           =   2300
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   6720
      Width           =   2300
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3000
      MaskColor       =   &H00008000&
      Picture         =   "EDITSTK.frx":3C72
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   120
      Picture         =   "EDITSTK.frx":75D4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "MEDICINE NAME"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   2000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "EDIT MEDICINE INFORMATION"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   15255
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "MEDICINE ID"
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
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   2000
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "BATCH NUMBER"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   2000
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "CATEGORY"
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
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2000
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "MANUFACTURER"
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
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   2000
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "QUANTITY"
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
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   2000
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "PRODUCTION DATE"
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
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   2000
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "EXPIRY DATE"
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
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   2000
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "BUYING PRICE"
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
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   2000
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "SELLING PRICE"
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
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   2000
   End
End
Attribute VB_Name = "EDITSTK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset

Private Sub Combo1_click()
rs2.Open "select * from addmed where medname = '" & Combo1.Text & "'", conn, adOpenDynamic, adLockOptimistic
text1.Text = rs2!medid
text2.Text = rs2!medname
Text3.Text = rs2!bno
Combo2.Text = rs2!cat
Text5.Text = rs2!mfgname
text6.Text = rs2!QTY
DTPicker1.Value = rs2!Exp
DTPicker2.Value = rs2!mfgdate
Text9.Text = rs2!buyrs
Text10.Text = rs2!sellrs
rs2.Close
End Sub



Private Sub Command1_Click()
admnfeat.Show
Unload Me
End Sub





Private Sub Command4_Click()
rs3.Open "select * from addmed where medid='" & text1.Text & "'", conn, adOpenDynamic, adLockOptimistic
If text1.Text = "" Or text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or text6.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "Please fill in the fields"
Else
rs3!medname = text2.Text
rs3!medid = text1.Text
rs3!bno = Text3.Text
rs3!cat = Combo2.Text
rs3!mfgname = Text5.Text
rs3!QTY = text6.Text
rs3!mfgdate = DTPicker2.Value
rs3!Exp = DTPicker1.Value
rs3!buyrs = Text9.Text
rs3!sellrs = Text10.Text
rs3.Update
fill
MsgBox "succesfully saved"
End If
text2.Text = ""
Text3.Text = ""
Text5.Text = ""
'Combo1.Text = ""
Text10.Text = ""
text1.Text = ""
text6.Text = ""
Text9.Text = ""
'Text2.Text = SetFocus
rs3.Close

End Sub

Private Sub Command5_Click()
text1.Text = clear
text2.Text = clear
Text3.Text = clear

Text5.Text = clear
text6.Text = clear
Text9.Text = clear
Text10.Text = clear
'Combo1.Text = clear
DTPicker1.Value = "1/1/2010"
DTPicker2.Value = "1/1/2010"
'Text2.Text = SetFocus
End Sub


Private Sub Form_Load()
'main
fill
text1.Enabled = False
RS1.Open "select * from addmed", conn, adOpenDynamic, adLockOptimistic
If Not RS1.EOF Then
While Not RS1.EOF

Combo1.AddItem RS1!medname
RS1.MoveNext
Wend
End If
End Sub


Private Sub GridHead()
Grid1.clear
Grid1.Cols = 10
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
Grid1.TextMatrix(0, 9) = "qty"
End Sub


Private Sub Grid1_Click()
text1.Text = Grid1.TextMatrix(Grid1.Row, 0)
text2.Text = Grid1.TextMatrix(Grid1.Row, 1)
Text3.Text = Grid1.TextMatrix(Grid1.Row, 2)
Combo2.Text = Grid1.TextMatrix(Grid1.Row, 4)
Text5.Text = Grid1.TextMatrix(Grid1.Row, 3)
text6.Text = Grid1.TextMatrix(Grid1.Row, 9)
Text9.Text = Grid1.TextMatrix(Grid1.Row, 7)
Text10.Text = Grid1.TextMatrix(Grid1.Row, 8)
DTPicker1.Value = Grid1.TextMatrix(Grid1.Row, 5)
DTPicker2.Value = Grid1.TextMatrix(Grid1.Row, 6)
End Sub
Private Sub fill()
GridHead
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
        If Not IsNull(rs!Exp) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!Exp
        If Not IsNull(rs!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!mfgdate
        If Not IsNull(rs!buyrs) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!buyrs
        If Not IsNull(rs!sellrs) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!sellrs
        If Not IsNull(rs!QTY) Then Grid1.TextMatrix(Grid1.Row, 9) = rs!QTY
    rs.MoveNext
    Loop
End If
rs.Close

End Sub

Private Sub Text7_Change()
Grid1.clear
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

rs.Open "select * from addmed where medname like '" & text7.Text & "%'", conn, adOpenDynamic, adLockOptimistic
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
