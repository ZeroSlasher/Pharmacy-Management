VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form viewsle 
   BackColor       =   &H8000000E&
   Caption         =   "VIEW SALES INFO"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   14850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   2880
      Width           =   4095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4935
      Left            =   1800
      TabIndex        =   5
      Top             =   3480
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   0
   End
   Begin VB.CommandButton cmdview 
      BackColor       =   &H00E87777&
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "viewsle.frx":0000
      Left            =   1920
      List            =   "viewsle.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   240
      Picture         =   "viewsle.frx":002B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "From"
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42294
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   615
      Left            =   7080
      TabIndex        =   4
      ToolTipText     =   "To"
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42294
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "SALES DETAILS"
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
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "                VIEW SALES INFORMATION"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   15135
   End
End
Attribute VB_Name = "viewsle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rse As New ADODB.Recordset

Private Sub Combo1_click()
If Combo1.Text = "date" Then
DTPicker1.Enabled = True
DTPicker2.Enabled = True
text1.Text = clear
text1.Enabled = False
ElseIf Combo1.Text = "billamount" Or Combo1.Text = "billno" Then
DTPicker1.Value = Now
DTPicker2.Value = Now
DTPicker1.Enabled = False
DTPicker2.Enabled = False
text1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If usern = "user" Then
USERFEAT.Show
Unload Me
ElseIf usern = "admin" Then
admnfeat.Show
Unload Me
End If
End Sub

Public Sub GridHead()
Grid1.Cols = 11
Grid1.Rows = 100
Grid1.TextMatrix(0, 0) = "MEDICINE NAME"
Grid1.TextMatrix(0, 1) = "CATGRY"
Grid1.TextMatrix(0, 2) = "MFG NAME"
Grid1.TextMatrix(0, 3) = "MFG DATE"
Grid1.TextMatrix(0, 4) = "EXP DATE"
Grid1.TextMatrix(0, 5) = "CUST NAME"
Grid1.TextMatrix(0, 6) = "QTY"
Grid1.TextMatrix(0, 7) = "PPU"
Grid1.TextMatrix(0, 8) = "TOTAL"
Grid1.TextMatrix(0, 9) = "BILLNO"
Grid1.TextMatrix(0, 10) = "BILLDATE"
End Sub

Private Sub cmdview_Click()
Grid1.clear
GridHead
    If Combo1.Text = "date" Then
    rse.Open "select * from billing where billdate>='" & DTPicker1.Value & "' and billdate<='" & DTPicker2.Value & "'", conn, adOpenDynamic, adLockOptimistic
    ElseIf Combo1.Text = "billamount" Then
    rse.Open "select * from billing where total>='" & text1.Text & "'", conn, adOpenDynamic, adLockOptimistic
    ElseIf Combo1.Text = "billno" Then
    rse.Open "select * from billing where billno='" & text1.Text & "'", conn, adOpenDynamic, adLockOptimistic
    End If
  If Not rse.EOF Then
    Grid1.Row = 0
        Do While Not rse.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
        Grid1.Rows = Grid1.Rows + 1
  End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rse!medname) Then Grid1.TextMatrix(Grid1.Row, 0) = rse!medname
        If Not IsNull(rse!cat) Then Grid1.TextMatrix(Grid1.Row, 1) = rse!cat
        If Not IsNull(rse!mfgname) Then Grid1.TextMatrix(Grid1.Row, 2) = rse!mfgname
        If Not IsNull(rse!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 3) = rse!mfgdate
        If Not IsNull(rse!expdate) Then Grid1.TextMatrix(Grid1.Row, 4) = rse!expdate
        If Not IsNull(rse!custname) Then Grid1.TextMatrix(Grid1.Row, 5) = rse!custname
        If Not IsNull(rse!QTY) Then Grid1.TextMatrix(Grid1.Row, 6) = rse!QTY
        If Not IsNull(rse!PPU) Then Grid1.TextMatrix(Grid1.Row, 7) = rse!PPU
        If Not IsNull(rse!total) Then Grid1.TextMatrix(Grid1.Row, 8) = rse!total
        If Not IsNull(rse!billno) Then Grid1.TextMatrix(Grid1.Row, 9) = rse!billno
        If Not IsNull(rse!billdate) Then Grid1.TextMatrix(Grid1.Row, 10) = rse!billdate
        rse.MoveNext
        Loop
End If
rse.Close
End Sub




Private Sub Form_Load()
'main
GridHead
DTPicker1.Value = Now
DTPicker2.Value = Now
DTPicker2.MaxDate = Now
DTPicker1.MinDate = "1/1/2010"
DTPicker1.Enabled = False
DTPicker2.Enabled = False
text1.Enabled = False
End Sub


