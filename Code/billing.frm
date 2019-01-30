VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form billing 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PHARMACY BILLING"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "PHARMACY BILL"
      Height          =   7935
      Left            =   1440
      TabIndex        =   33
      Top             =   720
      Width           =   11535
      Begin VB.TextBox Text14 
         Height          =   735
         Left            =   3840
         TabIndex        =   40
         Top             =   6960
         Width           =   2655
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   8520
         TabIndex        =   39
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   5040
         TabIndex        =   38
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   1560
         TabIndex        =   37
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         Height          =   735
         Left            =   8640
         TabIndex        =   36
         Top             =   6960
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid Grid3 
         Height          =   4935
         Left            =   1440
         TabIndex        =   35
         Top             =   1800
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   8
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin VB.Label Label13 
         Caption         =   "Bill date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   46
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "RideR's Pharmacy"
         BeginProperty Font 
            Name            =   "Hobo Std"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   44
         Top             =   360
         Width           =   9015
      End
      Begin VB.Label Label15 
         Caption         =   "Total amount"
         Height          =   735
         Left            =   2640
         TabIndex        =   43
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Bill no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   42
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
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
      Left            =   2280
      TabIndex        =   31
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "DISCARD BILL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   9840
      TabIndex        =   29
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   9840
      TabIndex        =   26
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "REMOVE ITEM FROM BILL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "ADD ITEM TO BILL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3720
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   3015
      Left            =   4920
      TabIndex        =   22
      ToolTipText     =   "Medicines in cart"
      Top             =   5160
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   2280
      TabIndex        =   21
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42681
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42681
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
      Left            =   2280
      TabIndex        =   17
      Top             =   3240
      Width           =   1695
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
      Left            =   6120
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Picture         =   "billing.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3015
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Available medicines"
      Top             =   5160
      Width           =   4570
      _ExtentX        =   8070
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
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
   Begin VB.CommandButton clear 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      Picture         =   "billing.frx":3813
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton PRW 
      BackColor       =   &H00FFFF80&
      Caption         =   "BILL PREVIEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   2775
   End
   Begin VB.TextBox text7 
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
      Left            =   6120
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox text6 
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
      Left            =   6120
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox text2 
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
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox text1 
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
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox text4 
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
      Left            =   6120
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   45
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "CUSTOMETR NAME"
      Height          =   495
      Left            =   4320
      TabIndex        =   34
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "MEDICINE id"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   32
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "DATE"
      Height          =   375
      Left            =   8280
      TabIndex        =   28
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "BILL NO"
      Height          =   495
      Left            =   8280
      TabIndex        =   27
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "            BILLING"
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
      TabIndex        =   25
      Top             =   1080
      Width           =   15255
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "EXP. DATE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "MFG. DATE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "MFG. NAME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "CATEGORY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "PHARMACY BILLING"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label QTY 
      BackColor       =   &H8000000E&
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label FNLAMT 
      BackColor       =   &H8000000E&
      Caption         =   "FINAL AMOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label PPU 
      BackColor       =   &H8000000E&
      Caption         =   "PRICE/UNIT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label MNANE 
      BackColor       =   &H8000000E&
      Caption         =   "MEDICINE NAME"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rse As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
Dim rs7 As New ADODB.Recordset
Dim rs8 As New ADODB.Recordset
Dim rs9 As New ADODB.Recordset
Dim rs10 As New ADODB.Recordset
Dim rec As New ADODB.Recordset

Private Sub Command2_Click() 'DDELETE FROM TABLE & GRID
rs7.Open "select qty from addmed where medname='" & Grid2.TextMatrix(Grid2.Row, 0) & "'", conn, adOpenDynamic, adLockOptimistic
rs7!QTY = rs7!QTY + Val(Grid2.TextMatrix(Grid2.Row, 6))
rs7.Update
rs7.Close
Grid1.clear
GridHead
fill
rs4.Open "select * from billing where billno='" & Text8.Text & "' and medname='" & Grid2.TextMatrix(Grid2.Row, 0) & "'", conn, adOpenDynamic, adLockOptimistic
rs4.Delete
rs4.Close
Grid2.SelectionMode = flexSelectionByRow
Grid2.RemoveItem (Grid2.Row)
End Sub

Private Sub Command4_Click() 'DISCARD BILL
If Grid2.Visible = True Then
rs8.Open "select * from billing where BILLNO='" & Text8.Text & "'", conn, adOpenDynamic, adLockOptimistic
rs8.Delete
rs8.Close
Grid2.clear
gridhead1
MsgBox "BILL DISCARDED"
Else
MsgBox "BILL NOT FOUND"
End If
text4.Text = ""
text4.Enabled = True
End Sub

Private Sub Command5_Click()
Frame1.Visible = False
txtclr
text4.Text = ""
text4.Enabled = True
Text8.Text = Val(Text8.Text) + 1
Grid2.clear
gridhead1
End Sub

Private Sub Grid1_Click()
Text10.Text = Grid1.TextMatrix(Grid1.Row, 0)
text1.Text = Grid1.TextMatrix(Grid1.Row, 1)
text2.Text = Grid1.TextMatrix(Grid1.Row, 2)
Text3.Text = Grid1.TextMatrix(Grid1.Row, 8)
DTPicker1.Value = Grid1.TextMatrix(Grid1.Row, 6)
DTPicker2.Value = Grid1.TextMatrix(Grid1.Row, 5)
text6.Text = Grid1.TextMatrix(Grid1.Row, 3)
End Sub

Private Sub Form_Load()
'main
Text9.Text = Date
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
rs5.Open "select max(billno) from billing", conn, adOpenDynamic, adLockOptimistic
If Not IsNull(rs5(0)) Then
Text8.Text = Val(rs5(0) + 1)
Else
Text8.Text = 1
End If
fill
Frame1.Visible = False
End Sub

Private Sub GridHead()
Grid1.clear
Grid1.Cols = 9
Grid1.Rows = 3
Grid1.TextMatrix(0, 0) = "MEDID"
Grid1.TextMatrix(0, 1) = "MEDNAME"
Grid1.TextMatrix(0, 2) = "CAT"
Grid1.TextMatrix(0, 3) = "SELLRS"
Grid1.TextMatrix(0, 4) = "QTY"
Grid1.TextMatrix(0, 5) = "EXP"
Grid1.TextMatrix(0, 6) = "MFGDATE"
Grid1.TextMatrix(0, 7) = "BNO"
Grid1.TextMatrix(0, 8) = "MFGNAME"
End Sub

Private Sub fill()
GridHead
rs.Open "select * from addmed where medname like'" & Trim(text1.Text) & "%'", conn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Grid1.Row = 0
    Do While Not rs.EOF
        If Grid1.Row = Grid1.Rows - 1 Then
            Grid1.Rows = Grid1.Rows + 1
        End If
        Grid1.Row = Grid1.Row + 1
        If Not IsNull(rs!medid) Then Grid1.TextMatrix(Grid1.Row, 0) = rs!medid
        If Not IsNull(rs!medname) Then Grid1.TextMatrix(Grid1.Row, 1) = rs!medname
        If Not IsNull(rs!bno) Then Grid1.TextMatrix(Grid1.Row, 7) = rs!bno
        If Not IsNull(rs!cat) Then Grid1.TextMatrix(Grid1.Row, 2) = rs!cat
        If Not IsNull(rs!mfgname) Then Grid1.TextMatrix(Grid1.Row, 8) = rs!mfgname
        If Not IsNull(rs!Exp) Then Grid1.TextMatrix(Grid1.Row, 5) = rs!Exp
        If Not IsNull(rs!mfgdate) Then Grid1.TextMatrix(Grid1.Row, 6) = rs!mfgdate
        If Not IsNull(rs!sellrs) Then Grid1.TextMatrix(Grid1.Row, 3) = rs!sellrs
         If Not IsNull(rs!QTY) Then Grid1.TextMatrix(Grid1.Row, 4) = rs!QTY
    rs.MoveNext
    Loop
End If
rs.Close
End Sub




Private Sub Command3_Click()
If text1.Text = "" Or text2.Text = "" Or Text3.Text = "" Or text4.Text = "" Or Text5.Text = "" Or text6.Text = "" Or text7.Text = "" Then
MsgBox "Please fill in the fields"
Else
RS1.Open "select qty from addmed where medid='" & Text10.Text & "'", conn, adOpenDynamic, adLockOptimistic
If Val(Text5.Text) > RS1!QTY Then
MsgBox "out of stock"
Else
gridhead1
savetotbl
savetogrid
text4.Enabled = False
End If
RS1.Close
End If
txtclr
End Sub

Public Sub gridhead1()
Grid2.clear
Grid2.Cols = 11
Grid2.Rows = 10
Grid2.TextMatrix(0, 0) = "MEDICINE NAME"
Grid2.TextMatrix(0, 1) = "CATGRY"
Grid2.TextMatrix(0, 2) = "MFG NAME"
Grid2.TextMatrix(0, 3) = "MFG DATE"
Grid2.TextMatrix(0, 4) = "EXP DATE"
Grid2.TextMatrix(0, 5) = "CUST NAME"
Grid2.TextMatrix(0, 6) = "QTY"
Grid2.TextMatrix(0, 7) = "PPU"
Grid2.TextMatrix(0, 8) = "TOTAL"
Grid2.TextMatrix(0, 9) = "BILLNO"
Grid2.TextMatrix(0, 10) = "BILLDATE"
End Sub



Public Sub savetogrid()
Grid2.Visible = True
rs10.Open "select * from billing where billno='" & Text8.Text & "'", conn, adOpenDynamic, adLockOptimistic
If Not rs10.EOF Then
    Grid2.Row = 0
    Do While Not rs10.EOF
        If Grid2.Row = Grid2.Rows - 1 Then
            Grid2.Rows = Grid2.Rows + 1
        End If
        Grid2.Row = Grid2.Row + 1
      
      
Grid2.TextMatrix(Grid2.Row, 0) = rs10!medname
Grid2.TextMatrix(Grid2.Row, 1) = rs10!cat
Grid2.TextMatrix(Grid2.Row, 2) = rs10!mfgname
Grid2.TextMatrix(Grid2.Row, 3) = rs10!mfgdate
Grid2.TextMatrix(Grid2.Row, 4) = rs10!expdate
Grid2.TextMatrix(Grid2.Row, 5) = rs10!custname
Grid2.TextMatrix(Grid2.Row, 6) = rs10!QTY
Grid2.TextMatrix(Grid2.Row, 7) = rs10!PPU
Grid2.TextMatrix(Grid2.Row, 8) = rs10!total
Grid2.TextMatrix(Grid2.Row, 9) = rs10!billno
Grid2.TextMatrix(Grid2.Row, 10) = rs10!billdate
    rs10.MoveNext
    Loop
End If
rs10.Close
End Sub

Public Sub savetotbl()
rs2.Open "select * from billing where billno = '" & Text8.Text & "' And medname = '" & text1.Text & "'", conn, adOpenDynamic, adLockOptimistic ' where billno='" & Text8.Text & "' and medname='" & Text1.Text & "'", conn, adOpenDynamic, adLockOptimistic
        
        If Not rs2.EOF Then
rs2!QTY = rs2!QTY + Val(Text5.Text)
rs2!total = rs2!total + Val(text7.Text)
    rs2.Update
    rs2.Close
Else
rs2.AddNew
rs2!medname = text1.Text
rs2!cat = text2.Text
rs2!mfgname = Text3.Text
rs2!mfgdate = DTPicker1.Value
rs2!expdate = DTPicker2.Value
rs2!custname = text4.Text
rs2!QTY = Val(Text5.Text)
rs2!PPU = Val(text6.Text)
rs2!total = Val(text7.Text)
rs2!billno = Text8.Text
rs2!billdate = Text9.Text
rs2.Update
rs2.Close
End If
rs3.Open "select qty from addmed where medname='" & text1.Text & "'", conn, adOpenDynamic, adLockOptimistic
rs3!QTY = rs3!QTY - Val(Text5.Text)
rs3.Update
rs3.Close

End Sub
Private Sub txtclr()
text1.Text = ""
text2.Text = ""
Text3.Text = ""
Text5.Text = ""
text6.Text = ""
text7.Text = ""
Text10.Text = ""
DTPicker1.Value = "1 / 1 / 2010"
DTPicker2.Value = "1 / 1 / 2010"
End Sub

Private Sub clear_Click()
txtclr
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




Private Sub PRW_Click()
If Grid2.Visible = False Then
MsgBox ("BILL NOT FOUND")
Else
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Frame1.Visible = True
GRIDHD
rec.Open "select * from billing where billno='" & Text8.Text & "'", conn, adOpenDynamic, adLockOptimistic
If Not rec.EOF Then
    Grid3.Row = 0
    Do While Not rec.EOF
        If Grid3.Row = Grid3.Rows - 1 Then
            Grid3.Rows = Grid3.Rows + 1
        End If
        Grid3.Row = Grid3.Row + 1
Grid3.TextMatrix(Grid3.Row, 0) = rec!medname
Grid3.TextMatrix(Grid3.Row, 1) = rec!cat
Grid3.TextMatrix(Grid3.Row, 2) = rec!mfgname
Grid3.TextMatrix(Grid3.Row, 3) = rec!mfgdate
Grid3.TextMatrix(Grid3.Row, 4) = rec!expdate
Text11.Text = rec!custname
Grid3.TextMatrix(Grid3.Row, 5) = rec!QTY
Grid3.TextMatrix(Grid3.Row, 6) = rec!PPU
Grid3.TextMatrix(Grid3.Row, 7) = rec!total
Text13.Text = rec!billno
Text12.Text = rec!billdate
    rec.MoveNext
    Loop
End If
For intRow = 0 To Grid3.Row
inttotal = inttotal + Val(Grid3.TextMatrix(intRow, 7))
Next intRow
Text14.Text = inttotal
rec.Close
End If

End Sub

Private Sub Text5_Change()
Dim a As Integer
Dim b As Integer
Dim c As Integer
a = Val(Text5.Text)
b = Val(text6.Text)
c = a * b
text7.Text = c
End Sub

Private Sub Text1_Change()
fill
End Sub
Public Sub GRIDHD()
Grid3.clear
Grid3.Cols = 8
Grid3.Rows = 10
Grid3.TextMatrix(0, 0) = "MEDICINE NAME"
Grid3.TextMatrix(0, 1) = "CATGRY"
Grid3.TextMatrix(0, 2) = "MFG NAME"
Grid3.TextMatrix(0, 3) = "MFG DATE"
Grid3.ColWidth(3) = 1200
Grid3.TextMatrix(0, 4) = "EXP DATE"
Grid3.ColWidth(4) = 1200
Grid3.TextMatrix(0, 5) = "QTY"
Grid3.TextMatrix(0, 6) = "PPU"
Grid3.TextMatrix(0, 7) = "TOTAL"
text4.Enabled = True
End Sub
