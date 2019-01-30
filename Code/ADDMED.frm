VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ADDMED 
   BackColor       =   &H8000000E&
   Caption         =   "ADD MEDICINE TO PHARMACY STOCK"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15075
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15075
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ADDMED.frx":0000
      Left            =   3240
      List            =   "ADDMED.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5520
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   735
      Left            =   10440
      TabIndex        =   22
      Top             =   2520
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   131006467
      CurrentDate     =   42660
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   10440
      TabIndex        =   21
      Top             =   3480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42660
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   10680
      Picture         =   "ADDMED.frx":002A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C000&
      Height          =   975
      Left            =   7680
      Picture         =   "ADDMED.frx":2C6B
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   360
      Picture         =   "ADDMED.frx":65CD
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   735
      Left            =   10440
      TabIndex        =   17
      Top             =   5640
      Width           =   3000
   End
   Begin VB.TextBox Text9 
      Height          =   735
      Left            =   10440
      TabIndex        =   16
      Top             =   4560
      Width           =   3000
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   7320
      Width           =   3000
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   3240
      TabIndex        =   10
      Top             =   6360
      Width           =   3000
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   3240
      TabIndex        =   9
      Top             =   4440
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3240
      TabIndex        =   8
      Top             =   3480
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      Caption         =   "      MEDICINE INFORMATION"
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
      TabIndex        =   23
      Top             =   1440
      Width           =   15135
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
      Height          =   735
      Left            =   7080
      TabIndex        =   15
      Top             =   5520
      Width           =   2535
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
      Height          =   735
      Left            =   7080
      TabIndex        =   14
      Top             =   4560
      Width           =   2535
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
      Height          =   735
      Left            =   7080
      TabIndex        =   13
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "DATE OF MANUFACTURE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   12
      Top             =   2520
      Width           =   2535
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
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   7320
      Width           =   2535
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
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Width           =   2535
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
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   5400
      Width           =   2535
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
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
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
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "ADD NEW MEDICINES"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   21.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "ADDMED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rse As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub Command3_Click()
admnfeat.Show
Unload Me
End Sub

Private Sub Command4_Click()
If text1.Text = "" Or text2.Text = "" Or Text3.Text = "" Or text4.Text = "" Or Text5.Text = "" Or text6.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "Please fill in the fields"
Else
rse.Open "select * from addmed", conn, adOpenDynamic, adLockOptimistic
rse.AddNew
rse!medid = text1.Text
rse!medname = text2.Text
rse!bno = Text3.Text
rse!cat = Combo1.Text
rse!mfgname = Text5.Text
rse!QTY = text6.Text
rse!Exp = DTPicker1.Value
rse!mfgdate = DTPicker2.Value
rse!buyrs = Val(Text9.Text)
rse!sellrs = Val(Text10.Text)
rse.Update
MsgBox "added successfully"
txtclr
text1.Text = Val(text1.Text + 1)
text2.SetFocus
rse.Close
End If
End Sub

Private Sub Command5_Click()
txtclr
End Sub
 
Private Sub Form_Load()
'main
text1.Enabled = False
rs.Open "select max(medid) from addmed", conn, adOpenDynamic, adLockOptimistic
If Not IsNull(rs(0)) Then
text1.Text = Val(rs(0) + 1)
Else
text1.Text = 1
End If
End Sub

Private Sub txtclr()
text2.Text = clear
Text3.Text = clear
text4.Text = clear
Text5.Text = clear
text6.Text = clear
DTPicker1 = "1/1/2010"
DTPicker2 = "1/1/2010"
Text9.Text = clear
Text10.Text = clear
End Sub


