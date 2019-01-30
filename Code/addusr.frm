VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form addusr 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "addusr.frx":0000
      Left            =   3360
      List            =   "addusr.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   7200
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   8880
      TabIndex        =   32
      Top             =   3960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42686
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3360
      TabIndex        =   31
      Top             =   3960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   42686
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   8640
      Picture         =   "addusr.frx":001B
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C000&
      Height          =   975
      Left            =   6240
      Picture         =   "addusr.frx":2C5C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Height          =   855
      Left            =   480
      Picture         =   "addusr.frx":65BE
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   12120
      Picture         =   "addusr.frx":9DD1
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Height          =   975
      Left            =   12120
      Picture         =   "addusr.frx":CA12
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2040
      Width           =   1515
   End
   Begin VB.TextBox Text11 
      Height          =   435
      Left            =   3360
      TabIndex        =   24
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Height          =   525
      Left            =   3360
      TabIndex        =   23
      Top             =   6000
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "addusr.frx":10374
      Left            =   3360
      List            =   "addusr.frx":1037E
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3480
      Width           =   2400
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   8880
      TabIndex        =   18
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      Caption         =   "      LOGIN INFORMATION"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   5160
      Width           =   15255
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000E&
      Caption         =   "ROLE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "USRNAME"
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
      TabIndex        =   19
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label11 
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
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "JOINING DATE"
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
      Left            =   6000
      TabIndex        =   16
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label9 
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
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label8 
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
      Height          =   615
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      Caption         =   "       STAFF INFORMATION"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   15135
   End
   Begin VB.Label Label6 
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
      Height          =   615
      Left            =   6000
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "CONTACT NO."
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
      Left            =   6000
      TabIndex        =   11
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   6000
      TabIndex        =   10
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "DATE OF BIRTH"
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
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label EMP_ID 
      BackColor       =   &H8000000E&
      Caption         =   "EMP_ID"
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
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "ADD NEW USERS"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "addusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rse1 As New ADODB.Recordset
Dim rse2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If text1.Text = "" Or text4.Text = "" Or Text5.Text = "" Or text6.Text = "" Or text7.Text = "" Or Text9.Text = "" Then
MsgBox "Please fill in the fields"
Else
rse1.Open "select * from adduser", conn, adOpenDynamic, adLockOptimistic
rse1.AddNew
rse1!stfname = text1.Text
rse1!empid = text2.Text
rse1!dob = DTPicker1.Value
rse1!blood = text4.Text
rse1!address = Text5.Text
rse1!cno = text6.Text
rse1!email = text7.Text
rse1!joindate = DTPicker2.Value
rse1!salary = Text9.Text
rse1!gender = Combo1.Text
rse1.Update
rse1.Close
MsgBox "added successfully"
txtclr
End If
End Sub

Private Sub Command2_Click()
txtclr
End Sub

Private Sub Command3_Click()
admnfeat.Show
Unload Me
End Sub

Private Sub Command4_Click()
If Text10.Text = "" Or Text11.Text = "" Then
MsgBox "Please fill in the fields"
Else
rse2.Open "select * from login where Username = '" & Text10.Text & "' and Type = '" & Combo2.Text & "'", conn, adOpenDynamic, adLockOptimistic
If Not rse2.EOF Then
MsgBox ("Please enter a unique Username")
Else
rse2.AddNew
rse2!UserName = Text10.Text
rse2!Password = Text11.Text
rse2!Type = Combo2.Text
rse2.Update

MsgBox "user added successfully"
txtclr2
rse2.Close

End If

End If

End Sub

Private Sub txtclr2()
Text10.Text = clear
Text11.Text = clear

End Sub

Private Sub Command5_Click()
txtclr2
End Sub

Private Sub Form_Load()
'main
text2.Enabled = False
rs.Open "select max(empid) from adduser", conn, adOpenDynamic, adLockOptimistic
If Not IsNull(rs(0)) Then
text2.Text = Val(rs(0) + 1)
Else
text2.Text = 1
End If
End Sub
Private Sub txtclr()
text1.Text = clear
text2.Text = clear
text4.Text = clear
Text5.Text = clear
text6.Text = clear
text7.Text = clear
Text9.Text = clear

DTPicker1.Value = "1 / 1 / 1990"
DTPicker2.Value = "1 / 1 / 1990"
End Sub

Private Sub Label1_Click()

End Sub
