VERSION 5.00
Begin VB.Form userfeat 
   BackColor       =   &H8000000E&
   Caption         =   "LOGGED IN AS USER"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14985
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9285
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   12120
      Top             =   6360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "CHANGE USER PASSWORD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   8055
      Begin VB.CommandButton Command3 
         BackColor       =   &H000080FF&
         Caption         =   "CHANGE PASSWORD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3360
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Left            =   3720
         TabIndex        =   13
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   3720
         TabIndex        =   12
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   3720
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "CONFIRM NEW PASSWORD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "NEW PASSWORD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "CURRENT PASSWORD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "USERFEAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   1935
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Picture         =   "USERFEAT.frx":3DCB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CommandButton USRLGOUT 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "USERFEAT.frx":B8F5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton VIEWSTK1 
      BackColor       =   &H8000000E&
      Height          =   1935
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "USERFEAT.frx":EFC2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CommandButton BILL1 
      BackColor       =   &H8000000E&
      Height          =   1935
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "USERFEAT.frx":1518F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Height          =   615
      Left            =   2040
      TabIndex        =   18
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "GREETINGS,"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6720
      TabIndex        =   16
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      TabIndex        =   15
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Logged in as: "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "USERFEAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim x As Integer

Private Sub BILL1_Click()
billing.Show
Unload Me
End Sub

Private Sub Command1_Click()
viewusr.Show
Unload Me
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
rs.Open "select * from login where username='" & un & "'", conn, adOpenDynamic, adLockOptimistic
If rs!Password = text1.Text And text2.Text = Text3.Text Then
rs!Password = text2.Text
rs.Update
MsgBox "Password successfully changed"
Else
MsgBox "Check your password"
End If
rs.Close
text1.Text = ""
text2.Text = ""
Text3.Text = ""
Frame1.Visible = False
End Sub

Private Sub Form_Load()
'main
Frame1.Visible = False
Label9.Caption = un
End Sub

Private Sub USRLGOUT_Click()
x = MsgBox("Are you sure you want to logout", vbYesNo, "CONFIRMATION")
If x = 6 Then
frmuser.Show
Unload Me
End If
End Sub

Private Sub VIEWSTK1_Click()
FLAG = 1
VIEWSTK.Show
Unload Me

End Sub
Private Sub Timer1_Timer()
Label6.Caption = Time
Label7.Caption = Date
End Sub
