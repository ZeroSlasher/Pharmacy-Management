VERSION 5.00
Begin VB.Form admnfeat 
   BackColor       =   &H8000000E&
   Caption         =   "LOGGED IN AS ADMIN"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   FillColor       =   &H00C0C000&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "CHANGE ADMIN PASSWORD"
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
      Left            =   3360
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton Command6 
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
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3360
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Left            =   4080
         TabIndex        =   25
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   4080
         TabIndex        =   24
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   4080
         TabIndex        =   23
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Confirm New Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   22
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Current Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   3735
      End
   End
   Begin VB.CommandButton ADDMDCN 
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   600
      MaskColor       =   &H00FFFF00&
      Picture         =   "ADMINFEAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
      Height          =   1335
      Left            =   2400
      Picture         =   "ADMINFEAT.frx":64AB
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   12600
      Picture         =   "ADMINFEAT.frx":D19D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7920
      Top             =   840
   End
   Begin VB.CommandButton ADMNlgout 
      BackColor       =   &H8000000E&
      Height          =   975
      Left            =   13800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ADMINFEAT.frx":10F68
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   1500
      Left            =   7800
      Picture         =   "ADMINFEAT.frx":14635
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Height          =   1455
      Left            =   7800
      Picture         =   "ADMINFEAT.frx":1932B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton MODSTK 
      Height          =   1500
      Left            =   600
      Picture         =   "ADMINFEAT.frx":1FD96
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton MODUSR 
      Height          =   1380
      Left            =   4200
      Picture         =   "ADMINFEAT.frx":269F7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   3255
   End
   Begin VB.CommandButton DLTUSR 
      BackColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   4200
      Picture         =   "ADMINFEAT.frx":2BE32
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FF0000&
      Height          =   1455
      Left            =   5880
      Picture         =   "ADMINFEAT.frx":32C6E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton USRADD 
      Height          =   1500
      Left            =   4200
      Picture         =   "ADMINFEAT.frx":389BF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton CMDVIEWSTK 
      BackColor       =   &H00C0C000&
      Height          =   1335
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ADMINFEAT.frx":3EE4E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   28
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label11 
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
      Height          =   495
      Left            =   600
      TabIndex        =   27
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "SALES"
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
      Left            =   7920
      TabIndex        =   18
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label6 
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
      Left            =   12240
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Left            =   11400
      TabIndex        =   11
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label3 
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
      Left            =   11400
      TabIndex        =   10
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "STAFF"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "MEDICINE"
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
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
End
Attribute VB_Name = "admnfeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub ADDMDCN_Click()
ADDMED.Show
Unload Me
End Sub

Private Sub BILL_Click()
billing.Show
Unload Me
End Sub

Private Sub CMDVIEWSTK_Click()
VIEWSTK.Show
Unload Me
End Sub


Private Sub Command1_Click()
viewusr.Show
Unload Me
End Sub

Private Sub Command2_Click()
viewsle.Show
Unload Me
End Sub

Private Sub Command3_Click()
billing.Show
Unload Me
End Sub

Private Sub Command4_Click()
Frame1.Visible = True
End Sub

Private Sub Command5_Click()
DELETEMED.Show
Unload Me
End Sub

Private Sub DLTUSR_Click()
DELETEUSR.Show
Unload Me
End Sub

Private Sub MODSTK_Click()
EDITSTK.Show
Unload Me
End Sub

Private Sub VIEWUSR_Click()
viewusr.Show
Unload Me
End Sub

Private Sub MODUSR_Click()
editusr.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Time
Label4.Caption = Date
End Sub

Private Sub USRADD_Click()
addusr.Show
Unload Me
End Sub

Private Sub ADMNlgout_Click()
Dim X As Integer
X = MsgBox("Are you sure you want to logout", vbYesNo, "CONFIRMATION")
If X = 6 Then
frmadmn.Show
Unload Me
Else
End If
End Sub


Private Sub Command6_Click()
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
Label3.Alignment = vbCenter
Label4.Alignment = vbCenter
Label12.Caption = un
Frame1.Visible = False
End Sub

