VERSION 5.00
Begin VB.Form frmuser 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USER LOGIN"
   ClientHeight    =   9420
   ClientLeft      =   2235
   ClientTop       =   1200
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000014&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton EXIT1 
      BackColor       =   &H80000014&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton LOGIN1 
      BackColor       =   &H80000014&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox text2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "PASSWORD"
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox text1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Text            =   "USERNAME"
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   5040
      Picture         =   "USERLOGIN.frx":0000
      Top             =   480
      Width           =   3840
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rse As New Recordset
Private Sub EXIT1_Click()
wlcm.Show
Unload Me
End Sub

Private Sub Form_Load()
'main
End Sub

Private Sub LOGIN1_Click()
If text1.Text = "" Then
MsgBox "Enter Username", vbInformation
text1.SetFocus
Exit Sub
End If
If text2.Text = "" Then
MsgBox "Enter password", vbInformation
text2.SetFocus
Exit Sub
End If
 If text1.Text <> "" And text2.Text <> "" Then
 rse.Open "select * from login where type='user' and username='" & Trim(text1.Text) & "' and password='" & Trim(text2.Text) & "'", conn, adOpenDynamic, adLockOptimistic

 If Not rse.EOF Then
  un = rse!UserName
 usern = "user"
USERFEAT.Show
 Unload Me
 Else
 MsgBox "Invalid username or password", vbCritical
 rse.Close
 End If
 End If
End Sub

Private Sub text1_Click()
text1.Text = ""
End Sub

Private Sub text1_GotFocus()
text1.Text = ""
End Sub

Private Sub text2_Click()
text2.Text = ""
End Sub

Private Sub text2_GotFocus()
text2.Text = ""
End Sub
