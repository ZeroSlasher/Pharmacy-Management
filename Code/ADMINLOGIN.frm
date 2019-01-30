VERSION 5.00
Begin VB.Form frmadmn 
   BackColor       =   &H80000014&
   Caption         =   "ADMINISTRATOR LOGIN"
   ClientHeight    =   9255
   ClientLeft      =   2850
   ClientTop       =   1005
   ClientWidth     =   15090
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000014&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton exit 
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
      Left            =   6840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton LOGIN 
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
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
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "PASSWORD"
      Top             =   5040
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
      Left            =   6000
      TabIndex        =   0
      Text            =   "USERNAME"
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   3450
      Left            =   6000
      Picture         =   "ADMINLOGIN.frx":0000
      Top             =   600
      Width           =   3285
   End
End
Attribute VB_Name = "frmadmn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rse As New Recordset
Private Sub EXIT_Click() 'EXIT BUTTON
wlcm.Show
Unload Me
End Sub

Private Sub Form_Load()
'main 'MODULE CONNECTION
End Sub

Private Sub LOGIN_Click() 'LOGIN BUTTON
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
 rse.Open "select * from login where type='admin' and username='" & Trim(text1.Text) & "' and password='" & Trim(text2.Text) & "'", conn, adOpenDynamic, adLockOptimistic
 If Not rse.EOF Then
 un = rse!UserName
 usern = "admin"
admnfeat.Show
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

Private Sub text2_Click()
text2.Text = ""
End Sub

Private Sub text1_GotFocus()
text1.Text = ""
End Sub

Private Sub text2_GotFocus()
text2.Text = ""
End Sub
