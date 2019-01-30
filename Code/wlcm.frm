VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form wlcm 
   BackColor       =   &H80000014&
   Caption         =   "PHARMACY MANAGEMENT"
   ClientHeight    =   9270
   ClientLeft      =   3120
   ClientTop       =   2340
   ClientWidth     =   15060
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   Picture         =   "wlcm.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   15060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "wlcm.frx":161E8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6000
      Top             =   6120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"wlcm.frx":1A7DD
      OLEDBString     =   $"wlcm.frx":1A868
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton USR 
      BackColor       =   &H80000014&
      Caption         =   "USER LOGIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton ADMN 
      BackColor       =   &H80000014&
      Caption         =   "ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PHARMACY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   26.25
         Charset         =   0
         Weight          =   850
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   10695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ZeroSlasher#1"
      BeginProperty Font 
         Name            =   "Slave only dreams to be king"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   10935
   End
End
Attribute VB_Name = "wlcm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FLAG As Integer
Private Sub ADMN_Click(Index As Integer)
frmadmn.Show
Unload Me
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub USR_Click()
frmuser.Show
Unload Me

End Sub
