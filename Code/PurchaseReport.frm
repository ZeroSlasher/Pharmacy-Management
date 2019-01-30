VERSION 5.00
Begin VB.Form PurchaseReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Report"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PurchaseReport.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   14760
      TabIndex        =   0
      Top             =   780
      Width           =   735
   End
End
Attribute VB_Name = "PurchaseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
