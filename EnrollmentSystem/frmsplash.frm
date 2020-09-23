VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmsplash.frx":0B3A
   ScaleHeight     =   4455
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   5760
      Picture         =   "frmsplash.frx":6B0B
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by: Ralph F. Leyga"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00800000&
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Unload Me
frmlogin.Show vbModal
End Sub

