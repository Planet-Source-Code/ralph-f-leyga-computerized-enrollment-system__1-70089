VERSION 5.00
Begin VB.Form frmaddsy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add School Year"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmaddsy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmaddsy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set rsadd = New ADODB.Recordset
rsadd.Open "Insert into tsy (sy) values ('" & Text1.Text & "');", db, 3, 3
MsgBox "New School Year is Added!", vbInformation
Unload Me
Set rsadd = Nothing
frmsy.Show vbModal
End Sub
