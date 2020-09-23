VERSION 5.00
Begin VB.Form frmupdatesy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update School Year"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmupdatesy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmupdatesy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set rsupdate = New ADODB.Recordset
rsupdate.Open "Update tsy set sy='" & Text1.Text & "'" & _
"where synumber=" & Text2.Text & "", db, 3, 3
MsgBox "Save!", vbInformation
Set rsupdate = Nothing
Unload Me
frmsy.Show vbModal
End Sub

