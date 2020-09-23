VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log In"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3795
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      ForeColor       =   &H00800000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Log In"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please type all the textbox below! Enter the Correct Username and correct password!"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmlogin.frx":0B3A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
login
End Sub

Private Sub Command2_Click()
End
End Sub

Public Sub login()
'On Error Resume Next
Set rspassword = New ADODB.Recordset
rspassword.Open "Select * From tuser", db, 3, 3
If rspassword.RecordCount > 0 Then
    Do Until rspassword.EOF
        If Text1.Text = rspassword!UserName And Text2.Text = rspassword!Password Then
            'frmmain.L1.Caption = Text1.Text
            Set rspassword = Nothing
            Unload Me
                frmmain.Show
                        ctr = 0
                Exit Do
                Else
          rspassword.MoveNext
            ctr = ctr + 1

               
        End If
        Loop
        If ctr > 0 Then
         Set rspassword = Nothing
                MsgBox "Invalid!", vbExclamation
                Text1.Text = ""
                Text2.Text = ""
                Text1.SetFocus
        ctr = 0
    End If
        End If
End Sub

