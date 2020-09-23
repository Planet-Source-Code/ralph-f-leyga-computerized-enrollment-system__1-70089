VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmadduser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add User Information"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadduser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox T5 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox T4 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox T3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox T2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox T1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Contact #:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Full name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "User name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2910
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   88194
            MinWidth        =   88194
            Text            =   "Enrollment System 2008"
            TextSave        =   "Enrollment System 2008"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If T1.Text <> "" And T2.Text <> "" And T3.Text <> "" And T4.Text <> "" And T5.Text <> "" Then
Set rsadd = New ADODB.Recordset
rsadd.Open "Insert into tuser (username,[password],fullname,contactnumber,address) values ('" & T1.Text & "','" & T2.Text & "','" & T3.Text & "','" & T4.Text & "','" & T5.Text & "');", db, 3, 3
MsgBox "New user is added!", vbInformation
Set rsadd = Nothing
frmuser.Timer1.Enabled = True
Unload Me
Else
MsgBox "All fields are required!", vbExclamation
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
