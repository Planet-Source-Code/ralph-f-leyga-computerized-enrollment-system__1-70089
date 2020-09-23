VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaddpar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Particular"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmaddpar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Particular Information"
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      Begin VB.TextBox t3 
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox t2 
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox t1 
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         ItemData        =   "frmaddpar.frx":0B3A
         Left            =   1200
         List            =   "frmaddpar.frx":0B50
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Amount:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Particular:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Year Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2205
      Width           =   4455
      _ExtentX        =   7858
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
Attribute VB_Name = "frmaddpar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If T1.Text <> "" And T2.Text <> "" And T3.Text <> "" Then
Set rsadd = New ADODB.Recordset
rsadd.Open "Insert into tparticular (yl,particular,payable) values ('" & T1.Text & "','" & T2.Text & "','" & T3.Text & "');", db, 3, 3
MsgBox "Data is Update!", vbInformation
Set rsadd = Nothing
frmparticular.Timer1.Enabled = True
Unload Me
Else
MsgBox "All fields are required!", vbExclamation
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

