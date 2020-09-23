VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpay 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Payment"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5280
      TabIndex        =   15
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compute"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reload"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Tid 
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox t4 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox t3 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox t2 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
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
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox t1 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Charge:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Payment:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Student Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2430
      Width           =   4710
      _ExtentX        =   8308
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
Attribute VB_Name = "frmpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If T3.Text <> "" And T2.Text <> "" And T4.Text <> "" And Text1.Text <> "" Then
Set rsupdate = New ADODB.Recordset
rsupdate.Open "Update tstudent set payment='" & Text1.Text & "', bal='" & T3.Text & "', charges='" & T4.Text & "'" & _
"where studentid=" & Tid.Text & "", db, 3, 3
MsgBox "Data is Update!", vbInformation
frmmain.Timer1.Enabled = True
Unload Me
Else
MsgBox "Cannot be Null!", vbExclamation
End If
End Sub

Private Sub Command2_Click()
T2.Text = ""
Call Form_Load
End Sub

Private Sub Command3_Click()
T3.Text = Val(T3.Text) - Val(T2.Text)
Text1.Text = Val(T2.Text) + Val(Text1.Text)
If T3.Text <= 0 Then
MsgBox "Invalid Amount!", vbExclamation
T2.Text = ""
Call Form_Load
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Set rs = New ADODB.Recordset
rs.Open "Select * from tstudent where studentid=" & frmmain.Text1.Text & "", db, 3, 3
Tid.Text = rs!studentid
T1.Text = rs!studentnumber
Text1.Text = rs!payment
T3.Text = rs!bal
T4.Text = rs!charges
Set rs = Nothing
End Sub

