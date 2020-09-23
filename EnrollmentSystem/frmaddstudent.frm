VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaddstudent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Student"
   ClientHeight    =   4710
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
   Icon            =   "frmaddstudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4455
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox tID 
      Height          =   315
      Left            =   4920
      TabIndex        =   19
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4455
      Begin VB.TextBox T10 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox T9 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox t8 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3480
         Width           =   1815
      End
      Begin VB.ComboBox t7 
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
         ItemData        =   "frmaddstudent.frx":0B3A
         Left            =   1560
         List            =   "frmaddstudent.frx":0B50
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ComboBox t5 
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
         ItemData        =   "frmaddstudent.frx":0B66
         Left            =   1560
         List            =   "frmaddstudent.frx":0B70
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox t6 
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox t4 
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
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
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
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   2775
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
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   2775
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Charges:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade/YL"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmaddstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If T1.Text <> "" And T2.Text <> "" And T3.Text <> "" And T4.Text <> "" And T5.Text <> "" And t6.Text <> "" And t7.Text <> "" And t8.Text <> "" And T9.Text <> "" Then
Set rsadd = New ADODB.Recordset
rsadd.Open "Insert into tstudent (studentnumber,lname,fname,mname,gender,age,yl,sy,charges,bal) values ('" & T1.Text & "','" & T2.Text & "','" & T3.Text & "','" & T4.Text & "','" & T5.Text & "','" & t6.Text & "','" & t7.Text & "','" & t8.Text & "','" & T9.Text & "','" & T10.Text & "');", db, 3, 3
MsgBox T1.Text & " is save!", vbInformation
frmstudent.Timer1.Enabled = True
Unload Me
Else
MsgBox "All fields are required!", vbExclamation
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
rs.Open "Select * from tstudent", db, 3, 3

If rs.RecordCount = 0 Then
        idno = 1
    Else
       rs.MoveLast
        idno = rs.Fields("examineeid") + 1
    End If
  
    T1.Text = Format(Now(), "yyyy-") & Format(idno, "0000000")
    cbosy
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsadd = Nothing
Set rs = Nothing
Set rsSY = Nothing
End Sub
Public Sub cbosy()
Set rsSY = New ADODB.Recordset
rsSY.Open "Select * from tSY order by sy asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsSY.EOF
        t8.AddItem rsSY!sy
        rsSY.MoveNext
    Loop
    Set rsSY = Nothing
    End Sub

Private Sub t7_Click()
On Error GoTo Error
Set rsPay = New ADODB.Recordset
rsPay.Open "Select sum(payable)as sumpay from tParticular where yl='" & t7.Text & "'", db, 3, 3
T9.Text = rsPay!sumpay
T10.Text = rsPay!sumpay
Set rsPay = Nothing
Exit Sub
Error:
MsgBox "No values!", vbExclamation
End Sub
