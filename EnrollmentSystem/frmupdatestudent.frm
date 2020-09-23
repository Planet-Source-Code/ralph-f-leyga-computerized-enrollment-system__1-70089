VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmupdatestudent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update Student"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4470
   Icon            =   "frmupdatestudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   4560
      TabIndex        =   24
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   14
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   2760
         Width           =   1815
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
         ItemData        =   "frmupdatestudent.frx":0B3A
         Left            =   1560
         List            =   "frmupdatestudent.frx":0B44
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
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
         ItemData        =   "frmupdatestudent.frx":0B56
         Left            =   1560
         List            =   "frmupdatestudent.frx":0B6C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
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
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Charges:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.TextBox tID 
      Height          =   315
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
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
      Left            =   1680
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   4500
      Width           =   4470
      _ExtentX        =   7885
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmupdatestudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If T1.Text <> "" And T2.Text <> "" And T3.Text <> "" And T4.Text <> "" And T5.Text <> "" And t6.Text <> "" And t7.Text <> "" And t8.Text <> "" Then
Set rsupdate = New ADODB.Recordset
rsupdate.Open "Update tstudent set lname='" & T2.Text & "', fname='" & T3.Text & "', mname='" & T4.Text & "', gender='" & T5.Text & "', age='" & t6.Text & "', yl='" & t7.Text & "', sy='" & t8.Text & "', charges='" & T9.Text & "', bal='" & T10.Text & "'" & _
"where studentid=" & Tid.Text & "", db, 3, 3
MsgBox "Data is update!", vbInformation
frmstudent.Timer1.Enabled = True
Unload Me
Else
MsgBox "Cannot be null!", vbExclamation
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cbosy
Set rs = New ADODB.Recordset
rs.Open "Select * from tstudent where studentid=" & frmstudent.Text1.Text & "", db, 3, 3
Tid.Text = rs!studentid
T1.Text = rs!studentnumber
T2.Text = rs!lname
T3.Text = rs!fname
T4.Text = rs!mname
T5.Text = rs!gender
t6.Text = rs!age
t7.Text = rs!yl
t8.Text = rs!sy
Set rs = Nothing
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

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
Set rsupdate = Nothing
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
