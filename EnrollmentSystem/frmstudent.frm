VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmstudent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Student List"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmstudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   6720
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   4560
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761024
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Refresh"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Search"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Records"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frmstudent.frx":0B3A
         Left            =   4680
         List            =   "frmstudent.frx":0B50
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
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
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Year Level:"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "School Year:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   6495
      _ExtentX        =   11456
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
Attribute VB_Name = "frmstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
DataGrid1.Visible = True
Set rsSYView = New ADODB.Recordset
rsSYView.Open "Select * from tstudent where sy='" & Combo1.Text & "' and yl='" & Combo2.Text & "' order by lname asc", db, 3, 3
Set DataGrid1.DataSource = rsSYView
dbgrid
DataGrid1.Enabled = True
End Sub

Private Sub Command1_Click()
frmaddstudent.Show vbModal
End Sub

Private Sub Command2_Click()
On Error GoTo Error
Dim repp As String
repp = MsgBox("Do you want to remove?", vbYesNo, "Confirm Delete")
If repp = vbYes Then
Set rsRemove = New ADODB.Recordset
rsRemove.Open "Select * from tstudent where studentID=" & Text1.Text & "", db, 3, 3
rsRemove.Delete
MsgBox "Data is remove.", vbInformation
Set rsRemove = Nothing
Call Command4_Click
End If
Exit Sub
Error:
        MsgBox "No Active Record!", vbExclamation
End Sub

Private Sub Command3_Click()
Dim strSearch As String
'Dim str1, str2, str3, str4 As String
strSearch = InputBox("Search for the last name.", "Search Option")
Set rsSYView = New ADODB.Recordset
'dbase
rsSYView.Open "Select * from tstudent where lname='" & strSearch & "' and sy='" & Combo1.Text & "' order by lname asc ", db, 3, 3
Set DataGrid1.DataSource = rsSYView
dbgrid
End Sub

Private Sub Command4_Click()
Call Combo2_Click
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
On Error GoTo err
Text1.Text = rsSYView!studentid
Exit Sub
err:
MsgBox "No Records!", vbExclamation
End Sub

Private Sub DataGrid1_DblClick()
frmupdatestudent.Show vbModal
End Sub

Private Sub Form_Load()
cbosy
End Sub
Public Sub cbosy()
Set rsSY = New ADODB.Recordset
rsSY.Open "Select * from tSY order by sy asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsSY.EOF
        Combo1.AddItem rsSY!sy
        rsSY.MoveNext
    Loop
    Set rsSY = Nothing
    End Sub


Public Sub dbgrid()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Width = 1500
DataGrid1.Columns(2).Width = 1500
DataGrid1.Columns(3).Width = 1500
DataGrid1.Columns(4).Width = 1400
DataGrid1.Columns(5).Visible = False
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(7).Visible = False
DataGrid1.Columns(8).Visible = False
DataGrid1.Columns(9).Visible = False
DataGrid1.Columns(10).Visible = False
DataGrid1.Columns(11).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsSY = Nothing
Set rsSYView = Nothing
End Sub

Private Sub Timer1_Timer()
Call Combo2_Click
Timer1.Enabled = False
End Sub
