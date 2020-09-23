VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Information"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5400
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   4440
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3413
      _Version        =   393216
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5880
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
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
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2970
      Width           =   7560
      _ExtentX        =   13335
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
   Begin VB.Label Label1 
      Caption         =   "Users Personal Information"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmadduser.Show vbModal
End Sub

Private Sub Command2_Click()
On Error GoTo Error
Dim repp As String
repp = MsgBox("Do you want to remove?", vbYesNo, "Confirm Delete")
If repp = vbYes Then
Set rsRemove = New ADODB.Recordset
rsRemove.Open "Select * from tuser where userID=" & Text1.Text & "", db, 3, 3
rsRemove.Delete
MsgBox "Data is remove.", vbInformation
Set rsRemove = Nothing
Timer1.Enabled = True
End If
Exit Sub
Error:
        MsgBox "No Active Record!", vbExclamation
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
Text1.Text = rs!userid
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
rs.Open "Select * from tuser order by username asc", db, 3, 3
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Width = 1500
DataGrid1.Columns(2).Width = 1700
DataGrid1.Columns(3).Width = 1700
DataGrid1.Columns(4).Width = 2000
DataGrid1.Columns(5).Visible = False
End Sub

Private Sub Timer1_Timer()
Call Form_Load
Timer1.Enabled = False
End Sub
