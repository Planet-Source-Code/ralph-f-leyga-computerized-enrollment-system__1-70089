VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmparticular 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Particulars"
   ClientHeight    =   4260
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
   Icon            =   "frmparticular.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   3360
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5160
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Records"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
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
         ItemData        =   "frmparticular.frx":0B3A
         Left            =   1680
         List            =   "frmparticular.frx":0B50
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Select Year Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4005
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
   Begin VB.Label Label2 
      Caption         =   "Total Payment:"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmparticular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
DataGrid1.Visible = True
Set rsView = New ADODB.Recordset
rsView.Open "Select * from tparticular where yl='" & Combo1.Text & "' order by particular asc", db, 3, 3
Set DataGrid1.DataSource = rsView
  dbgrid
  
  On Error GoTo Error
Set rsPay = New ADODB.Recordset
rsPay.Open "Select sum(payable)as sumpay from tParticular where yl='" & Combo1.Text & "'", db, 3, 3
Text1.Text = rsPay!sumpay
Set rsPay = Nothing
Exit Sub
Error:
MsgBox "No values!", vbExclamation
Text1.Text = 0
End Sub


Public Sub dbgrid()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Width = 800
DataGrid1.Columns(2).Width = 2300
DataGrid1.Columns(3).Width = 1000
End Sub

Private Sub Command1_Click()
frmaddpar.Show vbModal
End Sub

Private Sub Command2_Click()
On Error GoTo Error
Dim repp As String
repp = MsgBox("Do you want to remove?", vbYesNo, "Confirm Delete")
If repp = vbYes Then
Set rsRemove = New ADODB.Recordset
rsRemove.Open "Select * from tparticular where particularID=" & Text2.Text & "", db, 3, 3
rsRemove.Delete
MsgBox "Data is remove.", vbInformation
Set rsRemove = Nothing
Call Combo1_Click
End If
Exit Sub
Error:
        MsgBox "No Active Record!", vbExclamation
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
On Error GoTo err
Text2.Text = rsView!particularID
Exit Sub
err:
       MsgBox "No record!", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsPay = Nothing
Set rsView = Nothing
End Sub

Private Sub Timer1_Timer()
Combo1_Click
Timer1.Enabled = False
End Sub
