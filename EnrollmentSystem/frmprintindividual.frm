VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "School Information"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4815
   Icon            =   "frmprintindividual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Height          =   1125
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
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
      Height          =   1125
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3615
      Width           =   4815
      _ExtentX        =   8493
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
      Caption         =   "Address"
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
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Name of School"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Set rsupdate = New ADODB.Recordset
rsupdate.Open "Update tschoolinfo set schoolname='" & Text1.Text & "', address='" & Text2.Text & "'" & _
"where id=" & frmmain.Text13.Text & "", db, 3, 3
Set rsupdate = Nothing
MsgBox "Data is Update!", vbInformation
frmmain.Timer2.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Set rsinfo = New ADODB.Recordset
rsinfo.Open "Select * from tschoolinfo", db, 3, 3
Text1.Text = rsinfo!schoolname
Text2.Text = rsinfo!address
Set rsinfo = Nothing
End Sub
