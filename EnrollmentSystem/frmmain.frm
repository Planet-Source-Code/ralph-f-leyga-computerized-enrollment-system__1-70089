VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computerized Enrollment System 2008 - "
   ClientHeight    =   8430
   ClientLeft      =   150
   ClientTop       =   555
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   4680
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7680
      Top             =   3960
   End
   Begin VB.TextBox Text13 
      Height          =   315
      Left            =   4440
      TabIndex        =   23
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox Text12 
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   2280
      TabIndex        =   21
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Height          =   315
      Left            =   4200
      TabIndex        =   20
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   6240
      TabIndex        =   19
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   2160
      TabIndex        =   18
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   6240
      TabIndex        =   16
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   6240
      TabIndex        =   15
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   4200
      TabIndex        =   14
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   4200
      TabIndex        =   13
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Top             =   9000
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1095
      Left            =   8520
      TabIndex        =   11
      Top             =   8640
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
      _Version        =   393216
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   9000
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   12303
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
      Caption         =   "General Information"
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
   Begin VB.Frame Frame1 
      Caption         =   "Filter Records"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12135
      Begin VB.TextBox Ttotal 
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
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
         ItemData        =   "frmmain.frx":0B3A
         Left            =   4080
         List            =   "frmmain.frx":0B50
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1695
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
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Total Number of Students:"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Grade/YL"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "School Year:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8175
      Width           =   12150
      _ExtentX        =   21431
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
   Begin MSComctlLib.ImageList IMG1 
      Left            =   5040
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":340A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":48B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5190
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6344
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":74F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8224
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8676
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":936C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A098
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A972
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B24C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":C400
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":CCDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":D5B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":DE8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":E768
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":F042
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":F494
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":F8E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":FD38
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1018A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":10C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1152E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":11E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":126E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":12FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1340E
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":13CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1413A
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1567B
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15FE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":16685
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":16D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":174A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":17C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1835B
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":18B3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "IMG1"
      DisabledImageList=   "IMG1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   32
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   31
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   35
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   25
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu file 
      Caption         =   "&File Transaction"
      Begin VB.Menu studentinformation 
         Caption         =   "&Student Information"
         Shortcut        =   {F1}
      End
      Begin VB.Menu particular 
         Caption         =   "&Particular"
         Shortcut        =   {F2}
      End
      Begin VB.Menu changeinfo 
         Caption         =   "&Change School Information"
         Shortcut        =   ^D
      End
      Begin VB.Menu users 
         Caption         =   "&Users"
         Shortcut        =   ^U
      End
      Begin VB.Menu sySetting 
         Caption         =   "&School Year Setting"
         Shortcut        =   ^S
      End
      Begin VB.Menu bar01 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "&Close Application"
      End
   End
   Begin VB.Menu report 
      Caption         =   "&Report"
      Begin VB.Menu allstudent 
         Caption         =   "&Print All Student"
         Shortcut        =   {F3}
      End
      Begin VB.Menu printindividual 
         Caption         =   "&Print Student Individually"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About"
      Begin VB.Menu bar02 
         Caption         =   "-"
      End
      Begin VB.Menu aboutsoftware 
         Caption         =   "&About the Software"
         Shortcut        =   ^A
      End
      Begin VB.Menu sk 
         Caption         =   "&Shortcut Keys"
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub aboutsoftware_Click()
frmabout.Show vbModal
End Sub

Private Sub allstudent_Click()
If Combo2.Text <> "" Then
'DataGrid1.Enabled = False
Set DataReport2.DataSource = rsView
DataReport2.Sections("Section2").Controls("L1").Caption = Combo1.Text
DataReport2.Sections("Section2").Controls("L2").Caption = Combo2.Text
DataReport2.Sections("Section2").Controls("L3").Caption = Ttotal.Text
DataReport2.Sections("Section2").Controls("L4").Caption = Text9.Text
DataReport2.Sections("Section2").Controls("Label1").Caption = Text11.Text
DataReport2.Sections("Section2").Controls("Label2").Caption = Text12.Text
DataReport2.Show vbModal
Else
MsgBox "Select a Grade/Year Level", vbExclamation
End If
End Sub

Private Sub changeinfo_Click()
frmSetting.Show vbModal
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub Combo2_Click()
DataGrid1.Visible = True
On Error GoTo Error
Set rsView = New ADODB.Recordset
rsView.Open "Select * from tstudent where sy='" & Combo1.Text & "' and yl='" & Combo2.Text & "' order by lname asc", db, 3, 3
Set DataGrid1.DataSource = rsView
  dbgrid
Set rsView1 = New ADODB.Recordset
rsView1.Open "Select * from tparticular where yl='" & Combo2.Text & "' order by particular asc", db, 3, 3
Set DataGrid2.DataSource = rsView1
Set rsCount = New ADODB.Recordset
rsCount.Open "Select count(studentid) as sumstudent from tstudent where yl = '" & Combo2.Text & "'", db, 3, 3
Ttotal.Text = rsCount!sumstudent
Set rsCount = Nothing
Text9.Text = rsView!charges
Exit Sub
Error:
    MsgBox "No result!", vbExclamation
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
If Ttotal.Text <> 0 Then
Text1.Text = rsView!studentid
Text2.Text = rsView!studentnumber
Text3.Text = rsView!lname
Text4.Text = rsView!fname
Text5.Text = rsView!mname
Text6.Text = rsView!age
Text7.Text = rsView!gender
Text8.Text = rsView!bal
Text9.Text = rsView!charges
Text10.Text = rsView!payment
Else
MsgBox "No Record!", vbExclamation
End If
End Sub

Private Sub DataGrid1_DblClick()
frmpay.Show vbModal
End Sub





Private Sub Form_Load()
'========================POWERED BY: RALPH F. LEYGA==============================
'========================NOTRE DAME OF MIDSAYAP COLLEGE==========================
'========================TEXT OR CALL: 09057805663 ==============================
'========================E-mail: ralphleyga@yahoo.com============================
dbase
chnge
cbosy
End Sub

Private Sub particular_Click()
frmparticular.Show vbModal
End Sub

Private Sub printindividual_Click()
On Error Resume Next
If Text1.Text <> "" Then
'DataGrid1.Enabled = False
Set DataReport1.DataSource = rsView1
DataReport1.Sections("Section2").Controls("L1").Caption = Text2.Text
DataReport1.Sections("Section2").Controls("L2").Caption = Text3.Text
DataReport1.Sections("Section2").Controls("L3").Caption = Text4.Text
DataReport1.Sections("Section2").Controls("L4").Caption = Text5.Text
DataReport1.Sections("Section2").Controls("L5").Caption = Text6.Text
DataReport1.Sections("Section2").Controls("L6").Caption = Text7.Text
DataReport1.Sections("Section2").Controls("L7").Caption = Combo2.Text
DataReport1.Sections("Section2").Controls("L8").Caption = Combo1.Text
DataReport1.Sections("Section2").Controls("L9").Caption = Text10.Text
DataReport1.Sections("Section2").Controls("L10").Caption = Text8.Text
DataReport1.Sections("Section2").Controls("L11").Caption = Text9.Text
DataReport1.Sections("Section2").Controls("Label1").Caption = Text11.Text
DataReport1.Sections("Section2").Controls("Label2").Caption = Text12.Text
DataReport1.Show vbModal
Else
MsgBox "Select a student", vbExclamation
End If
'If DataGrid1.Visible = True Then
'frmprintindividual.Show vbModal
'Else
'MsgBox "Activate the data first!", vbExclamation
'End If
End Sub

Private Sub sk_Click()
frmshortcutkey.Show vbModal
End Sub

Private Sub studentinformation_Click()
frmstudent.Show vbModal
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
DataGrid1.Columns(4).Width = 1500
DataGrid1.Columns(5).Width = 900
DataGrid1.Columns(6).Width = 900
DataGrid1.Columns(7).Visible = False
DataGrid1.Columns(8).Visible = False
DataGrid1.Columns(9).Width = 1200
DataGrid1.Columns(10).Width = 1200
DataGrid1.Columns(11).Width = 1300
End Sub

Private Sub sySetting_Click()
frmsy.Show vbModal
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1:
    studentinformation_Click
Case 2:
    particular_Click
Case 3:
    changeinfo_Click
Case 4:
    users_Click
Case 5:
    allstudent_Click
Case 6:
    printindividual_Click
    Case 7:
    aboutsoftware_Click
End Select
End Sub

Private Sub Timer1_Timer()
Call Combo2_Click
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
chnge
Timer2.Enabled = False
End Sub

Public Sub chnge()
Set rsinfo = New ADODB.Recordset
rsinfo.Open "Select * from tschoolinfo", db, 3, 3
Text11.Text = rsinfo!schoolname
Text12.Text = rsinfo!address
Text13.Text = rsinfo!id
frmmain.Caption = "Computerized Enrollment System 2008 - " + Text11.Text
Set rsinfo = Nothing
End Sub

Private Sub Timer3_Timer()
frmsplash.Show vbModal
Timer3.Enabled = False
End Sub

Private Sub users_Click()
frmuser.Show vbModal
End Sub
