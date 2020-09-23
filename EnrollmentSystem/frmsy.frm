VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "School Year"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3030
   Icon            =   "frmsy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   2895
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
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Width           =   855
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
      Left            =   960
      TabIndex        =   1
      Top             =   4080
      Width           =   855
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
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsy.frx":0B3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lst 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList2"
      ForeColor       =   -2147483646
      BackColor       =   16761024
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SY  Number"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "School Year"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmsy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmaddsy.Show vbModal
End Sub

Private Sub Command2_Click()
On Error GoTo Error
Dim repp As String
repp = MsgBox("Do you want to remove?", vbYesNo, "Confirm Delete")
If repp = vbYes Then
Set rsRemove = New ADODB.Recordset
rsRemove.Open "Select * from tsy where synumber=" & Text1.Text & "", db, 3, 3
rsRemove.Delete
Set rsRemove = Nothing
Unload Me
MsgBox "Data is remove", vbExclamation
frmsy.Show vbModal
End If
Exit Sub
Error:
        MsgBox "No Active Record!", vbExclamation
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim criteria As String
Set rs = New ADODB.Recordset
    With rs
        criteria = "Select * from tsy order by sy"
        .Open criteria, db, 3, 3
            Do While Not .EOF
            
            lst.ListItems.Add , , !synumber, 1, 1
            lst.ListItems(lst.ListItems.Count).SubItems(1) = !sy
            .MoveNext
            Loop
        .close
    End With
  Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
End Sub

Private Sub lst_Click()
On Error Resume Next
Text1.Text = lst.SelectedItem.Text
Set rs = New ADODB.Recordset
Set rs = db.Execute("Select * from tsy where synumber=" & Text1.Text & "")
Text2.Text = rs!sy
Set rs = Nothing
End Sub

Private Sub lst_DblClick()
frmupdatesy.Text1.Text = Text2.Text
frmupdatesy.Text2.Text = Text1.Text
Unload Me
frmupdatesy.Show vbModal
End Sub
