VERSION 5.00
Begin VB.Form frmshortcutkey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shortcut Keys"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmshortcutkey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   "Shortcut Keys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "frmshortcutkey.frx":0B3A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl + A"
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "About the Software"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Student Individually"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Print All Students"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl + S"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "School Year Setting"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl + U"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl + D"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Change School Info"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Information"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut Keys"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menus"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   120
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "frmshortcutkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
