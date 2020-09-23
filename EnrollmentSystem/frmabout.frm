VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About the software..."
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmabout.frx":0B3A
   ScaleHeight     =   3615
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1800
      ScaleHeight     =   1515
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Computerized Enrollment System 2008"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "All rights Reserved"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Vb6.0 Beginner Programmer"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2007 Leygasoftware"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1800
      Width           =   5055
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmabout.frx":6B0B
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4815
      End
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
      Left            =   3960
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   120
      Picture         =   "frmabout.frx":6C29
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
