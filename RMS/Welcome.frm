VERSION 5.00
Begin VB.Form Welcome 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Welcome.frx":0000
   ScaleHeight     =   9930
   ScaleMode       =   0  'User
   ScaleWidth      =   18390
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7800
      Picture         =   "Welcome.frx":6C6DB
      ScaleHeight     =   975
      ScaleWidth      =   3375
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton cmdSignup 
      BackColor       =   &H008080FF&
      Caption         =   "SIGN UP"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu menAdmin 
      Caption         =   "Administrator Login"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu menAbout 
         Caption         =   "About TrainLine"
      End
      Begin VB.Menu menUpdate 
         Caption         =   "Check For Updates"
      End
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()
Login.Show
End Sub

Private Sub cmdSignup_Click()
SignUp.Show
End Sub

Private Sub menAbout_Click()
About.Show
End Sub

Private Sub menAdmin_Click()
AdminLogin.Show
End Sub

Private Sub menUpdate_Click()
Update.Show
End Sub
