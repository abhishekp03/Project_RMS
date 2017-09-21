VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User Login"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdLogin1 
      BackColor       =   &H00FF8080&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.PictureBox picUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1800
      Picture         =   "Login.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtLoginPass 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtLoginID 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label lblSignup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   3240
      TabIndex        =   4
      Top             =   4560
      Width           =   810
   End
   Begin VB.Label lblLoginJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Don't Have An Account?"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   3
      Top             =   4560
      Width           =   2250
   End
   Begin VB.Label lblLoginPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password  : "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Label lblLoginId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin1_Click()
Dim rs As New ADODB.Recordset
If txtLoginID.Text = "" Then
    MsgBox "Please Enter Your Username.", vbExclamation, "Error"
ElseIf txtLoginPass.Text = "" Then
    MsgBox "Please Enter Your Password.", vbExclamation, "Error"
Else
    On Error GoTo err:
    rs.Open "Select * from tblUserInfo Where Username = '" & txtLoginID.Text & "'", con, adOpenStatic, adLockReadOnly
    If rs!Password = txtLoginPass.Text Then
        grab = txtLoginID.Text
        fnamegrab = rs!Firstname
        Menu.Show
        Unload Me
        Exit Sub
    Else
        MsgBox "The Password You Have Entered Is Incorrect.", vbCritical, "Incorrect Password"
        txtLoginPass.SetFocus
        Exit Sub
    End If
End If
err:
    MsgBox "Username Not Found. Enter Correct Username", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub Form_Load()
Module1.Connection
On Error Resume Next
Set ado = New ADODB.Recordset
ado.Open "Select * From tblUserInfo", con, adOpenStatic, adLockPessimistic
End Sub

Private Sub lblSignup_Click()
SignUp.Show
End Sub
