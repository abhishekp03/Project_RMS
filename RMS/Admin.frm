VERSION 5.00
Begin VB.Form AdminLogin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Administrator Login"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdmin 
      BackColor       =   &H000000FF&
      Caption         =   "Administrator Login"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtLoginPass 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txtLoginID 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   2280
      Picture         =   "Admin.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblLoginPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
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
      Left            =   840
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
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
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "AdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdmin_Click()
MsgBox "This Section Is Currently Under Maintenance!", vbInformation, "Ooops"
Exit Sub
Dim rs As New ADODB.Recordset
If txtLoginID.Text = "" Then
    MsgBox "Please Enter Your Username.", vbExclamation, "Error"
ElseIf txtLoginPass.Text = "" Then
    MsgBox "Please Enter Your Password.", vbExclamation, "Error"
Else
    On Error GoTo err:
    rs.Open "Select * from tblAdminInfo Where Admin_name = '" & txtLoginID.Text & "'", con, adOpenStatic, adLockReadOnly
    If rs!Password = txtLoginPass.Text Then
        Administrator.Show
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
Administrator.Visible = False
End Sub
