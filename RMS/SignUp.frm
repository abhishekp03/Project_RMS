VERSION 5.00
Begin VB.Form SignUp 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   18270
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   120
      Picture         =   "SignUp.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   600
      TabIndex        =   16
      Top             =   120
      Width           =   600
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   6600
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   6000
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   5400
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   4800
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H000080FF&
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label lblusername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Preferred Username"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   15
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblSignupJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register With TrainLine - It's Free And Always Will Be!"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2760
      TabIndex        =   14
      Top             =   1080
      Width           =   13215
   End
   Begin VB.Label lblCnfrmPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   13
      Top             =   6600
      Width           =   1845
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   12
      Top             =   6000
      Width           =   1020
   End
   Begin VB.Label lblMobile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   11
      Top             =   5400
      Width           =   720
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   10
      Top             =   4800
      Width           =   705
   End
   Begin VB.Label lblLname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   9
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label lblFname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   0
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "SignUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ModeVal As Boolean

Private Sub Validate_User()
Dim rs As New ADODB.Recordset
rs.Open "Select * From tblUserInfo Where Username = '" & Text3.Text & "'", con, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then
        ModeVal = False
        Exit Sub
    Else
        ModeVal = True
    End If
Set rs = Nothing
End Sub

Private Sub cmdRegister_Click()
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Then
        MsgBox "Do Not Leave Empty Field(s).", vbExclamation, "Error"
        Exit Sub
    ElseIf Text6.Text <> Text7.Text Then
        MsgBox "Passwords Do Not Match.", vbExclamation, "Error"
        Exit Sub
    Else
        Call Validate_User
        If ModeVal = False Then
        On Error GoTo err:
        ado.AddNew
        ado!UserName = Text3.Text
        ado!Password = Text6.Text
        ado!Firstname = Text1.Text
        ado!Lastname = Text2.Text
        ado!Email = Text4.Text
        ado!Mobile = Text5.Text
        ado.Save
err:
        MsgBox "Account Created Successfully.", vbInformation, "Congratulations"
        Login.Show
        End If
    End If
End Sub
        

Private Sub Form_Load()
Module1.Connection
On Error Resume Next
Set ado = New ADODB.Recordset
ado.Open "Select * From tblUserInfo", con, adOpenStatic, adLockPessimistic
End Sub

Private Sub Picture1_Click()
Welcome.Show
SignUp.Hide
End Sub
