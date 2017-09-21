VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18270
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   10230
   ScaleWidth      =   18270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Plan My Travel"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   22935
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   120
         Picture         =   "Menu.frx":5147A
         ScaleHeight     =   570
         ScaleWidth      =   600
         TabIndex        =   13
         Top             =   120
         Width           =   600
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8760
         TabIndex        =   11
         Top             =   2040
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66977793
         CurrentDate     =   42495
         MaxDate         =   42504
         MinDate         =   42495
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Menu.frx":57ED4
         Left            =   11280
         List            =   "Menu.frx":57EF0
         TabIndex        =   7
         Text            =   "-Select-"
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Menu.frx":57F6D
         Left            =   4920
         List            =   "Menu.frx":57F89
         TabIndex        =   6
         Text            =   "-Select-"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find Trains"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Results"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   960
         TabIndex        =   1
         Top             =   3840
         Visible         =   0   'False
         Width           =   16455
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fare"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   3480
            TabIndex        =   28
            Top             =   3240
            Visible         =   0   'False
            Width           =   9135
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rs "
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
               Left            =   7200
               TabIndex        =   34
               Top             =   1320
               Width           =   270
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rs "
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
               Left            =   4680
               TabIndex        =   33
               Top             =   1320
               Width           =   270
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rs "
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
               Left            =   1905
               TabIndex        =   32
               Top             =   1320
               Width           =   270
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sleeper"
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
               Left            =   6960
               TabIndex        =   31
               Top             =   480
               Width           =   795
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AC  3-Tier"
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
               Left            =   4200
               TabIndex        =   30
               Top             =   480
               Width           =   1125
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AC  2-Tier"
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
               Left            =   1440
               TabIndex        =   29
               Top             =   480
               Width           =   1125
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   1815
            Left            =   2160
            TabIndex        =   12
            Top             =   360
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   3201
            _Version        =   393216
            BackColor       =   16777215
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
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
                  LCID            =   16393
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
                  LCID            =   16393
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Seat Availability"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   3480
            TabIndex        =   4
            Top             =   3240
            Visible         =   0   'False
            Width           =   9135
            Begin VB.CommandButton Command6 
               Caption         =   "Book Now"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   6360
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1680
               Width           =   2055
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Book Now"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   1680
               Width           =   2055
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Book Now"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   1680
               Width           =   2055
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AC  2-Tier"
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
               Left            =   1440
               TabIndex        =   26
               Top             =   480
               Width           =   1125
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AC  3-Tier"
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
               Left            =   4200
               TabIndex        =   25
               Top             =   480
               Width           =   1125
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sleeper"
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
               Left            =   6960
               TabIndex        =   24
               Top             =   480
               Width           =   795
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Label11"
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
               Left            =   1680
               TabIndex        =   23
               Top             =   1080
               Width           =   720
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Label12"
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
               Left            =   4440
               TabIndex        =   22
               Top             =   1080
               Width           =   750
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Label13"
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
               Left            =   6960
               TabIndex        =   21
               Top             =   1080
               Width           =   750
            End
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FF8080&
            Caption         =   "Check Fare"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2520
            Width           =   3255
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FF8080&
            Caption         =   "Check Availability"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4680
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2520
            Width           =   3255
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Journey"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
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
         Left            =   4080
         TabIndex        =   9
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
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
         Left            =   9960
         TabIndex        =   8
         Top             =   1200
         Width           =   1200
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOG OUT"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   15960
      TabIndex        =   27
      Top             =   9000
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hi "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   17700
      TabIndex        =   17
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "My Account"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8280
      TabIndex        =   16
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Booking History"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   15
      Top             =   4800
      Width           =   2595
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan My Travel"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   14
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   7200
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   7200
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   7200
      Shape           =   2  'Oval
      Top             =   1920
      Width           =   4095
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim flag As String


Private Sub Hrithik()
On Error GoTo err1:
DataGrid3.Refresh
rs.Close
err1:
Exit Sub
End Sub

Private Sub Command1_Click()

sfrom = Combo1.Text
sto = Combo2.Text
sdoj = DTPicker1.Value

If Combo1.Text = "-Select-" Or Combo1.Text = "" And Combo2.Text = "-Select-" Or Combo2.Text = "" Then
    MsgBox "Please Select Source And Destination.", vbExclamation, "Error"
    
ElseIf Combo1.Text = "-Select-" Or Combo1.Text = "" Then
    MsgBox "Please Select Source.", vbExclamation, "Error"
    
ElseIf Combo2.Text = "-Select-" Or Combo2.Text = "" Then
    MsgBox "Please Select Destination.", vbExclamation, "Error"
    
ElseIf Combo1.Text = Combo2.Text Then
    MsgBox "Source And Destination Cannot Be The Same.", vbExclamation, "Error"
    
ElseIf Combo1.Text = "Mughal Sarai Jn" Or Combo1.Text = "Nagpur Jn" Or Combo1.Text = "Malda Town" Then
    Frame2.Visible = True
    Call Hrithik
    rs.Open "Select Train_Number, Train_Name, Stoppage, Destination, Arrival, Departure From Trains Where Stoppage='" & Combo1.Text & "' And Destination='" & Combo2.Text & "'", con, adOpenStatic, adLockReadOnly
    GoTo err:
    
ElseIf Combo2.Text = "Mughal Sarai Jn" Or Combo2.Text = "Nagpur Jn" Or Combo2.Text = "Malda Town" Then
    Frame2.Visible = True
    Call Hrithik
    rs.Open "Select Train_Number, Train_Name, Source, Stoppage, Arrival, Departure From Trains Where Source='" & Combo1.Text & "' And Stoppage='" & Combo2.Text & "'", con, adOpenStatic, adLockReadOnly
    GoTo err:
    
Else
    Frame2.Visible = True
    Call Hrithik
    rs.Open "Select Train_Number, Train_Name, Source, Destination, Arrival, Departure From Trains Where Source='" & Combo1.Text & "' And Destination='" & Combo2.Text & "'", con, adOpenStatic, adLockReadOnly
    'If (rs.BOF = rs.EOF) Then
    'Frame2.Visible = False
    'MsgBox "No Train Found Between Selected Pair Of Stations.", vbExclamation, "Ooops!"
    'Exit Sub
    'End If
    
err:
    Set DataGrid3.DataSource = rs
    
End If
End Sub

Private Sub Command2_Click()
Frame3.Visible = True
Frame4.Visible = False
trngrab = rs.Fields(0)
tnamegrab = rs.Fields(1)

rs1.Open "Select * From Seats Where Train_Number='" & trngrab & "' and DateOfJourney='" & sdoj & "'", con, adOpenStatic, adLockOptimistic

Label11.Caption = rs1.Fields(2)
If rs1.Fields(2).Value > 0 Then
    Command4.Enabled = True
    Command4.BackColor = RGB(0, 255, 0)
Else
    Command4.Enabled = False
    Command4.BackColor = RGB(255, 0, 0)
End If

Label12.Caption = rs1.Fields(3)
If rs1.Fields(3).Value > 0 Then
    Command5.Enabled = True
    Command5.BackColor = RGB(0, 255, 0)
Else
    Command5.Enabled = False
    Command5.BackColor = RGB(255, 0, 0)
End If

Label13.Caption = rs1.Fields(4)
If rs1.Fields(4).Value > 0 Then
    Command6.Enabled = True
    Command6.BackColor = RGB(0, 255, 0)
Else
    Command6.Enabled = False
    Command6.BackColor = RGB(255, 0, 0)
End If

rs1.Close

End Sub

Private Sub Command3_Click()
Frame4.Visible = True
Frame3.Visible = False
trngrab = rs.Fields(0)
rs3.Open "Select * from Fare where Train_Number = '" & trngrab & "'", con, adOpenStatic, adLockReadOnly
Label18.Caption = "Rs " & rs3.Fields(1).Value
Label19.Caption = "Rs " & rs3.Fields(2).Value
Label20.Caption = "Rs " & rs3.Fields(3).Value
End Sub

Private Sub Command4_Click()
sclass = "AC 2-Tier"
sfare = 2500
BookTicket.Show
Unload Me
End Sub

Private Sub Command5_Click()
sclass = "AC 3-Tier"
sfare = 2000
BookTicket.Show
Unload Me
End Sub

Private Sub Command6_Click()
sclass = "Sleeper"
sfare = 1500
BookTicket.Show
Unload Me
End Sub



Private Sub Form_Load()
Module1.Connection
On Error Resume Next
Label2.Caption = Label2.Caption + " " + fnamegrab + "!"
Set ado = New ADODB.Recordset
ado.Open "Select * From Trains", con, adOpenStatic, adLockPessimistic

End Sub


Private Sub Label1_Click()
Frame1.Visible = True
End Sub

Private Sub Label14_Click()
MsgBox "You Have Been Logged Out Successfully.", vbInformation, "TrainLine"
Menu.Hide
Welcome.Show
End Sub

Private Sub Label3_Click()
Cancel.Show
End Sub

Private Sub Label4_Click()
Edit.Show
End Sub

Private Sub Picture1_Click()
Frame1.Visible = False
End Sub

Private Sub Picture2_Click()
Menu.Hide
BookTicket.Hide
Cancel.Hide
Edit.Hide
Ticket.Hide
Welcome.Show
End Sub
