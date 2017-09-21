VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Cancel 
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
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "Cancel.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Print Ticket"
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Cancel Ticket"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Bookings"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7920
      TabIndex        =   4
      Top             =   1680
      Width           =   2355
   End
End
Attribute VB_Name = "Cancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim temptrain As String
Dim tempclass As String
Dim tempdate As String
Dim result As String

Private Sub Command1_Click()
On Error Resume Next
result = MsgBox("Do You want To Cancel The Selected Ticket?", vbYesNo + vbQuestion, "Confirm Cancellation")
If result = vbNo Then
    Exit Sub
ElseIf result = vbYes Then
rs.Delete
rs.Save
rs.Close
temptrain = rs.Fields(6)
tempclass = rs.Fields(10)
tempdate = rs.Fields(2)
rs1.Open "Select * From Seats Where Train_Number='" & temptrain & "' And DateOfJourney='" & tempdate & "'", con, adOpenStatic, adLockOptimistic
If tempclass = "AC 2-Tier" Then
rs1.Fields(2).Value = rs1.Fields(2).Value + 1
ElseIf tempclass = "AC 3-Tier" Then
rs1.Fields(3).Value = rs1.Fields(3).Value + 1
ElseIf tempclass = "Sleeper" Then
rs1.Fields(4).Value = rs1.Fields(4).Value + 1
End If
rs1.Save
rs1.Close
End If
End Sub

Private Sub Command2_Click()
pnrgrab = rs.Fields(0)
Ticket.Show
End Sub

Private Sub Form_Load()
Module1.Connection
On Error Resume Next
Set ado = New ADODB.Recordset
ado.Open "Select * From tblUserInfo", con, adOpenStatic, adLockPessimistic
rs.Open "Select * From Reservation Where ByAccount='" & grab & "'", con, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = rs
End Sub
