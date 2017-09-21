VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Administrator 
   Caption         =   "Administrator"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   18270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18015
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   4320
         Width           =   17775
         Begin VB.CommandButton Command6 
            Caption         =   "Cancel"
            Height          =   1095
            Left            =   13560
            TabIndex        =   26
            Top             =   2640
            Width           =   4095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   1095
            Left            =   13560
            TabIndex        =   16
            Top             =   960
            Width           =   4095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Save Changes"
            Height          =   1095
            Left            =   9120
            TabIndex        =   15
            Top             =   840
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   4200
            TabIndex        =   14
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox Text2 
            Height          =   405
            Left            =   4200
            TabIndex        =   13
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox Text3 
            Height          =   405
            Left            =   4200
            TabIndex        =   12
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox Text4 
            Height          =   405
            Left            =   4200
            TabIndex        =   11
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox Text5 
            Height          =   405
            Left            =   4200
            TabIndex        =   10
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox Text6 
            Height          =   405
            Left            =   4200
            TabIndex        =   9
            Top             =   2760
            Width           =   2775
         End
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   4200
            TabIndex        =   8
            Top             =   3240
            Width           =   2775
         End
         Begin VB.TextBox Text8 
            Height          =   405
            Left            =   4200
            TabIndex        =   7
            Top             =   3720
            Width           =   2775
         End
         Begin VB.TextBox Text9 
            Height          =   405
            Left            =   4200
            TabIndex        =   6
            Top             =   4200
            Width           =   2775
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Arrival At Intermediate Station"
            Height          =   195
            Left            =   1320
            TabIndex        =   25
            Top             =   4320
            Width           =   2085
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Departure From Intermediate Station"
            Height          =   195
            Left            =   1320
            TabIndex        =   24
            Top             =   3840
            Width           =   2550
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Intermediate Station"
            Height          =   195
            Left            =   1320
            TabIndex        =   23
            Top             =   3360
            Width           =   1410
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Arrival"
            Height          =   195
            Left            =   1320
            TabIndex        =   22
            Top             =   2880
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Departure"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   2400
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Train Number"
            Height          =   195
            Left            =   1320
            TabIndex        =   20
            Top             =   480
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Train Name"
            Height          =   195
            Left            =   1320
            TabIndex        =   19
            Top             =   960
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Source"
            Height          =   195
            Left            =   1320
            TabIndex        =   18
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Destination"
            Height          =   195
            Left            =   1320
            TabIndex        =   17
            Top             =   1920
            Width           =   795
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Administrator.frx":0000
         Height          =   3855
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   6800
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
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   1095
         Left            =   13560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   1095
         Left            =   13560
         TabIndex        =   2
         Top             =   2760
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   1095
         Left            =   13560
         TabIndex        =   1
         Top             =   1560
         Width           =   4095
      End
   End
   Begin VB.Menu TrainDB 
      Caption         =   "Train Database"
   End
   Begin VB.Menu SeatDB 
      Caption         =   "Seat Database"
   End
   Begin VB.Menu PNRDB 
      Caption         =   "PNR Database"
   End
   Begin VB.Menu FareDB 
      Caption         =   "Fare Database"
   End
   Begin VB.Menu UserAC 
      Caption         =   "User Accounts"
   End
   Begin VB.Menu AdminAC 
      Caption         =   "Admin Accounts"
   End
End
Attribute VB_Name = "Administrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
