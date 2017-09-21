VERSION 5.00
Begin VB.Form BookTicket 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   18270
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "BookTicket.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   72
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   18255
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enter Card Details"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3480
         Width           =   3975
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   6600
         Picture         =   "BookTicket.frx":6A5A
         ScaleHeight     =   4935
         ScaleWidth      =   5175
         TabIndex        =   70
         Top             =   4200
         Width           =   5175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Debit Card"
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
         Left            =   10560
         TabIndex        =   69
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credit Card"
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
         Left            =   8280
         TabIndex        =   68
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7215
         Left            =   1080
         TabIndex        =   53
         Top             =   3120
         Visible         =   0   'False
         Width           =   18255
         Begin VB.TextBox Text12 
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
            Left            =   7200
            TabIndex        =   64
            Top             =   1680
            Width           =   5295
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Make Payment"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   4320
            Width           =   3495
         End
         Begin VB.ComboBox Combo6 
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
            ItemData        =   "BookTicket.frx":F66D
            Left            =   8280
            List            =   "BookTicket.frx":F68F
            TabIndex        =   60
            Text            =   "YYYY"
            Top             =   3360
            Width           =   1335
         End
         Begin VB.ComboBox Combo5 
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
            ItemData        =   "BookTicket.frx":F6CF
            Left            =   7200
            List            =   "BookTicket.frx":F6F7
            TabIndex        =   59
            Text            =   "MM"
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox Text11 
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
            IMEMode         =   3  'DISABLE
            Left            =   7200
            MaxLength       =   3
            PasswordChar    =   "*"
            TabIndex        =   58
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox Text10 
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
            Left            =   11520
            MaxLength       =   4
            TabIndex        =   57
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text9 
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
            Left            =   10080
            MaxLength       =   4
            TabIndex        =   56
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text8 
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
            Left            =   8640
            MaxLength       =   4
            TabIndex        =   55
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text7 
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
            Left            =   7200
            MaxLength       =   4
            TabIndex        =   54
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Card Expiry"
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
            Left            =   3840
            TabIndex        =   66
            Top             =   3360
            Width           =   1110
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CVV Number"
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
            Left            =   3840
            TabIndex        =   65
            Top             =   2520
            Width           =   1230
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cardholder Name"
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
            Left            =   3840
            TabIndex        =   63
            Top             =   1800
            Width           =   1635
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Card Number"
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
            Left            =   3840
            TabIndex        =   62
            Top             =   960
            Width           =   1245
         End
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "BookTicket.frx":F72B
         Left            =   8280
         List            =   "BookTicket.frx":F73B
         TabIndex        =   51
         Text            =   "-Select Your Card Type-"
         Top             =   2520
         Width           =   4575
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Card Type : "
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
         Left            =   6480
         TabIndex        =   67
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Your Merchant  : "
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
         Left            =   5640
         TabIndex        =   52
         Top             =   2520
         Width           =   2565
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6375
      Left            =   120
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   18015
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Undo"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5160
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Proceed To Payment"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5160
         Width           =   2895
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fare : "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10200
         TabIndex        =   49
         Top             =   4200
         Width           =   1515
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   14880
         TabIndex        =   48
         Top             =   3240
         Width           =   60
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   11640
         TabIndex        =   47
         Top             =   3240
         Width           =   60
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   6000
         TabIndex        =   46
         Top             =   3240
         Width           =   60
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   14880
         TabIndex        =   45
         Top             =   2640
         Width           =   60
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   11640
         TabIndex        =   44
         Top             =   2640
         Width           =   60
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   6000
         TabIndex        =   43
         Top             =   2640
         Width           =   60
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   14880
         TabIndex        =   42
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   11640
         TabIndex        =   41
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   6000
         TabIndex        =   40
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   39
         Top             =   3240
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   37
         Top             =   2040
         Width           =   210
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14760
         TabIndex        =   36
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11520
         TabIndex        =   35
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   34
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passenger No."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2520
         TabIndex        =   33
         Top             =   1320
         Width           =   1725
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   5760
      TabIndex        =   1
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Book Ticket"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "BookTicket.frx":F76A
      Left            =   14760
      List            =   "BookTicket.frx":F774
      TabIndex        =   9
      Text            =   "-Select-"
      Top             =   6720
      Width           =   975
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
      ItemData        =   "BookTicket.frx":F786
      Left            =   14760
      List            =   "BookTicket.frx":F790
      TabIndex        =   6
      Text            =   "-Select-"
      Top             =   6120
      Width           =   975
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
      ItemData        =   "BookTicket.frx":F7A2
      Left            =   14760
      List            =   "BookTicket.frx":F7AC
      TabIndex        =   3
      Text            =   "-Select-"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text6 
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
      Left            =   11520
      TabIndex        =   8
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text5 
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
      Left            =   11520
      TabIndex        =   5
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox Text4 
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
      Left            =   11520
      TabIndex        =   2
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text3 
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
      Left            =   5760
      TabIndex        =   7
      Top             =   6720
      Width           =   3255
   End
   Begin VB.TextBox Text2 
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
      Left            =   5760
      TabIndex        =   4
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
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
      Left            =   13080
      TabIndex        =   31
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
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
      Left            =   11520
      TabIndex        =   30
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
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
      Left            =   5280
      TabIndex        =   29
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Journey:"
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
      Left            =   3360
      TabIndex        =   28
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
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
      Left            =   13080
      TabIndex        =   27
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination:"
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
      Left            =   11520
      TabIndex        =   26
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
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
      Left            =   5280
      TabIndex        =   25
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
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
      Left            =   3360
      TabIndex        =   24
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
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
      Left            =   13080
      TabIndex        =   23
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Train Name:"
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
      Left            =   11520
      TabIndex        =   22
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
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
      Left            =   5280
      TabIndex        =   21
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Train Number:"
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
      Left            =   3360
      TabIndex        =   20
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note  :  A Maximum Of 3 Passengers Can Travel Per Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   10560
      TabIndex        =   19
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      Enabled         =   0   'False
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
      Left            =   3000
      TabIndex        =   18
      Top             =   6720
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Enabled         =   0   'False
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
      Left            =   3000
      TabIndex        =   17
      Top             =   6120
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      Enabled         =   0   'False
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
      Left            =   3000
      TabIndex        =   16
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label lblPsex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14760
      TabIndex        =   15
      Top             =   4920
      Width           =   465
   End
   Begin VB.Label lblPage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11520
      TabIndex        =   14
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lblPname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   13
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lblSr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   0
      Top             =   4920
      Width           =   1770
   End
End
Attribute VB_Name = "BookTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim flag As Long
Dim coach As String
Dim berth As Integer

Private Sub Command1_Click()

If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" And Combo1.Text = "-Select-" And Combo2.Text = "-Select-" And Combo3.Text = "-Select-" Then
    MsgBox "Enter The Details Of At Least One Passenger.", vbExclamation, "Error"
    Exit Sub
End If

If Text1.Text <> "" And Text4.Text <> "" And Combo1.Text <> "-Select-" Then
    Label1.Enabled = True
Else
    MsgBox "Please Fill In / Select The Particulars Properly."
    Exit Sub
End If

If Text2.Text = "" And Text5.Text = "" And Combo2.Text = "-Select-" And Text3.Text = "" And Text6.Text = "" And Combo3.Text = "-Select-" Then
    GoTo yahoo:
ElseIf Text2.Text <> "" And Text5.Text <> "" And Combo2.Text <> "-Select-" Then
    Label2.Enabled = True
Else
    MsgBox "Please Fill In / Select The Particulars Properly."
    Exit Sub
End If

If Text3.Text = "" And Text6.Text = "" And Combo3.Text = "-Select-" Then
    GoTo yahoo:
ElseIf Text3.Text <> "" And Text6.Text <> "" And Combo3.Text <> "-Select-" Then
    Label3.Enabled = True
Else
    MsgBox "Please Fill In / Select The Particulars Properly."
    Exit Sub
End If

yahoo:
Frame1.Visible = True

If Label1.Enabled = True Then
    Label24.Caption = Text1.Text
    Label25.Caption = Text4.Text
    Label26.Caption = Combo1.Text
End If

If Label2.Enabled = True Then
    Label22.Visible = True
    Label27.Caption = Text2.Text
    Label28.Caption = Text5.Text
    Label29.Caption = Combo2.Text
End If

If Label3.Enabled = True Then
    Label23.Visible = True
    Label30.Caption = Text3.Text
    Label31.Caption = Text6.Text
    Label32.Caption = Combo3.Text
End If

End Sub

Private Sub Command2_Click()
Frame2.Visible = True
End Sub

Private Sub Command3_Click()
    Frame1.Visible = False
    Label1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
End Sub

Private Sub Command4_Click()

If Text7.MaxLength <> 4 Or Text8.MaxLength <> 4 Or Text8.MaxLength <> 4 Or Text9.MaxLength <> 4 Or Text10.Text = "" Or Text11.MaxLength <> 3 Or Combo5.Text = "MM" Or Combo6.Text = "YYYY" Then
MsgBox "The Details Entered Are Incorrect. Please Verify Your Credentials.", vbExclamation, "Error"
Exit Sub
End If

If sclass = "AC 2-Tier" Then
    coach = "HA1"
ElseIf sclass = "AC 3-Tier" Then
    coach = "AB1"
ElseIf sclass = "Sleeper" Then
    coach = "S1"
End If


On Error Resume Next

If Label1.Enabled = True And Label2.Enabled = False And Label3.Enabled = False Then

sfare = sfare * 1

rs.Open "Select * from Reservation", con, adOpenStatic, adLockOptimistic
rs.AddNew
rs.Fields(1) = sdoj
rs.Fields(2) = Text1.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Combo1.Text
rs.Fields(5) = trngrab
rs.Fields(6) = tnamegrab
rs.Fields(7) = sfrom
rs.Fields(8) = sto
rs.Fields(9) = sclass
rs.Fields(10) = coach
rs.Fields(11) = sfare
rs.Fields(12) = grab
rs.Save
rs.Fields(0) = rs.Fields(13) + 67843321
rs.Save
flag = rs.Fields(0)
rs.Close

rs1.Open "Select * From Seats Where Train_Number ='" & trngrab & "' and DateOfJourney = '" & sdoj & "'", con, adOpenStatic, adLockOptimistic
If sclass = "AC 2-Tier" Then
    rs1.Fields(2).Value = rs1.Fields(2).Value - 1
ElseIf sclass = "AC 3-Tier" Then
    rs1.Fields(3).Value = rs1.Fields(3).Value - 1
ElseIf sclass = "Sleeper" Then
    rs1.Fields(4).Value = rs1.Fields(4).Value - 1
End If

rs1.Update
rs1.Save
rs1.Close


ElseIf Label1.Enabled = True And Label2.Enabled = True And Label3.Enabled = False Then
sfare = sfare * 2

rs.Open "Select * from Reservation", con, adOpenStatic, adLockOptimistic
rs.AddNew
rs.Fields(1) = sdoj
rs.Fields(2) = Text1.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Combo1.Text
rs.Fields(5) = trngrab
rs.Fields(6) = tnamegrab
rs.Fields(7) = sfrom
rs.Fields(8) = sto
rs.Fields(9) = sclass
rs.Fields(10) = coach
rs.Fields(11) = sfare
rs.Fields(12) = grab
rs.Save
rs.Fields(0) = rs.Fields(13) + 67843321
rs.Save
flag = rs.Fields(0)
rs.AddNew
rs.Fields(0) = flag
rs.Fields(1) = sdoj
rs.Fields(2) = Text2.Text
rs.Fields(3) = Text5.Text
rs.Fields(4) = Combo2.Text
rs.Fields(5) = trngrab
rs.Fields(6) = tnamegrab
rs.Fields(7) = sfrom
rs.Fields(8) = sto
rs.Fields(9) = sclass
rs.Fields(10) = coach
rs.Fields(11) = sfare
rs.Fields(12) = grab
rs.Save
rs.Close

rs1.Open "Select * From Seats Where Train_Number ='" & trngrab & "' and DateOfJourney = '" & sdoj & "'", con, adOpenStatic, adLockOptimistic
If sclass = "AC 2-Tier" Then
    rs1.Fields(2).Value = rs1.Fields(2).Value - 2
ElseIf sclass = "AC 3-Tier" Then
    rs1.Fields(3).Value = rs1.Fields(3).Value - 2
ElseIf sclass = "Sleeper" Then
    rs1.Fields(4).Value = rs1.Fields(4).Value - 2
End If
rs1.Update
rs1.Save
rs1.Close

Else


rs.Open "Select * from Reservation", con, adOpenStatic, adLockOptimistic
rs.AddNew
rs.Fields(1) = sdoj
rs.Fields(2) = Text1.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Combo1.Text
rs.Fields(5) = trngrab
rs.Fields(6) = tnamegrab
rs.Fields(7) = sfrom
rs.Fields(8) = sto
rs.Fields(9) = sclass
rs.Fields(10) = coach
rs.Fields(11) = sfare
rs.Fields(12) = grab
rs.Save
rs.Fields(0) = rs.Fields(13) + 67843321
rs.Save
flag = rs.Fields(0)
rs.AddNew
rs.Fields(0) = flag
rs.Fields(1) = sdoj
rs.Fields(2) = Text2.Text
rs.Fields(3) = Text5.Text
rs.Fields(4) = Combo2.Text
rs.Fields(5) = trngrab
rs.Fields(6) = tnamegrab
rs.Fields(7) = sfrom
rs.Fields(8) = sto
rs.Fields(9) = sclass
rs.Fields(10) = coach
rs.Fields(11) = sfare
rs.Fields(12) = grab
rs.Save
rs.AddNew
rs.Fields(0) = flag
rs.Fields(1) = sdoj
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text6.Text
rs.Fields(4) = Combo3.Text
rs.Fields(5) = trngrab
rs.Fields(6) = tnamegrab
rs.Fields(7) = sfrom
rs.Fields(8) = sto
rs.Fields(9) = sclass
rs.Fields(10) = coach
rs.Fields(11) = sfare
rs.Fields(12) = grab
rs.Save
rs.Close

rs1.Open "Select * From Seats Where Train_Number ='" & trngrab & "' and DateOfJourney = '" & sdoj & "'", con, adOpenStatic, adLockOptimistic
If sclass = "AC 2-Tier" Then
    rs1.Fields(2).Value = rs1.Fields(2).Value - 3
ElseIf sclass = "AC 3-Tier" Then
    rs1.Fields(3).Value = rs1.Fields(3).Value - 3
ElseIf sclass = "Sleeper" Then
    rs1.Fields(4).Value = rs1.Fields(4).Value - 3
End If
sfare = sfare * 3
rs1.Update
rs1.Save
rs1.Close


End If

MsgBox "Your Ticket Has been Booked Successfully", vbInformation, "Congratulations!"
Menu.Show
End Sub

Private Sub Command5_Click()
If Combo4.Text <> "-Select Your Card Type-" Then
Frame3.Visible = True
Picture1.Visible = False
Command5.Visible = False
End If
End Sub

Private Sub Form_Load()
Module1.Connection
On Error Resume Next
Set ado = New ADODB.Recordset
ado.Open "Select * From tblUserInfo", con, adOpenStatic, adLockPessimistic
Label6.Caption = trngrab
rs.Open "Select * From Trains Where Train_Number='" & trngrab & "'", con, adOpenStatic, adLockReadOnly
Label8.Caption = rs.Fields(1)
Label10.Caption = sfrom
Label12.Caption = sto
Label14.Caption = sdoj
Label16.Caption = sclass
rs.Close
Option1.Value = True
End Sub

Private Sub Picture2_Click()
BookTicket.Hide
Menu.Show
Menu.Frame1.Visible = True
End Sub
