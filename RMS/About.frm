VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      Height          =   1590
      Left            =   8160
      Picture         =   "About.frx":0000
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   3
      Top             =   2880
      Width           =   1590
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1590
      Left            =   5760
      Picture         =   "About.frx":6E49
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   2
      Top             =   2880
      Width           =   1590
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1635
      Left            =   3270
      Picture         =   "About.frx":D38C
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   2880
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1575
      Left            =   840
      Picture         =   "About.frx":15171
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   0
      Top             =   2880
      Width           =   1590
   End
   Begin VB.Label Label6 
      Caption         =   "Contributors:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":1C13C
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   9135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shivangi Prasad"
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
      Left            =   8160
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Soma Samanta"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muddassir Zafar"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   4680
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abhishek Pandey"
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
      Left            =   840
      TabIndex        =   4
      Top             =   4680
      Width           =   1845
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

