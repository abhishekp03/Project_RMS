VERSION 5.00
Begin VB.MDIForm MDIFormRMS 
   BackColor       =   &H8000000C&
   Caption         =   "TrainLine :: Online Reservation Portal"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18270
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MDIFormRMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menAboutBR_Click()
About.Show
End Sub

Private Sub menAdmin_Click()
Admin.Show
End Sub

Private Sub menUpdate_Click()
Update.Show
End Sub
