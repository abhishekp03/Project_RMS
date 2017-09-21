Attribute VB_Name = "Module1"
Option Explicit

Public con As New ADODB.Connection
Public ado As New ADODB.Recordset
Public grab As String
Public fnamegrab As String
Public trngrab As String
Public tnamegrab As String
Public pnrgrab As String
Public sfrom As String
Public sto As String
Public sclass As String
Public sdoj As String
Public sfare As Integer

Public Sub Connection()
Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Database.mdb"
con.Open
End Sub
