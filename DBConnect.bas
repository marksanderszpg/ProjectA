Attribute VB_Name = "DBConnect"
Public rs As New ADODB.Recordset
Public conn As New ADODB.Connection
Public sql As String
Public Con As String
Public CurrentUser As String
Public UserTitle As String
Public UserLog As Integer
Public rAdd As Boolean, rDelete As Boolean, rUpdate As Boolean, rPrint As Boolean
Public SuccessLogin As Boolean
Sub Main()
Con = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & App.Path & "\Datas.mdb;Jet OLEDB:Database Password=abc"
conn.Open Con
With rs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
End With
LoginFrm.Show
'ControlPanel.Show
End Sub
