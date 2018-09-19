VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Sub DisplayUsers(lstUsers As ListView)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
    sql = "Select * From Accounts  Order By Username"
   rs.Open sql, conn
   lstUsers.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstUsers.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = rs(0).Value
            lstItem.SubItems(2) = rs(2).Value
            lstItem.SubItems(5) = rs(1).Value
            rs.MoveNext
            Loop
End Sub
Function GetUsers(cboUser As ComboBox) As Boolean
'On Error GoTo hell
If rs.State = adStateOpen Then rs.Close
sql = "Select * From Accounts Order by Username"
rs.Open sql, conn
If Not (rs.EOF And rs.BOF) Then
cboUser.Clear
Do While Not rs.EOF
    GetUsers = True
    cboUser.AddItem rs.Fields(0).Value
    rs.MoveNext
    cboUser.ListIndex = 0
    Loop
 End If
 Exit Function
'hell:
'MsgBox Err.Description, vbCritical, ""
End Function
Public Function Login(Username As String, Password As String, msg As String) As Boolean
If rs.State = adStateOpen Then rs.Close
sql = "Select * From accounts where UserName='" & Username & "'"
rs.Open sql, conn
If (rs.EOF And rs.BOF) Then
DoEvents
msg = "The system cannot log you on.Please enter valid username and password."
Login = False
Else
    If rs.Fields(1).Value = Password Or rs.Fields(1).Value = "empty" Then
    DoEvents
    ControlPanel.Tag = rs.Fields(2).Value
    Login = True
    ElseIf IsNull(rs.Fields(0).Value) Then
    DoEvents
    Login = True
    Else
    Login = False
    msg = "Incorrect password"
    End If
End If
End Function
Sub AddUser(Username As String, Password As String, AccountType As String)

If rs.State = adStateOpen Then rs.Close
    sql = "Select * From Accounts Where username='" & Username & "'"
    rs.Open sql, conn
        If rs.RecordCount >= 1 Then
            MsgBox "Duplicate username found.Please enter another username.", vbInformation, ""
            Exit Sub
        End If
    With rs
        .AddNew
        !Username = Username
        !Password = Password
        !AccountType = AccountType
        
        .Update
        End With
        MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Sub UpdateUser(Username As String, Password As String, AccountType As String, Olduser As String)

If rs.State = adStateOpen Then rs.Close
    sql = "UPDATE Accounts SET Accounts.UserName='" & Username & "', Accounts.Password='" & Password & "', Accounts.AccountType='" & AccountType & "' WHERE (Accounts.UserName)='" & Olduser & "';"
    rs.Open sql, conn
    MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Sub DeleteUser(Username As String)
If rs.State = adStateOpen Then rs.Close
sql = "Delete * From Accounts Where username='" & Username & "'"
rs.Open sql, conn
MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Sub GetUserPrivileges(Username As String, AccountType As String)
If rs.State = adStateOpen Then rs.Close
sql = " SELECT Accounts.AccountType" & _
    " FROM Accounts" & _
    " Where username='" & Username & "'"
   rs.Open sql, conn
    AccountType = rs!AccountType
    
End Sub
Sub Initialize()
With mrs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
End With
End Sub
Function UserLoginTime(Username As String, LoginTime As Date) As Integer
If rs.State = adStateOpen Then rs.Close
    sql = " Select * From Userlog"
   rs.Open sql, conn
    With rs
        .AddNew
        !Username = Username
        !LoginTime = LoginTime
        .Update
   End With
   UserLoginTime = rs(0).Value
End Function
Function UserLogoutTime(Index As Integer, LogOutTime As String) As Integer
If rs.State = adStateOpen Then rs.Close
    sql = " Update Userlog set Userlog.LogoutTime=#" & LogOutTime & "# Where (userlog.logid)= " & Index & ""
rs.Open sql, conn
End Function
Sub DisplayUserLog(lstLog As ListView, mDate As Date)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
   sql = " Select * From Userlog Where (Userlog.Date)=#" & mDate & "#"
   rs.Open sql, conn
   lstLog.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstLog.ListItems.Add(, , a)
            lstItem.SubItems(1) = rs(1).Value
            lstItem.SubItems(2) = rs(2).Value
             If IsNull(rs(3).Value) Then
                lstItem.SubItems(3) = ""
             Else
                lstItem.SubItems(3) = rs(3).Value
             End If
            lstItem.SubItems(4) = rs(4).Value
            rs.MoveNext
            Loop
End Sub
