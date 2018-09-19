VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Public Sub AddRoom(rID As String, rName As String, rType As String, rCapacity As String, rRate As String, rDesc As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * from tblRoom Where rName='" & rName & "'"
rs.Open msql, conn
If rs.RecordCount >= 1 Then
MsgBox "Duplicate record found.", vbCritical, ""
Exit Sub
End If
With rs
    .AddNew
    !rID = rID
    !rName = rName
    !rType = rType
    !rCapacity = rCapacity
    !rRate = rRate
    !rDesc = rDesc
    .Update
End With
MsgBox "New entry successfully saved to the record.", vbInformation
End Sub
Public Sub UpdateRoom(rID As String, rName As String, rType As String, rCapacity As String, rRate As String, rDesc As String, OldID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblRoom set rID='" & rID & "',rName='" & rName & "',rtype='" & rType & "',rCapacity='" & rCapacity & "',rRate='" & rRate & "',rDesc='" & rDesc & "' Where rID='" & OldID & "'"
rs.Open msql, conn
MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Public Sub DeleteRoom(rID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Delete * From tblRoom Where rID='" & rID & "'"
rs.Open msql, conn
MsgBox "Record(s) Deleted", vbInformation, ""
End Sub
Public Sub DisplayRoom(lstDay As ListView, RoomSearch As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
    If RoomSearch <> "" Then
    msql = " SELECT tblRoom.rID, tblRoom.rName, tblRoom.rType, tblRoom.rCapacity, tblRoom.rRate, tblRoom.rDesc" & _
" From tblRoom" & _
" Where tblRoom.rName like '" & RoomSearch & "%'" & _
" GROUP BY tblRoom.rID, tblRoom.rName, tblRoom.rType, tblRoom.rCapacity, tblRoom.rRate, tblRoom.rDesc" & _
" ORDER BY tblRoom.rID;"
    Else
     msql = " SELECT tblRoom.rID, tblRoom.rName, tblRoom.rType, tblRoom.rCapacity, tblRoom.rRate, tblRoom.rDesc" & _
" From tblRoom" & _
" GROUP BY tblRoom.rID, tblRoom.rName, tblRoom.rType, tblRoom.rCapacity, tblRoom.rRate, tblRoom.rDesc" & _
" ORDER BY tblRoom.rID;"
    End If
 rs.Open msql, conn
   lstDay.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstDay.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = rs(0).Value
            lstItem.SubItems(2) = rs(1).Value
            lstItem.SubItems(3) = rs(2).Value
            lstItem.SubItems(4) = rs(3).Value
            lstItem.SubItems(5) = rs(4).Value
            lstItem.SubItems(6) = rs(5).Value
            rs.MoveNext
            Loop
End Sub
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblRoom.rID) AS MaxOfrID" & _
" FROM tblRoom;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function

Public Sub AddtoRoomStat(rID As String, rName As String, rCapacity As String, rRate As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * from tblRoomStat Where rname='" & rName & "'"
rs.Open msql, conn
If rs.RecordCount >= 1 Then
'MsgBox "Duplicate record found.", vbCritical, ""
Exit Sub
End If
With rs
    .AddNew
    !rID = rID
    !rName = rName
    !rmCapacity = rCapacity
    !rmRate = rRate
    .Update
End With
'MsgBox "New entry successfully saved to the record.", vbInformation
End Sub
Public Sub UpdatetoRoomStat(rID As String, rName As String, rCapacity As String, rRate As String, OldID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblRoomStat set rID='" & rID & "',rName='" & rName & "',rmCapacity='" & rCapacity & "',rmRate='" & rRate & "' Where rID='" & OldID & "'"
rs.Open msql, conn
'MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Public Sub DeleteRoomStat(rID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Delete * From tblRoomStat Where rID='" & rID & "'"
rs.Open msql, conn
'MsgBox "Record(s) Deleted", vbInformation, ""
End Sub
Public Sub DisplayRoomStat(lstDay As ListView, rmStat As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
If rmStat = "All" Then
msql = " SELECT tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" From tblRoomStat " & _
" GROUP BY tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" ORDER BY tblRoomStat.rID;"
Else
msql = " SELECT tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" From tblRoomStat " & _
" Where rmstat='" & rmStat & "'" & _
" GROUP BY tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" ORDER BY tblRoomStat.rID;"
End If
rs.Open msql, conn
   lstDay.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstDay.ListItems.Add(, , a, 1, 1)
            
            If rs(0).Value = "" Then
            lstItem.SubItems(1) = ""
            Else
             lstItem.SubItems(1) = rs(0).Value
             End If
            lstItem.SubItems(2) = rs(1).Value
            lstItem.SubItems(3) = rs(2).Value
            lstItem.SubItems(4) = rs(3).Value
            lstItem.SubItems(5) = rs(4).Value
            lstItem.SubItems(6) = rs(5).Value
            lstItem.SubItems(7) = rs(6).Value
            rs.MoveNext
            Loop


End Sub
Public Sub PrintRoomStat(rmStat As String)
Dim rsCompany As New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
If rmStat = "All" Then
msql = " SELECT tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" From tblRoomStat " & _
" GROUP BY tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" ORDER BY tblRoomStat.rID;"
Else
msql = " SELECT tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" From tblRoomStat " & _
" Where rmstat='" & rmStat & "'" & _
" GROUP BY tblRoomStat.cName, tblRoomStat.rID, tblRoomStat.rName, tblRoomStat.rmStat, tblRoomStat.rmCapacity, tblRoomStat.rmOccupants, tblRoomStat.rmRate" & _
" ORDER BY tblRoomStat.rID;"
End If
rs.Open msql, conn
If rsCompany.State = adStateOpen Then rsCompany.Close
    rsCompany.Open " Select * From CompanySetup Where Index=1", conn
    Set rptRooms.DataSource = rs
        rptRooms.Sections("Section5").Controls.Item("lblCashier").Caption = CurrentUser
        rptRooms.Sections("Section4").Controls.Item("lblCName").Caption = rsCompany(1).Value
        rptRooms.Sections("Section4").Controls.Item("lblLocation").Caption = rsCompany(2).Value
        rptRooms.Sections("Section4").Controls.Item("lblContact").Caption = rsCompany(3).Value
        rptRooms.Sections("Section2").Controls.Item("lblDate").Caption = Date
rptRooms.Show 1

End Sub

