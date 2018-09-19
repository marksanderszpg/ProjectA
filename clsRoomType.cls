VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoomType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Public Sub AddRoomType(rtID As String, rtName As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * from tblRoomType Where rtname='" & rtName & "'"
rs.Open msql, conn
If rs.RecordCount >= 1 Then
MsgBox "Duplicate record.", vbCritical, ""
Exit Sub
End If
With rs
    .AddNew
    !rtID = rtID
    !rtName = rtName
    .Update
End With
MsgBox "New entry successfully saved to the record.", vbInformation
End Sub
Public Sub UpdateRoomType(rtID As String, rtName As String, OldID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblRoomType set rtID='" & rtID & "',rtName='" & rtName & "' Where rtID='" & OldID & "'"
rs.Open msql, conn
MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Public Sub DeleteRoomType(rtID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Delete * From tblRoomType Where rtID='" & rtID & "'"
rs.Open msql, conn
MsgBox "Record(s) Deleted", vbInformation, ""
End Sub
Public Sub DisplayRoomType(lstDay As ListView, RoomSearch As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
    If RoomSearch <> "" Then
    msql = " SELECT tblRoomType.rtID, tblRoomType.rtName" & _
" From tblRoomType" & _
" Where tblRoomType.rtName like '" & RoomSearch & "%'" & _
" GROUP BY tblRoomType.rtID, tblRoomType.rtName" & _
" ORDER BY tblRoomType.rtID;"
    Else
    msql = " SELECT tblRoomType.rtID, tblRoomType.rtName" & _
" From tblRoomType" & _
" GROUP BY tblRoomType.rtID, tblRoomType.rtName" & _
" ORDER BY tblRoomType.rtID;"
    End If
 rs.Open msql, conn
   lstDay.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstDay.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = rs(0).Value
            lstItem.SubItems(2) = rs(1).Value
            rs.MoveNext
            Loop
End Sub
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblRoomType.rtID) AS MaxOfrtID" & _
" FROM tblRoomType;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function




