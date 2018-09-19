VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Sub Setup(txtCname As TextBox, txtLocation As TextBox, txtContact As TextBox, txtTIN As TextBox)
If mrs.State = adStateOpen Then mrs.Close
msql = "SELECT * FROM CompanySetup"
mrs.Open msql, conn
txtCname = IIf(IsNull(mrs(1).Value), "", mrs(1).Value)
'tx 'tComplete = IIf(IsNull(mrs(1).Value), "", mrs(2).Value)
txtLocation = IIf(IsNull(mrs(2).Value), "", mrs(2).Value)
txtContact = IIf(IsNull(mrs(3).Value), "", mrs(3).Value)
tTIN = IIf(IsNull(mrs(4).Value), "", mrs(4).Value)
End Sub
Sub UpdateSetup(txtCname As TextBox, txtLocation As TextBox, txtContact As TextBox, txtTIN As TextBox)
 If mrs.State = adStateOpen Then mrs.Close
    msql = "UPDATE CompanySetup SET CompanySetup.CName='" & txtCname & "', CompanySetup.Loc='" & txtLocation & "',CompanySetup.Contact='" & txtContact & "',CompanySetup.TIN='" & txtTIN & "' WHERE (((CompanySetup.Index)=1));"
    mrs.Open msql, conn
    MsgBox "Setup Updated!.", vbInformation, ""
End Sub
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblCheckOut.recordID) AS MaxOfcID" & _
" FROM tblCheckOut;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function
Public Sub SaveToTBLCheckOut(rID As String, cName As String, cContact As String, cAddress As String, rName As String, rCapacity As String, rRate As String, checkInTime As String, checkOutTime As String, daysStayed As String, checkInBy As String, checkOutBy As String, headCount As String, tCharges As String, tServices As String, totalCharges As String, inType As String, headExceed As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * from tblCheckOut Where recordid='" & rID & "'"
rs.Open msql, conn
If rs.RecordCount >= 1 Then
MsgBox "Duplicate record found", vbCritical, ""
Exit Sub
End If
With rs
    .AddNew
    !recordid = rID
    !cName = cName
    !cContact = cContact
    !cAddress = cAddress
    !rName = rName
    !rCapacity = rCapacity
    !rRate = rRate
    !checkInTime = checkInTime
    !checkOutTime = checkOutTime
    !daysStayed = daysStayed
    !checkInBy = checkInBy
    !checkOutBy = checkOutBy
    !headCount = headCount
    !tCharges = tCharges
    !tServices = tServices
    !totalCharges = totalCharges
    !inType = inType
    !headExceed = headExceed
    .Update
End With
MsgBox "Transaction Complete.", vbInformation
End Sub
Public Sub updateRoomStatToOccupied(rName As String, cName As String, rOccupants As String)
If rs.State = adStateOpen Then rs.Close
msql = "Update tblRoomStat Set cName='" & cName & "',rmOccupants='" & rOccupants & "', rmStat='Occupied' Where rname='" & rName & "'"
rs.Open msql, conn
End Sub
Public Sub updateRoomStatToReserved(rName As String, cName As String, rOccupants As String)
If rs.State = adStateOpen Then rs.Close
msql = "Update tblRoomStat Set cName='" & cName & "',rmOccupants='" & rOccupants & "', rmStat='Reserved' Where rname='" & rName & "'"
rs.Open msql, conn
End Sub
Public Sub updateRoomStatToAvailable(rName As String)
If rs.State = adStateOpen Then rs.Close
msql = "Update tblRoomStat Set rmStat='Available',cname='.' , rmOccupants = '0' Where rname='" & rName & "'"
rs.Open msql, conn
End Sub
Public Sub updateCheckOutBy(checkOutBy As String, recordid As String)
If rs.State = adStateOpen Then rs.Close
msql = "Update tblCheckOut Set checkOutBy='" & checkOutBy & "', transactionType='processed' Where recordID='" & recordid & "'"
rs.Open msql, conn
End Sub
Public Sub updateCheckOutByinReserved(checkOutBy As String, recordid As String)
If rs.State = adStateOpen Then rs.Close
msql = "Update tblCheckOut Set checkOutBy='" & checkOutBy & "' Where recordID='" & recordid & "'"
rs.Open msql, conn
End Sub
Public Sub setCustomerToClose(cName As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblCustomer Set hasTransaction='close' where cName='" & cName & "' "
rs.Open msql, conn
End Sub
Public Sub setCustomerToOpen(cName As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblCustomer Set hasTransaction='open' where cName='" & cName & "' "
rs.Open msql, conn
End Sub
Public Sub setInTypeToCheckIn(recordid As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblCheckout set intype='Check-in' where recordid='" & recordid & "'"
rs.Open msql, conn
End Sub

Public Sub cancelReserverationSetToCancelled(recordid As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblCheckout set intype='Cancelled' where recordid='" & recordid & "'"
rs.Open msql, conn
End Sub

