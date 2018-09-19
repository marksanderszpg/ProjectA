VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Public Sub AddCustomer(cID As String, cName As String, cGender As String, cAge As String, cContact As String, cAddress As String, cType As String, mDateTime As String, cCusnameForAddOn As String)
'close the recordset for new record to fetch
If rs.State = adStateOpen Then rs.Close

'declare your sql statement
msql = " Select * from tblCustomer Where cName='" & cName & "'"

'open the sql and establish connection
rs.Open msql, conn

'Check duplicate
If rs.RecordCount >= 1 Then
MsgBox "Duplicate record found", vbCritical, ""
Exit Sub
End If

'add record to database
With rs
    .AddNew
    !cID = cID
    !cName = cName
    !cGender = cGender
    !cAge = cAge
    !cContact = cContact
    !cAddress = cAddress
    !cType = cType
    !mDateTime = mDateTime
    !cCusnameForAddOn = cCusnameForAddOn
    .Update
End With
MsgBox "New entry successfully saved to the record.", vbInformation
End Sub

''sql statement used to update record
Public Sub UpdateCustomer(cID As String, cName As String, cGender As String, cAge As String, cContact As String, cAddress As String, cType As String, OldID As String, cCusnameForAddOn As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblCustomer set cID='" & cID & "',cName='" & cName & "'" & _
" ,cGender='" & cGender & "',cAge='" & cAge & "',cContact='" & cContact & "',cAddress='" & cAddress & "',ctype='" & cType & "',cCusnameForAddOn='" & cCusnameForAddOn & "' Where cID='" & OldID & "'"
rs.Open msql, conn
MsgBox "Record(s) Updated", vbInformation, ""
End Sub

''sql statement used to delete record
Public Sub DeleteCustomer(cID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Delete * From tblCustomer Where cID='" & cID & "'"
rs.Open msql, conn
MsgBox "Record(s) Deleted", vbInformation, ""
End Sub

''get and display record of customer from database
Public Sub DisplayCustomer(lstDay As ListView, CustomerSearch As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
    If CustomerSearch <> "" Then
    msql = " SELECT tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
           " From tblCustomer" & _
           " Where tblCustomer.cName like '" & CustomerSearch & "%'" & _
    " GROUP BY tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
    " ORDER BY tblCustomer.cID;"
    Else
     msql = " SELECT tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
           " From tblCustomer" & _
    " GROUP BY tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
    " ORDER BY tblCustomer.cID;"
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
            lstItem.SubItems(7) = rs(6).Value
            If IsNull(rs(7).Value) Then
            lstItem.SubItems(8) = ""
            Else
            lstItem.SubItems(8) = rs(7).Value
            End If
            '
            If IsNull(rs(8).Value) Then
            lstItem.SubItems(9) = ""
            Else
            lstItem.SubItems(9) = rs(8).Value
            End If
            rs.MoveNext
            Loop
End Sub
Public Sub DisplayCustomerMain(lstDay As ListView)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close

    msql = " SELECT tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
           " From tblCustomer" & _
           " Where tblCustomer.ctype='Main'" & _
    " GROUP BY tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
    " ORDER BY tblCustomer.cID;"
    
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
            lstItem.SubItems(7) = rs(6).Value
            If IsNull(rs(7).Value) Then
            lstItem.SubItems(8) = ""
            Else
            lstItem.SubItems(8) = rs(7).Value
            End If
            '
            If IsNull(rs(8).Value) Then
            lstItem.SubItems(9) = ""
            Else
            lstItem.SubItems(9) = rs(8).Value
            End If
            rs.MoveNext
            Loop
End Sub

'get the maximum value of the customer id
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblCustomer.cID) AS MaxOfcID" & _
" FROM tblCustomer;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function







