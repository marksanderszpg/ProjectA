VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Public Sub AddExpense(ID As String, expenseName As String, ecashout As String, netx As String, edate, qty As String, myprice As String, stotal As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * from tblExpenses"
rs.Open msql, conn
With rs
    .AddNew
    !ID = ID
    !expenseName = expenseName
    !ecashout = ecashout
    !netx = netx
    !edate = Date
    !qty = qty
    !myprice = myprice
    !stotal = stotal
    .Update
End With
MsgBox "New entry successfully saved to the record.", vbInformation
End Sub
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblExpenses.ID) AS MaxOfID" & _
" FROM tblExpenses;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function
Public Sub DisplayCustomer(lstDay As ListView)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
        msql = "Select * from tblexpenses order by id"
   
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
            lstItem.SubItems(8) = rs(7).Value
          '  lstItem.SubItems(9) = rs(8).Value
            rs.MoveNext
            Loop
End Sub
Public Sub DisplayExpense(lstCash As ListView, datefrom As Date, dateto As Date)
If rs.State = adStateOpen Then rs.Close
'msql = " SELECT Sum(tblInvoice.TotalPrice) AS SumOfTotalPrice, tblInvoice.InvoiceID, daterec" & _
     " From tblInvoice" & _
     " Where tblInvoice.DateRec >= #" & DateFrom & "#  And tblInvoice.Daterec <= #" & DateTo & "#" & _
     " GROUP BY tblInvoice.InvoiceID, daterec;"
     msql = " SELECT tblExpenses.expenseName, Sum(tblExpenses.myprice) AS SumOfmyprice, Sum(tblExpenses.qty) AS SumOfqty, Sum(tblExpenses.stotal) AS SumOfstotal, tblExpenses.edate" & _
" From tblExpenses " & _
" Where tblExpenses.edate >= #" & datefrom & "# and tblExpenses.edate <= #" & dateto & "#" & _
" GROUP BY tblExpenses.expenseName, tblExpenses.edate;"
     rs.Open msql, conn
lstCash.ListItems.Clear
Do While Not rs.EOF
a = a + 1
    Set lstItem = lstCash.ListItems.Add(, , a, 1, 1)
        lstItem.SubItems(1) = rs(0).Value
        lstItem.SubItems(2) = Format(rs(1).Value, "##,##0.00")
        lstItem.SubItems(3) = Format(rs(2).Value, "##,##0.00")
        lstItem.SubItems(4) = Format(rs(3).Value, "##,##0.00")
        lstItem.SubItems(5) = rs(4).Value
        rs.MoveNext
        Loop
End Sub


''print expenses of the hotel
Public Sub PrintExpenses(datefrom As Date, dateto As Date)
Dim rsCompany As New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
        msql = "Select * from tblexpenses  Where tblExpenses.edate >= #" & datefrom & "# and tblExpenses.edate <= #" & dateto & "#  order by id"
        '" Where tblInvoice.DateRec >=#" & MyDate & "# and tblInvoice.DateRec <= #" & DateTo & "#  " & _

 rs.Open msql, conn
 If rsCompany.State = adStateOpen Then rsCompany.Close
    rsCompany.Open " Select * From CompanySetup Where Index=1", conn
    
    Set rptExpense.DataSource = rs
    
        rptExpense.Sections("Section5").Controls.Item("lblCashier").Caption = CurrentUser
        rptExpense.Sections("Section4").Controls.Item("lblCName").Caption = rsCompany(1).Value
        rptExpense.Sections("Section4").Controls.Item("lblLocation").Caption = rsCompany(2).Value
        rptExpense.Sections("Section4").Controls.Item("lblContact").Caption = rsCompany(3).Value
        rptExpense.Sections("Section2").Controls.Item("lblDate").Caption = Date
rptExpense.Show 1

End Sub
