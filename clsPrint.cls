VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Public Sub DisplayCashOnHand(lstCash As ListView, DateFrom As Date, DateTo As Date)
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Sum(tblInvoice.TotalPrice) AS SumOfTotalPrice, tblInvoice.InvoiceID, daterec" & _
     " From tblInvoice" & _
     " Where tblInvoice.DateRec >= #" & DateFrom & "#  And tblInvoice.Daterec <= #" & DateTo & "#" & _
     " GROUP BY tblInvoice.InvoiceID, daterec;"
     rs.Open msql, conn
lstCash.ListItems.Clear
Do While Not rs.EOF
a = a + 1
    Set lstItem = lstCash.ListItems.Add(, , a, 1, 1)
        lstItem.SubItems(1) = rs(1).Value
        lstItem.SubItems(2) = Format(rs(0).Value, "##,##0.00")
        lstItem.SubItems(3) = rs(2).Value
        rs.MoveNext
        Loop
End Sub
Public Sub PrintCash(MyDate As Date, TotalPrice As String, TotalExpense As String, NetTotal As String, DateTo As Date)
Dim rsCompany As New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Sum(tblInvoice.TotalPrice) AS SumOfTotalPrice, tblInvoice.InvoiceID, daterec" & _
     " From tblInvoice" & _
     " Where tblInvoice.DateRec >=#" & MyDate & "# and tblInvoice.DateRec <= #" & DateTo & "#  " & _
     " GROUP BY tblInvoice.InvoiceID,daterec;"
     rs.Open msql, conn
    If rsCompany.State = adStateOpen Then rsCompany.Close
    rsCompany.Open " Select * From CompanySetup Where Index=1", conn
    Set rptCash.DataSource = rs
        rptCash.Sections("Section5").Controls.Item("lblCashier").Caption = CurrentUser
        rptCash.Sections("Section4").Controls.Item("lblCName").Caption = rsCompany(1).Value
        rptCash.Sections("Section4").Controls.Item("lblLocation").Caption = rsCompany(2).Value
        rptCash.Sections("Section4").Controls.Item("lblContact").Caption = rsCompany(3).Value
        rptCash.Sections("Section5").Controls.Item("lblTotal").Caption = TotalPrice
        rptCash.Sections("Section5").Controls.Item("lblExpense").Caption = TotalExpense
        rptCash.Sections("Section5").Controls.Item("lblNet").Caption = NetTotal

rptCash.Sections("Section2").Controls.Item("lblDate").Caption = MyDate
rptCash.Show 1
End Sub

