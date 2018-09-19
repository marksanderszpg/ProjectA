VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblInvoice.InvoiceID) AS MaxOfInvoiceID" & _
" FROM tblInvoice;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function
Sub AddInvoice(InvoiceID As String, Customername As String, TotalPrice As String, Tendered As String, Change As String, Daterec As String, ProcessedBy As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * From tblInvoice Where InvoiceID='" & InvoiceID & "'"
rs.Open msql, conn
If rs.RecordCount >= 1 Then
MsgBox "Duplicate ID found.", vbCritical, ""
Exit Sub
End If
    With rs
        .AddNew
        !InvoiceID = InvoiceID
        !Customername = Customername
        !TotalPrice = TotalPrice
        !Tendered = Tendered
        !Change = Change
        !Daterec = Daterec
        !ProcessedBy = ProcessedBy
        .Update
     End With
End Sub

