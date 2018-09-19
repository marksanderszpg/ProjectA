VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ListOfReserveRoomToBeConfirm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Reserved Room to be Confirm"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Select"
      Height          =   375
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Top             =   3960
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstRoom 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilList"
      SmallIcons      =   "ilList"
      ForeColor       =   12278016
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Record ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contact"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Room name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Capacity"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Check-In Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Check-Out Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Days Stayed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Check-In By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Head Count"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Charges"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Services"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Total Charge"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Head Exceed"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   6600
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListOfReserveRoomToBeConfirm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Reserved Rooms to Confirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2970
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   2400
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "ListOfReserveRoomToBeConfirm.frx":0420
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "ListOfReserveRoomToBeConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msql As String
Private Sub cmdCancel_Click()
Unload Me
End Sub
Public Sub GetRecordToConfirm(lstSearch As ListView, SearchValue As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
If Trim(SearchValue) = "" Then
msql = " SELECT tblCheckOut.recordID, tblCheckOut.cName, tblCheckOut.cContact, tblCheckOut.cAddress, tblCheckOut.rName, tblCheckOut.rCapacity, tblCheckOut.rRate, tblCheckOut.checkInTime, tblCheckOut.checkOutTime, tblCheckOut.daysStayed, tblCheckOut.checkInBy, tblCheckOut.headCount, tblCheckOut.tCharges, tblCheckOut.tServices, tblCheckOut.totalCharges, tblCheckOut.headExceed" & _
" From tblCheckOut" & _
" Where transactionType='unprocess' and inType='Reserved'" & _
" GROUP BY tblCheckOut.recordID, tblCheckOut.cName, tblCheckOut.cContact, tblCheckOut.cAddress, tblCheckOut.rName, tblCheckOut.rCapacity, tblCheckOut.rRate, tblCheckOut.checkInTime, tblCheckOut.checkOutTime, tblCheckOut.daysStayed, tblCheckOut.checkInBy, tblCheckOut.headCount, tblCheckOut.tCharges, tblCheckOut.tServices, tblCheckOut.totalCharges, tblCheckOut.headExceed" & _
" ORDER BY tblCheckOut.cName;"
Else
msql = " SELECT tblCheckOut.recordID, tblCheckOut.cName, tblCheckOut.cContact, tblCheckOut.cAddress, tblCheckOut.rName, tblCheckOut.rCapacity, tblCheckOut.rRate, tblCheckOut.checkInTime, tblCheckOut.checkOutTime, tblCheckOut.daysStayed, tblCheckOut.checkInBy, tblCheckOut.headCount, tblCheckOut.tCharges, tblCheckOut.tServices, tblCheckOut.totalCharges, tblCheckOut.headExceed" & _
" From tblCheckOut" & _
" Where tblCheckOut.cname like '" & SearchValue & "%' and transactionType='unprocess' and inType='Reserved'" & _
" GROUP BY tblCheckOut.recordID, tblCheckOut.cName, tblCheckOut.cContact, tblCheckOut.cAddress, tblCheckOut.rName, tblCheckOut.rCapacity, tblCheckOut.rRate, tblCheckOut.checkInTime, tblCheckOut.checkOutTime, tblCheckOut.daysStayed, tblCheckOut.checkInBy, tblCheckOut.headCount, tblCheckOut.tCharges, tblCheckOut.tServices, tblCheckOut.totalCharges, tblCheckOut.headExceed" & _
" ORDER BY tblCheckOut.cName;"
End If
rs.Open msql, conn
lstRoom.ListItems.Clear
Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstRoom.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = rs(0).Value 'record ID
            lstItem.SubItems(2) = rs(1).Value 'cName
            lstItem.SubItems(3) = rs(2).Value 'contact
            lstItem.SubItems(4) = rs(3).Value 'add
            lstItem.SubItems(5) = rs(4).Value 'rname
            lstItem.SubItems(6) = rs(5).Value 'capacity
            lstItem.SubItems(7) = rs(6).Value 'rate
            lstItem.SubItems(8) = rs(7).Value 'in time
            lstItem.SubItems(9) = rs(8).Value ' out time
            lstItem.SubItems(10) = rs(9).Value 'day stayed
            lstItem.SubItems(11) = rs(10).Value ' in by
            lstItem.SubItems(12) = rs(11).Value ' head count
            lstItem.SubItems(13) = rs(12).Value 'charges
            lstItem.SubItems(14) = rs(13).Value ' services
            lstItem.SubItems(15) = rs(14).Value ' total charges
            lstItem.SubItems(16) = rs(15).Value 'head exceed
            rs.MoveNext
            Loop
End Sub

Private Sub cmdDelete_Click()
lstRoom_DblClick
End Sub

Private Sub Form_Load()
GetRecordToConfirm lstRoom, ""
End Sub

Private Sub lstRoom_DblClick()
If lstRoom.ListItems.Count > 0 Then
ConfirmReserveFrm.txtrCode = lstRoom.SelectedItem.ListSubItems(1).Text
ConfirmReserveFrm.txtrName = lstRoom.SelectedItem.ListSubItems(5).Text
ConfirmReserveFrm.txtcap = lstRoom.SelectedItem.ListSubItems(6).Text
ConfirmReserveFrm.txtRate = lstRoom.SelectedItem.ListSubItems(7).Text
ConfirmReserveFrm.txtCustName = lstRoom.SelectedItem.ListSubItems(2).Text
ConfirmReserveFrm.txtCon = lstRoom.SelectedItem.ListSubItems(3).Text
ConfirmReserveFrm.txtAdd = lstRoom.SelectedItem.ListSubItems(4).Text
ConfirmReserveFrm.Text1 = lstRoom.SelectedItem.ListSubItems(11).Text
ConfirmReserveFrm.Text2 = lstRoom.SelectedItem.ListSubItems(12).Text
ConfirmReserveFrm.Text6 = lstRoom.SelectedItem.ListSubItems(13).Text
ConfirmReserveFrm.Text7 = lstRoom.SelectedItem.ListSubItems(14).Text
ConfirmReserveFrm.DTPicker1.Value = lstRoom.SelectedItem.ListSubItems(8).Text
ConfirmReserveFrm.DTPicker2.Value = lstRoom.SelectedItem.ListSubItems(9).Text
ConfirmReserveFrm.txtdays = lstRoom.SelectedItem.ListSubItems(10).Text
Unload Me
Else
MsgBox "Record is empty.", vbInformation, ""
End If
End Sub
Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
GetRecordToConfirm lstRoom, ""
Else
GetRecordToConfirm lstRoom, txtSearch.Text
End If
End Sub

