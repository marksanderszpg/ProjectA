VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form IncomeFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income (Cash on Hand) Form"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotal 
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtTExpense 
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   2280
      Top             =   120
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
            Picture         =   "IncomeFrm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstCash 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cash"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Recorded"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97452033
      CurrentDate     =   40207
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97452033
      CurrentDate     =   40207
   End
   Begin MSComctlLib.ListView lstExpense 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   7080
      TabIndex        =   14
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Expense Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Recorded"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   7080
      TabIndex        =   17
      Top             =   780
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   6960
      Picture         =   "IncomeFrm.frx":059A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      Height          =   4215
      Left            =   6960
      Top             =   720
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   10080
      X2              =   12000
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expense"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Top             =   4440
      Width           =   1950
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   9840
      TabIndex        =   15
      Top             =   4320
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   120
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Income"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   13
      Top             =   800
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Left            =   5400
      TabIndex        =   12
      Top             =   1920
      Width           =   780
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   5760
      Picture         =   "IncomeFrm.frx":0A20
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash On Hand"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Left            =   5400
      TabIndex        =   4
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   4320
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Income"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   4440
      Width           =   1845
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   3240
      X2              =   5160
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "IncomeFrm.frx":11C2
      Top             =   0
      Width           =   12960
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   120
      Picture         =   "IncomeFrm.frx":1BB6
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "IncomeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsData As New clsPrint
Dim clsExpense As New clsExpense
Function ComputeTotal()
Dim a As Integer, b As Integer, myTotal As Currency
 a = lstCash.ListItems.Count
For b = 1 To a
myTotal = Val(myTotal) + CCur(lstCash.ListItems(b).SubItems(2))
Next
lblTotal.Caption = Format(myTotal, "##,##0.00")
End Function
Function ComputeTotalExpense()
Dim a As Integer, b As Integer, myTotal As Currency
 a = lstExpense.ListItems.Count
For b = 1 To a
myTotal = Val(myTotal) + CCur(lstExpense.ListItems(b).SubItems(4))
Next
Label4.Caption = Format(myTotal, "##,##0.00")
End Function


Private Sub cmdDelete_Click()
If lstCash.ListItems.Count <= 0 Then
MsgBox "No data to print", vbCritical, ""
Else
clsData.PrintCash DTPicker1.Value, lblTotal.Caption, Label4.Caption, txttotal.Text, DTPicker2.Value
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
clsData.DisplayCashOnHand lstCash, DTPicker1.Value, DTPicker2.Value
clsExpense.DisplayExpense lstExpense, DTPicker1.Value, DTPicker2.Value
ComputeTotal
ComputeTotalExpense
End Sub

Private Sub DTPicker1_Click()
clsData.DisplayCashOnHand lstCash, DTPicker1.Value, DTPicker2.Value
clsExpense.DisplayExpense lstExpense, DTPicker1.Value, DTPicker2.Value
ComputeTotal
ComputeTotalExpense
End Sub

Private Sub DTPicker2_Change()
If DTPicker2.Value < DTPicker1.Value Then
MsgBox "Invalid", vbCritical, ""
DTPicker2.Value = DTPicker1.Value
Else
clsData.DisplayCashOnHand lstCash, DTPicker1.Value, DTPicker2.Value
clsExpense.DisplayExpense lstExpense, DTPicker1.Value, DTPicker2.Value
ComputeTotal
ComputeTotalExpense
End If
End Sub

Private Sub DTPicker2_Click()
If DTPicker2.Value < DTPicker1.Value Then
MsgBox "Invalid", vbCritical, ""
DTPicker2.Value = DTPicker1.Value
Else
clsData.DisplayCashOnHand lstCash, DTPicker1.Value, DTPicker2.Value
clsExpense.DisplayExpense lstExpense, DTPicker1.Value, DTPicker2.Value
ComputeTotal
ComputeTotalExpense
End If
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
'clsData.DisplayCashOnHand lstCash, DTPicker1.Value
'ComputeTotal
'GetTotalExpense
End Sub


Public Sub GetTotalExpense()
Dim localRs As New ADODB.Recordset, mysql As String
If localRs.State = adStateOpen Then localRs.Close
mysql = "SELECT Sum(tblExpenses.netx) AS SumOfnetx FROM tblExpenses;"
localRs.Open mysql, conn
'Do While localRs.EOF
txtTExpense.Text = IIf(IsNull(localRs(0).Value), "", localRs(0).Value)
'localRs.MoveNext
'Loop


End Sub

Private Sub Label4_Change()
txttotal.Text = Format(CCur(lblTotal.Caption) - CCur(Label4.Caption), "##,##0.00")

End Sub

Private Sub Label4_Click()
txttotal.Text = Format(CCur(lblTotal.Caption) - CCur(Label4.Caption), "##,##0.00")

End Sub

Private Sub lblTotal_Change()
txttotal.Text = Format(CCur(lblTotal.Caption) - CCur(Label4.Caption), "##,##0.00")

End Sub

Private Sub lblTotal_Click()
txttotal.Text = Format(CCur(lblTotal.Caption) - CCur(Label4.Caption), "##,##0.00")

End Sub

Private Sub txtTExpense_Change()
txttotal.Text = Format(CCur(lblTotal.Caption) - CCur(txtTExpense.Text), "##,##0.00")

End Sub
