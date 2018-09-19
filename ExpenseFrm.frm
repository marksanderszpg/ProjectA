VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ExpenseFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expenses Form"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Select Date"
      Height          =   2295
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   375
         Left            =   1920
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   360
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
         Left            =   2280
         TabIndex        =   25
         Top             =   1080
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
      Begin VB.Label Label11 
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
         Left            =   1200
         TabIndex        =   27
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label10 
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
         Left            =   1200
         TabIndex        =   26
         Top             =   1080
         Width           =   780
      End
   End
   Begin VB.TextBox txtReturn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6840
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4695
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtprice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtqty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1800
         TabIndex        =   13
         Text            =   "0"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txteName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtnet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtcOut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return"
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
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Code"
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
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Name"
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash-out"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   855
   End
   Begin MSComctlLib.ListView lstExpenses 
      CausesValidation=   0   'False
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Expense name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cash out"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Net"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   240
      Top             =   5520
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
            Picture         =   "ExpenseFrm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   2880
      X2              =   4800
      Y1              =   7320
      Y2              =   7320
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
      Left            =   2640
      TabIndex        =   22
      Top             =   6840
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Expense"
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
      Left            =   1560
      TabIndex        =   21
      Top             =   6960
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses Form"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1305
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "ExpenseFrm.frx":059A
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "ExpenseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsData As New clsExpense
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If CheckField = False Then
clsData.AddExpense txtCode, txteName, txtcOut, txtnet, Date, txtqty.Text, txtprice.Text, txttotal.Text
Unload Me
ExpenseFrm.Show 1, ControlPanel
End If
End Sub

Private Sub Command1_Click()
Frame2.Visible = True
End Sub
Function ComputeTotal()
Dim a As Integer, b As Integer, myTotal As Currency
 a = lstExpenses.ListItems.Count
For b = 1 To a
myTotal = Val(myTotal) + CCur(lstExpenses.ListItems(b).SubItems(8))
Next
lblTotal.Caption = Format(myTotal, "##,##0.00")
End Function

Private Sub Command2_Click()
If lstExpenses.ListItems.Count > 0 Then
clsData.PrintExpenses DTPicker1.Value, DTPicker2.Value
Else
MsgBox "No record to print", vbInformation, ""
End If

End Sub

Private Sub Command3_Click()
Frame2.Visible = False
End Sub

Private Sub DTPicker1_Change()
clsData.DisplayExpense lstExpenses, DTPicker1.Value, DTPicker2.Value
'ComputeTotal

End Sub

Private Sub DTPicker1_Click()
clsExpense.DisplayExpense lstExpenses, DTPicker1.Value, DTPicker2.Value
'ComputeTotal

End Sub

Private Sub DTPicker2_Change()
If DTPicker2.Value < DTPicker1.Value Then
MsgBox "Invalid", vbCritical, ""
DTPicker2.Value = DTPicker1.Value
Else
'clsData.DisplayCashOnHand lstCash, DTPicker1.Value, DTPicker2.Value
clsData.DisplayExpense lstExpenses, DTPicker1.Value, DTPicker2.Value
'ComputeTotal
'ComputeTotalExpense
End If
End Sub

Private Sub DTPicker2_Click()
If DTPicker2.Value < DTPicker1.Value Then
MsgBox "Invalid", vbCritical, ""
DTPicker2.Value = DTPicker1.Value
Else
'clsData.DisplayCashOnHand lstCash, DTPicker1.Value, DTPicker2.Value
clsData.DisplayExpense lstExpenses, DTPicker1.Value, DTPicker2.Value
'ComputeTotal
'ComputeTotalExpense
End If
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
txtCode.Text = clsData.GetID
clsData.DisplayCustomer lstExpenses
ComputeTotal
End Sub

Private Sub txtcOut_Change()
'If CCur(txtnet.Text) = CCur(txttotal.Text) Then
'txtnet.Text = "0.00"
'End If
If Trim(txtcOut.Text) = "" Then
txtcOut.Text = "0.00"
txtcOut.Text = Format(txtcOut, "##,##0.00")
End If
txtnet.Text = Format(CCur(txtcOut.Text) - CCur(txttotal.Text), "##,##0.00")
txtcOut.Text = Format(txtcOut, "##,##0.00")
txtnet.Text = Format(txtnet, "##,##0.00")

End Sub

Private Sub txtcOut_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtnet_Change()
If CCur(txtnet.Text) > CCur(txtcOut.Text) Then
    MsgBox "Cash return should not exceed the Cash-out amount", vbInformation, ""
    txtnet.Text = "0.00"
End If
'If CCur(txtnet.Text) = CCur(txttotal.Text) Then
'txtnet.Text = "0.00"
'End If
txtnet.Text = Format(txtnet, "##,##0.00")
'txtnet.Text = Format(CCur(txtcOut.Text) - CCur(txttotal.Text), "##,##0.00")
End Sub

Private Sub txtprice_Change()
If Trim(txtprice.Text) = "" Then
txtprice.Text = "0.00"
End If
txttotal.Text = Format(CCur(txtqty.Text) * CCur(txtprice.Text), "##,##0.00")
txtprice.Text = Format(txtprice, "##,##0.00")
'On Error GoTo errhandler:
'errhandler:
'txtprice.Text = "0.00"
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtqty_Change()
If Trim(txtqty.Text) = "" Then
txtqty.Text = "0"
End If
txttotal.Text = Format(CCur(txtqty.Text) * CCur(txtprice.Text), "##,##0.00")
'On Error GoTo errhandler:
'errhandler:
'txtqty.Text = "0"
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtReturn_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Function CheckField() As Boolean
CheckField = True
    If Trim(txtCode.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    'txtCustName.SetFocus
    Exit Function
    ElseIf Trim(txteName.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txteName.SetFocus
    Exit Function
    ElseIf Trim(txtcOut.Text) <= 0 Then
    MsgBox "Please enter cash-out amount.", vbCritical, ""
    txtcOut.SetFocus
    Exit Function
    ElseIf CCur(txtnet.Text) > CCur(txtcOut.Text) Then
    MsgBox "Cash return should not exceed the Cash-out amount", vbInformation, ""
    txtReturn.Text = "0.00"
    Exit Function
    End If
CheckField = False
End Function

Private Sub txttotal_Change()
txtnet.Text = Format(CCur(txtcOut.Text) - CCur(txttotal.Text), "##,##0.00")

End Sub
