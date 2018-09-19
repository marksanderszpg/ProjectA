VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CustomerFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Form"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAdd 
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
      TabIndex        =   30
      Text            =   "."
      Top             =   7920
      Width           =   2655
   End
   Begin VB.ComboBox cbotype 
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
      Height          =   330
      ItemData        =   "CustomerFrm.frx":0000
      Left            =   5760
      List            =   "CustomerFrm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "-"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Main Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4695
      Left            =   -240
      TabIndex        =   22
      Top             =   7440
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   375
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4200
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         CausesValidation=   0   'False
         Height          =   3375
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5953
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1023
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Contact Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Gender"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "800"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Contact"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Address"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Date Recorded"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: double click the item to select"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
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
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11175
      Begin VB.TextBox txtage 
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
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cbogender 
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
         Height          =   330
         ItemData        =   "CustomerFrm.frx":001C
         Left            =   1800
         List            =   "CustomerFrm.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtCon 
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
         Left            =   5640
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtCustCode 
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
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtCustName 
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
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Format: Lastname, Firstname, MI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
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
         Left            =   4560
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         TabIndex        =   6
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1410
      End
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
      Left            =   6480
      TabIndex        =   0
      Top             =   2760
      Width           =   3495
   End
   Begin MSComctlLib.ListView lstCustomer 
      CausesValidation=   0   'False
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6165
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Contact Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Age"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Address"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Date Recorded"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   10920
      Top             =   2640
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
            Picture         =   "CustomerFrm.frx":0038
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   5760
      TabIndex        =   31
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   5760
      TabIndex        =   29
      Top             =   7080
      Width           =   435
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Customer (for Add-ons)"
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
      Left            =   8520
      TabIndex        =   28
      Top             =   6960
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   10440
      Picture         =   "CustomerFrm.frx":05D2
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view Customer's Information."
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
      TabIndex        =   11
      Top             =   360
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Maintenance"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Name"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "CustomerFrm.frx":0D75
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "CustomerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsData As New clsCustomer
Dim OldID As String

Private Sub cbotype_Change()
If cbotype.Text = "Add-on" Then
Label12.Visible = True
Text1.Visible = True
Frame2.Visible = True
clsData.DisplayCustomerMain ListView1
Else
Label12.Visible = False
Text1.Visible = False
End If
End Sub

Private Sub cbotype_Click()
If cbotype.Text = "Add-on" Then
Label12.Visible = True
Text1.Visible = True
Frame2.Visible = True
clsData.DisplayCustomerMain ListView1
Else
Label12.Visible = False
Text1.Visible = False
End If
End Sub

Private Sub cmdAdd_Click()
'If cmdAdd.Caption = "&Add" Then
'cmdAdd.Caption = "&Save"
'Clear
'txtCustName.SetFocus
'Else
    If CheckField = False Then
   ' cmdAdd.Caption = "&Add"
        clsData.AddCustomer txtCustCode, txtCustName, cbogender.Text, txtage, txtCon, txtAdd, cbotype.Text, Date & "-" & Time, Text1.Text
        Unload Me
        CustomerFrm.Show 1, ControlPanel
    End If
'End If
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
If vbYes = MsgBox("Delete record?", vbQuestion + vbYesNo, "Confirm Record Delete") Then
        clsData.DeleteCustomer txtCustCode
     Unload Me
    CustomerFrm.Show 1, ControlPanel
    End If
End Sub
Public Sub Clear()
txtCustCode = clsData.GetID
txtCustName = ""
'txtAdd = ""
txtCon = ""
cbogender.ListIndex = -1
txtage = ""
'cbotype.ListIndex = -1
End Sub

Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
cmdEdit.Caption = "&Update"
txtCustName.SetFocus
Else
    If CheckField = False Then
    cmdEdit.Caption = "&Edit"
    clsData.UpdateCustomer txtCustCode, txtCustName, cbogender.Text, txtage, txtCon, txtAdd, cbotype.Text, OldID, Text1.Text
    Unload Me
    CustomerFrm.Show 1, ControlPanel
    End If
End If
End Sub

Private Sub cmdRefresh_Click()
Clear
End Sub
Function CheckField() As Boolean
CheckField = True
    If Trim(txtCustName.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtCustName.SetFocus
    Exit Function
    ElseIf Trim(txtAdd.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtAdd.SetFocus
    Exit Function
    ElseIf Trim(txtCon.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtCon.SetFocus
    Exit Function
    ElseIf Trim(txtage.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtage.SetFocus
    Exit Function
    ElseIf cbogender.Text = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    cbogender.SetFocus
    Exit Function
  '  ElseIf cbotype.Text = "" Then
   ' MsgBox "All fields are required.", vbCritical, ""
    'cbotype.Text = ""
    'Exit Function
    End If
CheckField = False
End Function

Private Sub Command1_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
clsData.DisplayCustomer lstCustomer, ""
txtCustCode = clsData.GetID
cbotype.ListIndex = 0
End Sub





Private Sub ListView1_DblClick()
Text1.Text = ListView1.SelectedItem.ListSubItems(2).Text
Frame2.Visible = False
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
OldID = Item.ListSubItems(1)
txtCustCode = Item.ListSubItems(1)
txtCustName = Item.ListSubItems(2)
cbogender.Text = Item.ListSubItems(3)
txtage = Item.ListSubItems(4)
'txtAdd = Item.ListSubItems(6)
txtCon = Item.ListSubItems(5)
'cbotype.Text = Item.ListSubItems(7)
Text1.Text = Item.ListSubItems(9)
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtCon_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
clsData.DisplayCustomer lstCustomer, ""
Else
clsData.DisplayCustomer lstCustomer, txtSearch.Text
End If
End Sub
