VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RoomStatusFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Status"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Add-on Customers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   1320
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   375
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3600
         Width           =   855
      End
      Begin MSComctlLib.ListView lstAddOns 
         CausesValidation=   0   'False
         Height          =   3135
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1023
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   855
   End
   Begin VB.ComboBox cbostat 
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
      ItemData        =   "RoomStatusFrm.frx":0000
      Left            =   2040
      List            =   "RoomStatusFrm.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin MSComctlLib.ListView lstType 
      CausesValidation=   0   'False
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7223
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Room ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Room Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Room Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Capacity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Occupant"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Room Rate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   7920
      Top             =   480
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
            Picture         =   "RoomStatusFrm.frx":0038
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click the item to view the Add-ons"
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
      TabIndex        =   6
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Room Status"
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
      TabIndex        =   3
      Top             =   720
      Width           =   1830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Status"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "RoomStatusFrm.frx":05D2
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "RoomStatusFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsData As New clsRoom
Private Sub cbostat_Change()
clsData.DisplayRoomStat lstType, cbostat.Text
End Sub
Private Sub cbostat_Click()
clsData.DisplayRoomStat lstType, cbostat.Text
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If lstType.ListItems.Count > 0 Then
clsData.PrintRoomStat cbostat.Text
Else
MsgBox "No record to print", vbInformation, ""
End If
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
End Sub

Private Sub Form_Load()
clsData.DisplayRoomStat lstType, "All"
cbostat.ListIndex = 0
End Sub

Public Sub ViewAddOns()
Dim lstItem As ListItem, a As Integer, mrs As New ADODB.Recordset
If mrs.State = adStateOpen Then mrs.Close

    msql = " SELECT tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
           " From tblCustomer" & _
           " Where tblCustomer.ctype='Add-on' and tblCustomer.cCusnameForAddOn='" & lstType.SelectedItem.ListSubItems(1).Text & "'" & _
    " GROUP BY tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType,mDateTime,cCusnameForAddOn" & _
    " ORDER BY tblCustomer.cID;"
    
 mrs.Open msql, conn
   lstAddOns.ListItems.Clear
   Do While Not mrs.EOF
    a = a + 1
        Set lstItem = lstAddOns.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = mrs!cName
          
            mrs.MoveNext
            Loop
End Sub

Private Sub lstType_DblClick()
'ViewAddOns
'Frame1.Visible = True
'ViewAddOns
End Sub
