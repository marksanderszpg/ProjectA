VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RoomFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Form"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4080
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
      TabIndex        =   22
      Top             =   4080
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
      TabIndex        =   21
      Top             =   4080
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
      TabIndex        =   20
      Top             =   4080
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
      TabIndex        =   19
      Top             =   4080
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
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4695
      Begin VB.TextBox txtcap 
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
         TabIndex        =   17
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox cboRtype 
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
         ItemData        =   "RoomFrm.frx":0000
         Left            =   1800
         List            =   "RoomFrm.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtRate 
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
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   795
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "RoomFrm.frx":001C
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtrCode 
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
      Begin VB.TextBox txtrName 
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
         Caption         =   "Capacity"
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
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type"
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
         Top             =   1320
         Width           =   1035
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
         TabIndex        =   14
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Name"
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
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Code"
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
         Top             =   240
         Width           =   1065
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
      Left            =   7200
      TabIndex        =   0
      Top             =   4200
      Width           =   3015
   End
   Begin MSComctlLib.ListView lstRoom 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   4920
      TabIndex        =   10
      Top             =   840
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Room ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Room Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Room Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Capacity"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   8760
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
            Picture         =   "RoomFrm.frx":0020
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   9480
      Picture         =   "RoomFrm.frx":05BA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view Room's Information."
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
      TabIndex        =   13
      Top             =   360
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Maintenance"
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
      TabIndex        =   12
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Room name"
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
      TabIndex        =   11
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "RoomFrm.frx":0E86
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "RoomFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsData As New clsRoom
Dim OldID As String
Private Sub cmdAdd_Click()
'If cmdAdd.Caption = "&Add" Then
'cmdAdd.Caption = "&Save"
'Clear
'txtrName.SetFocus
'Else
    If CheckField = False Then
    'cmdAdd.Caption = "&Add"
        clsData.AddRoom txtrCode, txtrName, cboRtype.Text, txtcap, txtRate, txtDesc
        clsData.AddtoRoomStat txtrCode, txtrName, txtcap, txtRate
        Unload Me
        RoomFrm.Show 1, ControlPanel
    End If
'End If
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
If vbYes = MsgBox("Delete record?", vbQuestion + vbYesNo, "Confirm Record Delete") Then
        clsData.DeleteRoom txtrCode
        clsData.DeleteRoomStat txtrCode
     Unload Me
    RoomFrm.Show 1, ControlPanel
    End If
End Sub
Public Sub Clear()
txtrCode = clsData.GetID
txtrName = ""
txtDesc = "."
txtRate = ""
cboRtype.ListIndex = -1
txtcap = ""
'cbotype.ListIndex = -1
End Sub
Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
cmdEdit.Caption = "&Update"
txtrName.SetFocus
Else
    If CheckField = False Then
    cmdEdit.Caption = "&Edit"
    clsData.UpdateRoom txtrCode, txtrName, cboRtype.Text, txtcap, txtRate, txtDesc, OldID
    clsData.UpdatetoRoomStat txtrCode, txtrName, txtcap, txtRate, OldID
    Unload Me
    RoomFrm.Show 1, ControlPanel
    End If
End If
End Sub

Private Sub cmdRefresh_Click()
Clear
End Sub
Function CheckField() As Boolean
CheckField = True
    If Trim(txtrName.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtrName.SetFocus
    Exit Function
    'ElseIf Trim(txtDesc.Text) = "" Then
   ' MsgBox "All fields are required.", vbCritical, ""
   ' txtDesc.SetFocus
   ' Exit Function
    ElseIf Trim(txtRate.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtRate.SetFocus
    Exit Function
    ElseIf Trim(txtcap.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtcap.SetFocus
    Exit Function
    ElseIf cboRtype.Text = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    cbogender.SetFocus
    Exit Function
    End If
CheckField = False
End Function

Private Sub Form_Load()
clsData.DisplayRoom lstRoom, ""
txtrCode = clsData.GetID
GetRoomType
End Sub





Private Sub lstRoom_ItemClick(ByVal Item As MSComctlLib.ListItem)
OldID = Item.ListSubItems(1)
txtrCode = Item.ListSubItems(1)
txtrName = Item.ListSubItems(2)
cboRtype.Text = Item.ListSubItems(3)
txtcap = Item.ListSubItems(4)
txtDesc = Item.ListSubItems(6)
txtRate = Item.ListSubItems(5)
'cbotype.Text = Item.ListSubItems(7)
End Sub

Private Sub txtcap_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
clsData.DisplayRoom lstRoom, ""
Else
clsData.DisplayRoom lstRoom, txtSearch.Text
End If
End Sub

Public Sub GetRoomType()
Dim msql As String
If rs.State = adStateOpen Then rs.Close
msql = "SELECT tblRoomType.rtName" & _
" From tblRoomType" & _
" GROUP BY tblRoomType.rtName" & _
" ORDER BY tblRoomType.rtName;"
rs.Open msql, conn
cboRtype.Clear
Do While Not rs.EOF
cboRtype.AddItem rs!rtName
rs.MoveNext
Loop
End Sub
