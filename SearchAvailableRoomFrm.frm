VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SearchAvailableRoomFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Available Room"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStat 
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
      Left            =   3720
      TabIndex        =   6
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
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
      Left            =   4800
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
      Left            =   1680
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
   End
   Begin MSComctlLib.ListView lstRoom 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
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
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Room name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Capacity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   6240
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
            Picture         =   "SearchAvailableRoomFrm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Available Room"
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
      Width           =   1980
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
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "SearchAvailableRoomFrm.frx":0420
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "SearchAvailableRoomFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msql As String
Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Sub SearchRoom(lstSearch As ListView, SearchValue As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
If Trim(SearchValue) = "" Then
msql = " SELECT tblRoomStat.rName, tblRoomStat.rmCapacity, tblRoomStat.rmRate" & _
" From tblRoomStat" & _
" Where rmstat='Available'" & _
" GROUP BY tblRoomStat.rName, tblRoomStat.rmCapacity, tblRoomStat.rmRate" & _
" ORDER BY tblRoomStat.rName;"
Else
msql = " SELECT tblRoomStat.rName, tblRoomStat.rmCapacity, tblRoomStat.rmRate" & _
" From tblRoomStat" & _
" Where rname like '" & SearchValue & "%' and rmstat='Available'" & _
" GROUP BY tblRoomStat.rName, tblRoomStat.rmCapacity, tblRoomStat.rmRate" & _
" ORDER BY tblRoomStat.rName;"
End If
rs.Open msql, conn
lstRoom.ListItems.Clear
Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstRoom.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = rs(0).Value
            lstItem.SubItems(2) = rs(1).Value
            lstItem.SubItems(3) = rs(2).Value
            rs.MoveNext
            Loop
End Sub
Private Sub cmdDelete_Click()
lstRoom_DblClick
End Sub
Private Sub Form_Load()
SearchRoom lstRoom, ""
End Sub
Private Sub lstRoom_DblClick()
If lstRoom.ListItems.Count > 0 Then
    If txtStat.Text = "Check-in" Then
    CheckInFrm.txtrName.Text = lstRoom.SelectedItem.ListSubItems(1).Text
    CheckInFrm.txtcap.Text = lstRoom.SelectedItem.ListSubItems(2).Text
    CheckInFrm.txtRate.Text = lstRoom.SelectedItem.ListSubItems(3).Text
    Unload Me
    Else
    RoomReserveFrm.txtrName.Text = lstRoom.SelectedItem.ListSubItems(1).Text
    RoomReserveFrm.txtcap.Text = lstRoom.SelectedItem.ListSubItems(2).Text
    RoomReserveFrm.txtRate.Text = lstRoom.SelectedItem.ListSubItems(3).Text
    Unload Me
    End If
Else
MsgBox "Record is empty.", vbInformation, ""
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
SearchRoom lstRoom, ""
Else
SearchRoom lstRoom, txtSearch.Text
End If
End Sub
