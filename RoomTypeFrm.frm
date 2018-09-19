VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RoomTypeFrm 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Type"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
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
      TabIndex        =   13
      Top             =   5880
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
      TabIndex        =   12
      Top             =   5880
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
      TabIndex        =   11
      Top             =   5880
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
      TabIndex        =   10
      Top             =   5880
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
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      Begin VB.TextBox txtrtID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtrtName 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type Name"
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
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type ID"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1305
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
      Left            =   2280
      TabIndex        =   0
      Top             =   5280
      Width           =   2535
   End
   Begin MSComctlLib.ListView lstType 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
      _ExtentX        =   8281
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Room Type ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Room Type Name"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   4440
      Top             =   1800
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
            Picture         =   "RoomTypeFrm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Type Maintenance"
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
      TabIndex        =   9
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by type name"
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
      TabIndex        =   7
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -120
      X2              =   10560
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4080
      Picture         =   "RoomTypeFrm.frx":059A
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "RoomTypeFrm.frx":0E66
      Top             =   0
      Width           =   12960
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
      TabIndex        =   8
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "RoomTypeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsData As New clsRoomType
Dim OldID As String

Private Sub cmdAdd_Click()
'If cmdAdd.Caption = "&Add" Then
'cmdAdd.Caption = "&Save"
'Clear
'txtrtName.SetFocus
'Else
    If CheckField = False Then
  '  cmdAdd.Caption = "&Add"
        clsData.AddRoomType txtrtID, txtrtName
        Unload Me
        RoomTypeFrm.Show 1, ControlPanel
    End If
'End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Public Sub Clear()
txtrtID = clsData.GetID
txtrtName = ""
End Sub
Function CheckField() As Boolean
CheckField = True
    If Trim(txtrtName.Text) = "" Then
    MsgBox "All fields are required.", vbCritical, ""
    txtrtName.SetFocus
    Exit Function
    End If
CheckField = False
End Function

Private Sub cmdDelete_Click()
If vbYes = MsgBox("Delete record?", vbQuestion + vbYesNo, "Confirm Record Delete") Then
        clsData.DeleteRoomType txtrtID.Text
     Unload Me
    RoomTypeFrm.Show 1, ControlPanel
    End If
End Sub

Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
cmdEdit.Caption = "&Update"
txtrtName.SetFocus
Else
    If CheckField = False Then
    cmdEdit.Caption = "&Edit"
    clsData.UpdateRoomType txtrtID, txtrtName, OldID
    Unload Me
    RoomTypeFrm.Show 1, ControlPanel
    End If
End If
End Sub

Private Sub cmdRefresh_Click()
Clear
End Sub

Private Sub Form_Load()
clsData.DisplayRoomType lstType, ""
txtrtID = clsData.GetID
End Sub

Private Sub lstType_ItemClick(ByVal Item As MSComctlLib.ListItem)
OldID = Item.ListSubItems(1)
txtrtID = Item.ListSubItems(1)
txtrtName = Item.ListSubItems(2)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
clsData.DisplayRoomType lstType, ""
Else
clsData.DisplayRoomType lstType, txtSearch.Text
End If
End Sub
