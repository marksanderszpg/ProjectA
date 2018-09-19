VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AccountsFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Users"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
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
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5295
      Begin VB.TextBox txtVerify 
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtpass 
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
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtUser 
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
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cbotype 
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
         ItemData        =   "AccountsFrm.frx":0000
         Left            =   2160
         List            =   "AccountsFrm.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password"
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
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "User Type"
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView lstUser 
      CausesValidation=   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Account Type"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Password"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   360
      Top             =   3600
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
            Picture         =   "AccountsFrm.frx":001B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
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
            Picture         =   "AccountsFrm.frx":0079
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4680
      Picture         =   "AccountsFrm.frx":0613
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   10680
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users Information"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view and set user accounts and their type."
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
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   0
      Picture         =   "AccountsFrm.frx":0E12
      Top             =   0
      Width           =   12960
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   -240
      Picture         =   "AccountsFrm.frx":1806
      Top             =   6000
      Width           =   12960
   End
   Begin VB.Image Image4 
      Height          =   7920
      Left            =   120
      Picture         =   "AccountsFrm.frx":1E07
      Stretch         =   -1  'True
      Top             =   -1200
      Width           =   8685
   End
End
Attribute VB_Name = "AccountsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Dim clsData As New clsUsers
Dim OldID As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
'If rAdd = False Then MsgBox "Access Denied. The operation is not allowed", vbCritical, "Restriction": Exit Sub
    'If cmdAdd.Caption = "&Add" Then
       ' cmdAdd.Caption = "&Save"
        Lockfalse
        Clear
   ' Else
        If ValidateEntry = False Then
  '      cmdAdd.Caption = "&Add"
        txtUser.SetFocus
        Lockfalse
        clsData.AddUser txtUser, txtPass, cbotype.Text
        Clear
        Unload Me
        AccountsFrm.Show 1, ControlPanel
    End If
'End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
'If rDelete = False Then MsgBox "Access Denied. The operation is not allowed", vbCritical, "Restriction": Exit Sub
If txtUser = "" Then Exit Sub
    If txtUser = CurrentUser Then
        MsgBox "Can't delete your own profile.", vbCritical, ""
    Else
        If vbYes = MsgBox("Are you sure you want to delete this user" & vbCrLf & "Username: " & OldID, vbQuestion + vbYesNo, "") Then
            clsData.DeleteUser txtUser
    End If
    End If
    Unload Me
    AccountsFrm.Show 1, ControlPanel
End Sub


Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
cmdEdit.Caption = "&Update"
Lockfalse
txtUser.SetFocus
Else
    If ValidateEntry = False Then
        Lockfalse
        clsData.UpdateUser txtUser, txtPass, cbotype, OldID
        Unload Me
        AccountsFrm.Show 1, ControlPanel
    End If
End If
End Sub



Private Sub Command2_Click()
Unload Me
AccountsFrm.Show 1
End Sub

Private Sub Form_Activate()
Set clsData = New clsUsers
clsData.DisplayUsers lstUser
If CurrentUser = "sampleUser" Then
cmdDelete.Visible = False
cmdEdit.Visible = False
cmdAdd.Visible = False
Else
cmdDelete.Visible = True
cmdEdit.Visible = True
cmdAdd.Visible = True
End If
End Sub

Public Sub Lockfalse()
txtUser.Locked = False
txtPass.Locked = False
txtVerify.Locked = False
cbotype.Locked = False
End Sub
Public Sub Clear()
txtUser = ""
txtPass = ""
txtVerify = ""
cbotype.ListIndex = 0
End Sub
Public Function ValidateEntry() As Boolean
ValidateEntry = True
If Trim(txtUser) = "" Then
    MsgBox "Please enter username", vbInformation, ""
    Exit Function
ElseIf Trim(txtPass) <> Trim(txtVerify) Then
    MsgBox "Password mismatch.", vbInformation, ""
    Exit Function
 End If
ValidateEntry = False
End Function

Private Sub Form_Load()
If ControlPanel.Tag = "privilige" Then
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdAdd.Enabled = True
Else
If ControlPanel.Tag = "unprivilige" Then
cmdDelete.Visible = False
cmdEdit.Visible = False
cmdAdd.Visible = False
ControlPanel.Tag = ""
End If
End If
End Sub


Private Sub lstUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Dim a As Boolean, b As Boolean, c As Boolean, d As Boolean
OldID = Item.ListSubItems(1).Text
txtUser = Item.ListSubItems(1).Text
txtPass = Item.ListSubItems(5).Text
txtVerify = Item.ListSubItems(5).Text
cbotype.Text = Item.ListSubItems(2).Text
clsData.GetUserPrivileges txtUser, UserTitle
End Sub
