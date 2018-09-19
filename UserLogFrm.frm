VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form UserLogFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Log"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton jcbutton4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Filter Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5775
      Begin VB.CommandButton jcbutton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&OK"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97583105
         CurrentDate     =   40013
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User log as of:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
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
            Picture         =   "UserLogFrm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstUserLog 
      CausesValidation=   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   128
      BackColor       =   -2147483643
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Login Time"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Logout Time"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Information"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6360
      Picture         =   "UserLogFrm.frx":005E
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   -120
      Picture         =   "UserLogFrm.frx":092A
      Top             =   5160
      Width           =   12960
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "UserLogFrm.frx":0F2B
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "UserLogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim clsUserData As New clsUsers
'Dim clsPrint As New clsPrint



Private Sub DTPicker2_Change()
clsUserData.DisplayUserLog lstUserLog, DTPicker2.Value
End Sub

Private Sub Form_Activate()
clsUserData.DisplayUserLog lstUserLog, DTPicker2.Value
End Sub

Private Sub Form_Load()
DTPicker2.Value = Date
End Sub

Private Sub jcbutton1_Click()
If lstUserLog.ListItems.Count > 0 Then
'clsPrint.PrintUserLog DTPicker2.Value
Else
MsgBox " No data to print.", vbInformation, ""
End If
End Sub

Private Sub jcbutton3_Click()
Unload Me
UserLogFrm.Show 1, MainFrm
End Sub

Private Sub jcbutton4_Click()
Unload Me
End Sub
