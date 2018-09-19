VERSION 5.00
Begin VB.Form LoginFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2820
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.PictureBox cmdUsers 
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      ScaleHeight     =   300
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   11730
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   3240
      Picture         =   "LoginFrm.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   -360
      Picture         =   "LoginFrm.frx":0770
      Top             =   2520
      Width           =   12960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -240
      X2              =   10440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN FORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   -480
      Picture         =   "LoginFrm.frx":0D71
      Top             =   0
      Width           =   12960
   End
   Begin VB.Image Image5 
      Height          =   11700
      Left            =   -1080
      Picture         =   "LoginFrm.frx":1765
      Stretch         =   -1  'True
      Top             =   -7440
      Width           =   9225
   End
End
Attribute VB_Name = "LoginFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














Dim clsData As New clsUsers
Private Sub cmdCancel_Click()
Unload Me
'MainFrm.Show 1
End Sub
Private Sub cmdOK_Click()
Dim msg As String
    If cboUser.ListCount = 0 Then Unload Me: CurrentUser = "privilige"
        If cboUser.Text <> "" Then
            Set clsData = New clsUsers
                If True = clsData.Login(cboUser.Text, txtPass, msg) Then
                DoEvents
                CurrentUser = cboUser
                UserLog = clsData.UserLoginTime(CurrentUser, Time())
                clsData.GetUserPrivileges cboUser, UserTitle
                'addtemp
                Unload Me
                ControlPanel.Show
                Else
                CurrentUser = ""
                MsgBox msg, vbCritical, ""
                txtPass = ""
                End If
       Set clsData = Nothing
   End If
End Sub
Private Sub Form_Activate()
cboUser.SetFocus
End Sub
Private Sub Form_Load()
Set clsData = New clsUsers
If clsData.GetUsers(cboUser) = False Then Load Me
Set clsData = Nothing
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOK_Click
End If
End Sub

