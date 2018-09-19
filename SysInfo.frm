VERSION 5.00
Begin VB.Form SysInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Info"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtsName 
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
      Height          =   1440
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1365
      Width           =   5415
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   0
      Picture         =   "SysInfo.frx":0000
      Top             =   3600
      Width           =   12960
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6360
      Picture         =   "SysInfo.frx":0601
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Displays the Information about the software"
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
      TabIndex        =   5
      Top             =   360
      Width           =   3195
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
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -120
      X2              =   10560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1365
      Width           =   1095
   End
   Begin VB.Label lblLocname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking System"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   915
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1470
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "SysInfo.frx":0ECD
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "SysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
Unload Me
End Sub
