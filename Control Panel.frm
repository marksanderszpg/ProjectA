VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ControlPanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel"
   ClientHeight    =   7245
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox OfficeMenu1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11640
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   1080
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   8760
      Top             =   5880
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
            Picture         =   "Control Panel.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8040
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0582
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgForm 
      Left            =   9240
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":07B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListLeft 
      Left            =   7440
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0F6C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListLV 
      Left            =   6840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1028
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1086
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":10E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1142
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":11FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":125C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1318
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1376
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":13D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1432
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1490
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":14EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":154C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":15AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1608
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1666
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":16C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1722
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":1780
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgQuickLaunch 
      Left            =   3120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":17DE
            Key             =   "rmstat"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":20B8
            Key             =   "user"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":2992
            Key             =   "reserve"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":326C
            Key             =   "creserve"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":3B46
            Key             =   "cin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":4420
            Key             =   "logoff"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":4CFA
            Key             =   "shutdown"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":55D4
            Key             =   "room"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":5EAE
            Key             =   "cout"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6788
            Key             =   "transaction"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7062
            Key             =   "customers"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":793C
            Key             =   "roomtype"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":8216
            Key             =   "set"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listQuickLaunch 
      CausesValidation=   0   'False
      Height          =   5655
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9975
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgQuickLaunch"
      SmallIcons      =   "imgQuickLaunch"
      ForeColor       =   128
      BackColor       =   16777215
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Control Panel.frx":8AF0
      NumItems        =   0
   End
   Begin VB.Frame xFrame6 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short-cut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   900
      End
      Begin VB.Image Image8 
         Height          =   315
         Left            =   0
         Picture         =   "Control Panel.frx":8C52
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5580
      End
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   2760
      TabIndex        =   7
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   7800
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   7800
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking System"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   585
      Left            =   4200
      TabIndex        =   3
      Top             =   6240
      Width           =   3330
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   0
      Picture         =   "Control Panel.frx":90D8
      Top             =   6960
      Width           =   12960
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "Control Panel.frx":96D9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
   Begin VB.Image Image2 
      Height          =   8850
      Left            =   3240
      Picture         =   "Control Panel.frx":A0CD
      Top             =   600
      Width           =   24000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log-Off"
      End
      Begin VB.Menu mnuAccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Data"
      End
      Begin VB.Menu mnuUserLog 
         Caption         =   "User Log"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Query"
      Begin VB.Menu mnuRooms 
         Caption         =   "Rooms"
      End
      Begin VB.Menu mnurType 
         Caption         =   "Room Type"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "Customers"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIn 
         Caption         =   "Check-in"
      End
      Begin VB.Menu mnuOUt 
         Caption         =   "Check-out"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReserve 
         Caption         =   "Reservation"
      End
      Begin VB.Menu mnuConfirm 
         Caption         =   "Confirm Reservation"
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncome 
         Caption         =   "Income Report"
      End
      Begin VB.Menu mnuExpense 
         Caption         =   "Expense Report"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStat 
         Caption         =   "Room Status"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Help"
      Begin VB.Menu MnuPO 
         Caption         =   "System Info"
      End
      Begin VB.Menu mnuSalesRPT 
         Caption         =   "Company Info"
      End
   End
End
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddQuickLaunchItems()
    
    listQuickLaunch.ListItems.Clear
    
    listQuickLaunch.ListItems.Add _
    , "RoomStat", "Room Status", "rmstat", "rmstat"

    listQuickLaunch.ListItems.Add _
    , "Reserve", "Reserve", "reserve", "reserve"

 
    listQuickLaunch.ListItems.Add _
    , "ConfirmReserve", "Confirm Reserve", "creserve", "creserve"
    
    listQuickLaunch.ListItems.Add _
    , "CheckIN", "Check-In", "cin", "cin"
    
     listQuickLaunch.ListItems.Add _
    , "CheckOut", "Check-Out", "cout", "cout"
    
    listQuickLaunch.ListItems.Add _
    , "Customers", "Customers", "user", "user"
    
    listQuickLaunch.ListItems.Add _
    , "Rooms", "Rooms", "room", "room"
    
    listQuickLaunch.ListItems.Add _
    , "RoomType", "Room Type", "roomtype", "roomtype"
    
    listQuickLaunch.ListItems.Add _
    , "Income", "Income", "transaction", "transaction"
    
   listQuickLaunch.ListItems.Add _
    , "Expense", "Expense", "customers", "customers"


    listQuickLaunch.ListItems.Add _
    , "HeadExceed", "Settings", "set", "set"
    
    ' listQuickLaunch.ListItems.Add _
    ', "PrintSchedule", "Class Schedule", "printschedule", "printschedule"
    
    
  ''''''''''''''''''''''''''''''''''''''''''''
    
    'listQuickLaunch.ListItems.Add _
   ' , "Logoff", "Log-off", "logoff", "logoff"
    
   ' listQuickLaunch.ListItems.Add _
  '  , "Shutdown", "Shutdown", "shutdown", "shutdown"

End Sub

Private Sub AddQuickLaunchItemsLimited()
    
    listQuickLaunch.ListItems.Clear
    
    listQuickLaunch.ListItems.Add _
    , "RoomStat", "Room Status", "rmstat", "rmstat"

    listQuickLaunch.ListItems.Add _
    , "Reserve", "Reserve", "reserve", "reserve"

 
    listQuickLaunch.ListItems.Add _
    , "ConfirmReserve", "Confirm Reserve", "creserve", "creserve"
    
    listQuickLaunch.ListItems.Add _
    , "CheckIN", "Check-In", "cin", "cin"
    
     listQuickLaunch.ListItems.Add _
    , "CheckOut", "Check-Out", "cout", "cout"
    
    listQuickLaunch.ListItems.Add _
    , "Customers", "Customers", "user", "user"
    
    listQuickLaunch.ListItems.Add _
    , "Rooms", "Rooms", "room", "room"
    
    listQuickLaunch.ListItems.Add _
    , "RoomType", "Room Type", "roomtype", "roomtype"
    
   ' listQuickLaunch.ListItems.Add _
    , "Income", "Income", "transaction", "transaction"
    
  ' listQuickLaunch.ListItems.Add _
    , "Expense", "Expense", "customers", "customers"


 '   listQuickLaunch.ListItems.Add _
  '  , "HeadExceed", "Settings", "set", "set"
    
    ' listQuickLaunch.ListItems.Add _
    ', "PrintSchedule", "Class Schedule", "printschedule", "printschedule"
    
    
  ''''''''''''''''''''''''''''''''''''''''''''
    
    'listQuickLaunch.ListItems.Add _
   ' , "Logoff", "Log-off", "logoff", "logoff"
    
   ' listQuickLaunch.ListItems.Add _
  '  , "Shutdown", "Shutdown", "shutdown", "shutdown"

End Sub
Private Sub Form_Activate()
If ControlPanel.Tag = "admin" Then
mnuAccounts.Enabled = True
mnuBackup.Enabled = True
mnuSettings.Enabled = True
mnuIncome.Enabled = True
mnuExpense.Enabled = True
AddQuickLaunchItems
Else
mnuAccounts.Enabled = False
mnuBackup.Enabled = False
mnuSettings.Enabled = False
mnuIncome.Enabled = False
mnuExpense.Enabled = False
AddQuickLaunchItemsLimited
End If
End Sub

Private Sub Form_Load()
If ControlPanel.Tag = "admin" Then
AddQuickLaunchItems
Else
AddQuickLaunchItemsLimited
End If
End Sub

Private Sub listQuickLaunch_DblClick()
Select Case listQuickLaunch.SelectedItem.Key
Case "RoomStat"
RoomStatusFrm.Show
Case "Reserve"
RoomReserveFrm.Show
Case "ConfirmReserve"
ConfirmReserveFrm.Show
Case "CheckIN"
CheckInFrm.Show
Case "Customers"
CustomerFrm.Show
Case "Rooms"
RoomFrm.Show
Case "RoomType"
RoomTypeFrm.Show
Case "HeadExceed"
HeadExceedChargeFrm.Show
Case "Logoff"
If vbYes = MsgBox("Log-off?", vbQuestion + vbYesNo, "") Then
Unload Me
FrmLogin.Show
End If
Case "Shutdown"
If vbYes = MsgBox("Shutdown?", vbQuestion + vbYesNo, "") Then
End
End If
Case "CheckOut"
CheckOutFrm.Show
Case "Income"
IncomeFrm.Show
Case "Expense"
ExpenseFrm.Show
End Select
End Sub

Private Sub mnuAccounts_Click()
AccountsFrm.Show
End Sub

Private Sub mnuBackup_Click()
BackupLocFrm.Show
End Sub

Private Sub mnuConfirm_Click()
ConfirmReserveFrm.Show
End Sub

Private Sub mnuCustomers_Click()
CustomerFrm.Show
End Sub

Private Sub mnuExit_Click()
If vbYes = MsgBox("Shutdown?", vbQuestion + vbYesNo, "") Then
End
End If
End Sub

Private Sub mnuExpense_Click()
ExpenseFrm.Show
End Sub

Private Sub mnuIn_Click()
CheckInFrm.Show
End Sub

Private Sub mnuIncome_Click()
IncomeFrm.Show
End Sub

Private Sub mnuLogOut_Click()
If vbYes = MsgBox("Log-off?", vbQuestion + vbYesNo, "") Then
Unload Me
LoginFrm.Show
End If
End Sub

Private Sub mnuOUt_Click()
CheckOutFrm.Show
End Sub

Private Sub MnuPO_Click()
SysInfo.Show
End Sub

Private Sub mnuReserve_Click()
RoomReserveFrm.Show
End Sub
Private Sub mnuRooms_Click()
RoomFrm.Show
End Sub
Private Sub mnurType_Click()
RoomTypeFrm.Show
End Sub

Private Sub mnuSalesRPT_Click()
CompanyFrm.Show
End Sub

Private Sub mnuSettings_Click()
HeadExceedChargeFrm.Show
End Sub

Private Sub mnuStat_Click()
RoomStatusFrm.Show
End Sub

Private Sub mnuUserLog_Click()
UserLogFrm.Show
End Sub
