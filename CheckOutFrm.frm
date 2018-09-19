VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CheckOutFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check-Out Form"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9510
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
      Height          =   825
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   54
      Top             =   7320
      Width           =   3135
   End
   Begin VB.TextBox txtInvoice 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Frame xFrame6 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   2400
      TabIndex        =   40
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtTendered 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         IMEMode         =   3  'DISABLE
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   375
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   1005
         TabIndex        =   49
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   1065
         TabIndex        =   47
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   1650
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
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
         TabIndex        =   41
         Top             =   45
         Width           =   810
      End
      Begin VB.Image Image8 
         Height          =   315
         Left            =   0
         Picture         =   "CheckOutFrm.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5580
      End
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
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9255
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
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2280
         TabIndex        =   38
         Top             =   1920
         Width           =   3615
      End
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   855
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtdays 
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
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   615
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   360
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text2 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text3 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text4 
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text5 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1800
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3840
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   19
         Top             =   240
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
         Format          =   97452033
         CurrentDate     =   40013
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97452033
         CurrentDate     =   40013
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   3135
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
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check-in/Reserved by"
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
         TabIndex        =   39
         Top             =   1920
         Width           =   2010
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
         TabIndex        =   35
         Top             =   1200
         Width           =   780
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
         Left            =   2640
         TabIndex        =   34
         Top             =   1200
         Width           =   435
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
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check-in"
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
         TabIndex        =   32
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check-out"
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
         Left            =   3000
         TabIndex        =   31
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
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
         Left            =   6000
         TabIndex        =   30
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label12 
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
         Left            =   4440
         TabIndex        =   29
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label14 
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
         Left            =   4440
         TabIndex        =   28
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head count"
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
         TabIndex        =   27
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderStyle     =   2  'Dash
         X1              =   240
         X2              =   9120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head exceed"
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
         Left            =   2040
         TabIndex        =   26
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price exceed"
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
         Left            =   3960
         TabIndex        =   25
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addition amount"
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
         Left            =   6360
         TabIndex        =   24
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charges"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   240
         TabIndex        =   23
         Top             =   3240
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Services"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   240
         TabIndex        =   22
         Top             =   4080
         Width           =   1485
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   4320
         TabIndex        =   21
         Top             =   3240
         Width           =   915
      End
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "C&heck-out"
      Height          =   375
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "C&lear"
      Height          =   375
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label13 
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
      Left            =   1200
      TabIndex        =   55
      Top             =   7320
      Width           =   735
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
      Left            =   6840
      TabIndex        =   53
      Top             =   6480
      Width           =   465
   End
   Begin VB.Label Label25 
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
      Left            =   6240
      TabIndex        =   52
      Top             =   6480
      Width           =   540
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
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
      Left            =   720
      TabIndex        =   51
      Top             =   6840
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-Out Transaction"
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
      TabIndex        =   37
      Top             =   120
      Width           =   1950
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
      Left            =   5160
      TabIndex        =   36
      Top             =   6840
      Width           =   1065
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "CheckOutFrm.frx":0486
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "CheckOutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Dim clsData As New clsProducts
Dim clsInvoice As New clsInvoice
Private Sub cmdDelete_Click()
If Text8.Text = "" Then
MsgBox "No transaction available"
Else
xFrame6.Visible = True
txtTotalAmount.Text = Text8.Text
txtTendered.SetFocus
End If
End Sub
Private Sub cmdRefresh_Click()
Unload Me
CheckOutFrm.Show 1, ControlPanel
End Sub
Private Sub Command1_Click()
GetPrice
ListOfRoomsToCheckOutFrm.Show 1
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
xFrame6.Visible = False
End Sub
Public Sub GetPrice()
If mrs.State = adStateOpen Then mrs.Close
msql = "Select * from tblheadexceed"
mrs.Open msql, conn
Text4.Text = IIf(IsNull(mrs(0).Value), "", Format(mrs(0).Value, "##,##0.00"))
End Sub

Private Sub Command4_Click()
If CCur(txtTendered.Text) >= CCur(txtTotalAmount.Text) Then
clsData.updateRoomStatToAvailable txtrName
clsData.updateCheckOutBy CurrentUser, txtrCode
clsInvoice.AddInvoice txtInvoice, txtCustName, txtTotalAmount, txtTendered, txtChange, lblDate.Caption, CurrentUser
PrintInvoice
clsData.setCustomerToOpen txtCustName
Unload Me
Else
 MsgBox "Please enter payment greater than or equal to the total amount.", vbCritical, ""
End If
End Sub

Private Sub DTPicker2_Change()
If DTPicker2.Value < DTPicker1.Value Then
MsgBox "Invalid date", vbCritical, ""
ElseIf DTPicker2.Value = DTPicker1.Value Then
txtdays.Text = "1"
Else
txtdays.Text = DTPicker2.Value - DTPicker1.Value
End If
End Sub

Private Sub Form_Load()
txtInvoice.Text = clsInvoice.GetID
lblDate.Caption = Format(Date, "mm-dd-yyyy")
End Sub
Private Sub Text2_Change()
Text3.Text = Val(Text2.Text) - Val(txtcap.Text)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Private Sub Text3_Change()
If Text3.Text < 0 Then
Text3.Text = 0
End If
Text5.Text = Format(Val(Text3.Text) * CCur(Text4.Text), "##,##0.00")
End Sub



Private Sub Text5_Change()
Text6 = Format((CCur(txtRate.Text) * Val(txtdays.Text)) + CCur(Text5.Text), "##,##0.00")
End Sub
Private Sub Text6_Change()
Text8.Text = Format(CCur(Text6.Text) + CCur(Text7.Text), "##,##0.00")
End Sub
Private Sub Text7_Change()
Text7.Text = Format(CCur(Text7.Text), "##,##.00")
Text8.Text = Format(CCur(Text6.Text) + CCur(Text7.Text), "##,##0.00")
End Sub



Private Sub txtdays_Change()
Text6 = Format((CCur(txtRate.Text) * Val(txtdays.Text)) + CCur(Text5.Text), "##,##0.00")

End Sub

Private Sub txtTendered_Change()
On Error Resume Next
txtChange.Text = Format(txtTendered.Text - txtTotalAmount.Text, "##,##0.00")
If txtTendered.Text < 0 Then
txtTendered.Text = "0.00"
End If
txtTendered.Text = Format(txtTendered, "##,##0.00")
End Sub
Private Sub txtTendered_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
   End If
End Sub
Public Sub PrintInvoice()
Dim rsCompany As New ADODB.Recordset
If rsCompany.State = adStateOpen Then rsCompany.Close

rsCompany.Open " Select * From CompanySetup Where Index=1", conn

Set rpt_Invoice.DataSource = rsCompany

rpt_Invoice.Sections("Section2").Controls.Item("lblCName").Caption = rsCompany(1).Value
rpt_Invoice.Sections("Section2").Controls.Item("lblLocation").Caption = rsCompany(3).Value
rpt_Invoice.Sections("Section2").Controls.Item("lblContact").Caption = rsCompany(4).Value

rpt_Invoice.Sections("Section2").Controls.Item("lbluser").Caption = CurrentUser

rpt_Invoice.Sections("Section2").Controls.Item("lblInvoiceID").Caption = txtInvoice

rpt_Invoice.Sections("Section2").Controls.Item("lblTotal").Caption = txtTotalAmount.Text
rpt_Invoice.Sections("Section2").Controls.Item("lblTendered").Caption = txtTendered.Text
rpt_Invoice.Sections("Section2").Controls.Item("lblChange").Caption = txtChange.Text

rpt_Invoice.Sections("Section2").Controls.Item("lblName").Caption = txtCustName.Text
rpt_Invoice.Sections("Section2").Controls.Item("lblAdd").Caption = txtAdd.Text
rpt_Invoice.Sections("Section2").Controls.Item("lblCon").Caption = txtCon.Text
rpt_Invoice.Sections("Section2").Controls.Item("lblRoom").Caption = txtrName.Text

rpt_Invoice.Sections("Section2").Controls.Item("lblin").Caption = DTPicker1.Value
rpt_Invoice.Sections("Section2").Controls.Item("lblout").Caption = DTPicker2.Value

rpt_Invoice.Sections("Section2").Controls.Item("lblstay").Caption = txtdays.Text
rpt_Invoice.Show 1
End Sub

