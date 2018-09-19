VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CheckInFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check-In Form"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9525
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
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Top             =   7080
      Width           =   3135
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check-in"
      Height          =   375
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   855
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
      TabIndex        =   15
      Top             =   6840
      Width           =   2655
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
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   9255
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
         TabIndex        =   33
         Top             =   3120
         Width           =   4815
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
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   3360
         Width           =   2175
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
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   2520
         Width           =   2175
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
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   27
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   25
         Top             =   1920
         Width           =   615
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
         TabIndex        =   5
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   360
         Left            =   8640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   360
         Left            =   4440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   495
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   3135
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1200
         Width           =   3135
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
         TabIndex        =   35
         Top             =   240
         Width           =   615
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2895
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   1200
         Width           =   855
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   34
         Top             =   2520
         Width           =   915
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
         TabIndex        =   32
         Top             =   3360
         Width           =   1485
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
         TabIndex        =   31
         Top             =   2520
         Width           =   1440
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
         TabIndex        =   30
         Top             =   1920
         Width           =   1575
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
         TabIndex        =   28
         Top             =   1920
         Width           =   1125
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
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderStyle     =   2  'Dash
         X1              =   240
         X2              =   9120
         Y1              =   1800
         Y2              =   1800
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
         TabIndex        =   24
         Top             =   1920
         Width           =   1065
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
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1440
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
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   1035
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
         TabIndex        =   19
         Top             =   240
         Width           =   435
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
         TabIndex        =   18
         Top             =   240
         Width           =   945
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
         TabIndex        =   17
         Top             =   240
         Width           =   780
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
         Left            =   5040
         TabIndex        =   14
         Top             =   720
         Width           =   1095
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
         Left            =   7440
         TabIndex        =   13
         Top             =   1200
         Width           =   435
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
         Left            =   5040
         TabIndex        =   12
         Top             =   1200
         Width           =   780
      End
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
      Left            =   480
      TabIndex        =   40
      Top             =   7080
      Width           =   735
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
      Left            =   4680
      TabIndex        =   16
      Top             =   6840
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check-In Transaction"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "CheckInFrm.frx":0000
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "CheckInFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Dim clsData As New clsProducts
Private Sub cmdCancel_Click()
SearchCustomerFrm.txtStat = "Check-in"
SearchCustomerFrm.Show 1

End Sub

Private Sub cmdDelete_Click()
If CheckField = False Then
clsData.SaveToTBLCheckOut txtrCode, txtCustName, txtCon, txtAdd, txtrName, txtcap, txtRate, DTPicker1.Value, DTPicker2.Value, txtdays, CurrentUser, "", Text2, Text6, Text7, Text8, "Check-in", Text3
clsData.updateRoomStatToOccupied txtrName, txtCustName, Text2
clsData.setCustomerToClose txtCustName
Unload Me
End If
End Sub

Private Sub cmdRefresh_Click()
Unload Me
CheckInFrm.Show 1, ControlPanel
End Sub

Private Sub Command1_Click()
SearchAvailableRoomFrm.txtStat.Text = "Check-in"
SearchAvailableRoomFrm.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub DTPicker2_Click()
'DTPicker2_Change
If txtdays.Text < 0 Then
MsgBox "Invalid Date", vbCritical, ""
DTPicker2.Value = Date + 1
End If
txtdays = DTPicker2.Value - DTPicker1.Value
End Sub
Private Sub DTPicker2_Change()
DTPicker2_Click
If txtdays.Text < 0 Then
MsgBox "Invalid Date", vbCritical, ""
DTPicker2.Value = Date + 1
End If
txtdays = DTPicker2.Value - DTPicker1.Value
End Sub



Private Sub Form_Load()
txtrName = ""
txtCustName = ""
DTPicker1.Value = Date
DTPicker2.Value = Date + 1
txtdays = DTPicker2.Value - DTPicker1.Value
GetPrice
txtrCode.Text = clsData.GetID

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
Text7.Text = Format(CCur(Text7.Text), "##,##0.00")
Text8.Text = Format(CCur(Text6.Text) + CCur(Text7.Text), "##,##0.00")
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Public Sub GetPrice()
If mrs.State = adStateOpen Then mrs.Close
msql = "Select * from tblheadexceed"
mrs.Open msql, conn
Text4.Text = IIf(IsNull(mrs(0).Value), "", Format(mrs(0).Value, "##,##0.00"))
End Sub


Function CheckField() As Boolean
CheckField = True
    If Trim(txtCustName.Text) = "" Then
    MsgBox "Select Customer.", vbCritical, ""
    txtCustName.SetFocus
    Exit Function
    ElseIf Trim(txtrName.Text) = "" Then
    MsgBox "Select Room.", vbCritical, ""
    txtrName.SetFocus
    Exit Function
    ElseIf Trim(Text2.Text) = "" Then
    MsgBox "Enter headcount.", vbCritical, ""
    Text2.SetFocus
    Exit Function
    ElseIf Trim(Text2.Text) = 0 Then
    MsgBox "Enter headcount.", vbCritical, ""
    Text2.SetFocus
    Exit Function
    End If
CheckField = False
End Function

Private Sub txtdays_Change()
Text6 = Format((CCur(txtRate.Text) * Val(txtdays.Text)) + CCur(Text5.Text), "##,##0.00")
End Sub
