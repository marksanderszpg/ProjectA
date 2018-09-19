VERSION 5.00
Begin VB.Form HeadExceedChargeFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Head Exceed Charge"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   0
      Text            =   "200"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "HeadExceedChargeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Sub GetPrice()
If mrs.State = adStateOpen Then mrs.Close
msql = "Select * from tblheadexceed"
mrs.Open msql, conn
txtprice.Text = IIf(IsNull(mrs(0).Value), "", mrs(0).Value)
End Sub

Private Sub cmdDelete_Click()
UpdateHeadExceed
End Sub

Private Sub Form_Load()
GetPrice
End Sub
Sub UpdateHeadExceed()
 If mrs.State = adStateOpen Then mrs.Close
    msql = "UPDATE tblheadexceed SET headprice='" & txtprice & "'"
    mrs.Open msql, conn
    MsgBox "Setup Updated!.", vbInformation, ""
End Sub


Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
