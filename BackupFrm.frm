VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BackupFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Back-up Database"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton jcbutton2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   2970
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3855
   End
   Begin VB.FileListBox File1 
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
      Height          =   2970
      Left            =   4680
      Pattern         =   "*.bac"
      TabIndex        =   5
      Top             =   5760
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
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
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtBackupPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   6840
      Width           =   4695
   End
   Begin VB.PictureBox cmdOK 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   3360
      Picture         =   "BackupFrm.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -3240
      X2              =   7440
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Backup"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing....pls wait."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   2100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   3720
      TabIndex        =   4
      Top             =   7320
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drive:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "BackupFrm.frx":086B
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "BackupFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub ccc_Click()

End Sub

Private Sub cmdOK_Click()
DataBackup
End Sub



Private Sub Command2_Click()
Unload Me
BackupFrm.Show 1, MainFrm
End Sub

Private Sub Dir1_Change()
Dir1_Click
End Sub

Private Sub Dir1_Click()
On Error GoTo hell:
File1.Path = Dir1.Path
BackupLocFrm.File1.Path = Dir1.Path
txtBackupPath = Dir1.Path
Exit Sub
hell:
MsgBox " Error drive is not accessible", vbInformation, ""
Drive1.Drive = "c:"
End Sub



Private Sub Drive1_Change()
On Error GoTo heaven
Dir1.Path = Drive1.Drive
Exit Sub
heaven:
MsgBox " Error drive is not accessible", vbInformation, ""
Drive1.Drive = "c:"
End Sub
Sub DataBackup()
Dim FS As New FileSystemObject, a As Integer
ProgressBar1.Value = 0
FS.CopyFile App.Path & "\data.mdb", txtBackupPath & "\Data" & File1.ListCount & ".bac", True
For a = 1 To 10
DoEvents
    Sleep 500
    ProgressBar1 = ProgressBar1 + 10
    BackupFrm.Height = 6240
Next
Set FS = Nothing
MsgBox "Backup successful", vbInformation, ""
BackupFrm.Height = 5520
ProgressBar1.Value = 0
Exit Sub
End Sub

Private Sub Form_Load()
Drive1.Drive = "c:"
End Sub

Private Sub jcbutton1_Click()

End Sub

Private Sub jcbutton2_Click()
If Trim(txtBackupPath.Text) = "" Then
MsgBox " Please enter backup location.", vbCritical, ""
Else
BackupLocFrm.txtBackupPath.Text = txtBackupPath.Text
Unload Me
End If
End Sub
