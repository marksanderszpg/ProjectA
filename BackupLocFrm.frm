VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BackupLocFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Backup"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdsearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&browse"
      Height          =   345
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   855
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
      Left            =   2280
      Pattern         =   "*.bac"
      TabIndex        =   6
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtFilename 
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
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
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
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -360
      X2              =   10320
      Y1              =   2280
      Y2              =   2280
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
      TabIndex        =   5
      Top             =   3240
      Width           =   2100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   825
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
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "BackupLocFrm.frx":0000
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "BackupLocFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Dim DataFileName As String
Sub DataBackup()
Dim FS As New FileSystemObject, a As Integer
DataFileName = "" + Me.txtBackupPath.Text + "\" + Me.txtFilename.Text + ".library"
'ProgressBar1.Value = 0
FS.CopyFile App.Path & "\data.mdb", DataFileName, True
'FS.CopyFile App.Path & "\data.mdb", txtBackupPath & "\Data" & File1.ListCount & ".bac", True ' original
'FS.CopyFile App.Path & "\Data.mdb", txtBackupPath & " & Trim(txtFilename.Text) & " & File1.ListCount & ".bac", True
'For a = 1 To 10
'DoEvents
 '   Sleep 500
  '  ProgressBar1 = ProgressBar1 + 10
    'BackupFrm.Height = 6240
'Next
Set FS = Nothing
MsgBox "Backup successful", vbInformation, ""
'BackupFrm.Height = 5520
'ProgressBar1.Value = 0
Exit Sub
End Sub
Private Sub cmdCancel_Click()
Unload Me
'Unload bg
End Sub

Private Sub cmdOK_Click()
If Trim(txtBackupPath.Text) = "" Or txtFilename.Text = "" Then
MsgBox "Please enter backup path and filename.", vbCritical, ""
Else
DataBackup
Unload Me
'Unload bg
End If
End Sub

Private Sub cmdSearch_Click()
BackupFrm.Show 1
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Trim(txtBackupPath.Text) = "" Or txtFilename.Text = "" Then
MsgBox "Please enter backup path and filename.", vbCritical, ""
Else
DataBackup
Unload Me
'Unload bg
End If
End Sub

Private Sub Form_Load()
txtBackupPath.Text = App.Path & "\backup"
txtFilename.Text = Format(Date, "mm-dd-yyyy") & "DataBackUp"

End Sub
