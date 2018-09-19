VERSION 5.00
Begin VB.Form CompanyFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Information"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox tin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "111"
      Top             =   6000
      Width           =   5175
   End
   Begin VB.TextBox Contact 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox Loc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   5175
   End
   Begin VB.TextBox ComName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -600
      X2              =   10080
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3015
      Left            =   240
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Information"
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
      TabIndex        =   9
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Company Information"
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
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "CompanyFrm.frx":0000
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "CompanyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Setup As New clsProducts
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
    cmdEdit.Caption = "&Update"
    ComName.Locked = False
    Loc.Locked = False
    Contact.Locked = False
    tin.Locked = False
        Else
            If ValidateEntry = False Then
                cmdEdit.Caption = "&Edit"
                Setup.UpdateSetup ComName, Loc, Contact, tin
                ComName.Locked = False
               ' txtComplete.Locked = False
                Loc.Locked = False
                Contact.Locked = False
                tin.Locked = False
                Unload Me
             End If
End If
End Sub
Private Sub Form_Activate()
Setup.Setup ComName, Loc, Contact, tin
End Sub
Private Function ValidateEntry() As Boolean
ValidateEntry = True
    If Trim(ComName) = "" Then
        MsgBox "Don't leave the field blank", vbInformation, ""
        ComName = ""
        ComName.SetFocus
        Exit Function
    ElseIf Trim(Loc) = "" Then
        MsgBox "Don't leave the field blank", vbInformation, ""
        Loc = ""
        Loc.SetFocus
        Exit Function
     ElseIf Trim(Contact) = "" Then
        MsgBox "Don't leave the field blank", vbInformation, ""
        Contact = ""
        Contact.SetFocus
        Exit Function
     ElseIf Trim(tin) = "" Then
        MsgBox "Don't leave the field blank", vbInformation, ""
        tin = ""
        tin.SetFocus
     Exit Function
    
    End If
ValidateEntry = False
    
    
End Function
