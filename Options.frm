VERSION 5.00
Begin VB.Form Options 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Unlim 
      BackColor       =   &H80000012&
      Caption         =   "Check1"
      Height          =   255
      Left            =   3000
      MaskColor       =   &H00000000&
      TabIndex        =   19
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox PlayerSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Text            =   "100"
      Top             =   3705
      Width           =   735
   End
   Begin VB.TextBox HelicopterSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Text            =   "50"
      Top             =   3345
      Width           =   735
   End
   Begin VB.TextBox HelicopterAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Text            =   "7"
      Top             =   2925
      Width           =   735
   End
   Begin VB.TextBox PlayerHealth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "5"
      Top             =   2565
      Width           =   735
   End
   Begin VB.TextBox PlayerName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2978
      TabIndex        =   2
      Text            =   "Name"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Unlimited Helicopters"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   4020
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Save 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Close 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000007&
      Caption         =   "1 - 200"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Speed"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3660
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Caption         =   "1 - 200"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Helicopter Speed"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   3300
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "1 - 7"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "2 - 100"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Helicopter Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Health Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1178
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1418
      TabIndex        =   3
      Top             =   2115
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   26.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   1178
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COPS AND ROBBERS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   458
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MakeSound As Long
Private Sub Close_Click()
Me.Hide
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End Sub


Private Sub Form_Load()
Open App.Path & "\Files\Config.cfg" For Input As 1
Input #1, PlayerSpeedL, HelicopterSpeedL, PlayerNameL, PlayerHealthL, SoundL, HelicopterAmountL, UnlimL
PlayerSpeed.Text = PlayerSpeedL
HelicopterSpeed.Text = HelicopterSpeedL
PlayerName.Text = PlayerNameL
PlayerHealth.Text = PlayerHealthL
Unlim.Value = Unlim
HelicopterAmount.Text = HelicopterAmountL
Unlim.Value = UnlimL
Close #1
End Sub

Private Sub Label14_Click()
PlayerName.Text = "Name"
PlayerSpeed.Text = "100"
HelicopterSpeed.Text = "25"
PlayerHealth.Text = "90"
Unlim.Value = 0
HelicopterAmount.Text = "5"
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End Sub

Private Sub Save_Click()
On Error Resume Next

If PlayerSpeed.Text < 1 Or PlayerSpeed.Text > 200 Then
MsgBox "There was an error with the above information"
Exit Sub
ElseIf HelicopterSpeed.Text < 1 Or HelicopterSpeed.Text > 200 Then
MsgBox "There was an error with the above information"
Exit Sub
ElseIf PlayerHealth.Text < 1 Or PlayerHealth.Text > 100 Then
MsgBox "There was an error with the above information"
Exit Sub
ElseIf HelicopterAmount.Text < 1 Or HelicopterAmount.Text > 7 Then
MsgBox "There was an error with the above information"
Exit Sub
ElseIf HelicopterSpeed.Text < 1 Or HelicopterSpeed.Text > 200 Then
End If
If Err.Number <> 0 Then
    MsgBox "There was an error with the options you selected please try again. " & Err.Number & " " & Err.Description
    Exit Sub
End If

Open App.Path & "\Files\Config.cfg" For Output As #1
Write #1, PlayerSpeed, HelicopterSpeed, PlayerName, PlayerHealth, SoundOn, HelicopterAmount, Unlim.Value
Close #1

MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End Sub





