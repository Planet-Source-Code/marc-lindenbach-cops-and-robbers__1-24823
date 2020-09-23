VERSION 5.00
Begin VB.Form Setup 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   5160
   End
   Begin VB.Label SayHi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1448
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Car2 
      Height          =   465
      Left            =   1920
      Picture         =   "Setup.frx":0000
      Top             =   5400
      Width           =   870
   End
   Begin VB.Image Car 
      Height          =   465
      Left            =   1920
      Picture         =   "Setup.frx":04FD
      Top             =   2280
      Width           =   870
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1995
      TabIndex        =   7
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   1335
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Robber Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1275
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Survival Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   1268
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cram Productions"
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
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Beta 1.0"
      BeginProperty Font 
         Name            =   "Metrostyle"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   735
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MakeSound As Long

Private Sub Form_Activate()
If PlayerName = "Name" Or PlayerName = "" Then
    SayHi.Caption = "Please Choose A Name!"
Else
    SayHi.Caption = "Hello " & PlayerName & "!"
End If
Label2.Caption = "Beta " & App.Major & "." & App.Minor
End Sub

Private Sub Form_Load()
Open App.Path & "\Files\Config.cfg" For Input As 1
Input #1, PlayerSpeedL, HelicopterSpeedL, PlayerNameL, PlayerHealthL, SoundOnL, HelicopterAmountL, UnlimL
PlayerSpeed = PlayerSpeedL
HelicopterSpeed = HelicopterSpeedL
PlayerName = PlayerNameL
PlayerHealth = PlayerHealthL
Unlim = UnlimL
HelicopterAmount = HelicopterAmountL
Close #1

End Sub

Private Sub Label4_Click()
GameType = "Survival"
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
Me.Hide
Main.Show
MainIntro.Show
MainIntro.GameType.Caption = GameType
MainIntro.Breif.Caption = About.A2.Caption
MainIntro.Breif.Caption = MainIntro.Breif.Caption & "The controls are left key for left, right key for right, up key for up, down key for down, and space for nitro."
End Sub

Private Sub Label5_Click()
GameType = "Robber"
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
Me.Hide
Main.Show
MainIntro.Show
MainIntro.GameType.Caption = GameType
MainIntro.Breif.Caption = About.A4.Caption
MainIntro.Breif.Caption = MainIntro.Breif.Caption & "The controls are left key for left, right key for right, up key for up, down key for down, and space for nitro."
End Sub

Private Sub Label6_Click()
Options.Show
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End Sub

Private Sub Label7_Click()
About.Show
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End Sub

Private Sub Label8_Click()
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End
End Sub

Private Sub Timer1_Timer()
Car.Left = Car.Left + 50
If Car.Left >= 4680 Then
    Car.Left = -870
End If

Car2.Left = Car2.Left - 50
If Car2.Left <= -870 Then
    Car2.Left = 4680
End If
End Sub
