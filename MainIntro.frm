VERSION 5.00
Begin VB.Form MainIntro 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Enter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Breif 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Point Of This Game is..."
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   135
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label GameType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Robber"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   975
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "MainIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MovedLeft As Integer
Private Sub Enter_Click()
Open App.Path & "\Files\Config.cfg" For Input As 1
Input #1, PlayerSpeedL, HelicopterSpeedL, PlayerNameL, PlayerHealthL, SoundOnL, HelicopterAmountL, UnlimL
PlayerSpeed = PlayerSpeedL
HelicopterSpeed = HelicopterSpeedL
PlayerName = PlayerNameL
PlayerHealth = PlayerHealthL
Unlim = UnlimL
HelicopterAmount = HelicopterAmountL
Close #1
HelicopterLeft = 0
PlayerOldSpeed = PlayerSpeed
PlayerScore = 0
Nitro1 = 0
Gun = 0
SpeedUp = False
SpeedNum = 100
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
Me.Hide
Main.Show
Main.Mover.Enabled = True
For I = 0 To 6
    Main.Helicopter(I).Left = Main.HelicopterPad(I).Left
    Main.Helicopter(I).Top = Main.HelicopterPad(I).Top
    Main.Helicopter(I).Visible = True
Next I
Main.Player.Top = Main.CarStart.Top
Main.Player.Left = Main.CarStart.Left
Main.Survival.Enabled = True
Nitro1 = 0
OldHelicopterSpeed = HelicopterSpeed
GotObjective = False
Main.GamePause.Caption = ""
PauseGame = False
Main.TimeTimer.Enabled = True
FullTime = ""
SecTime = 0
MinTime = 0

End Sub

