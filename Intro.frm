VERSION 5.00
Begin VB.Form Intro 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer IntroTimer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer IntroTimer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer IntroTimer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer IntroTimer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Skip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skip Intro"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5273
      TabIndex        =   2
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label EnterGame 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   5033
      TabIndex        =   1
      Top             =   4073
      Width           =   1935
   End
   Begin VB.Label ShowCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COPS N ROBBERS"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   170.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   7215
      Left            =   113
      TabIndex        =   0
      Top             =   893
      Width           =   11775
   End
   Begin VB.Image Helicopter 
      Height          =   3435
      Left            =   5760
      Picture         =   "Intro.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Image Explosion 
      Height          =   3135
      Left            =   9120
      Picture         =   "Intro.frx":0515
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   2925
   End
   Begin VB.Image Car 
      Height          =   2625
      Left            =   6840
      Picture         =   "Intro.frx":08A0
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   5295
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MakeSound As Long

Private Sub EnterGame_Click()
MakeSound = sndPlaySound("Media\Sound\gun.wav", &H1)
Me.Hide
Setup.Show

End Sub

Private Sub Form_Load()
Intro.Height = 30
IntroTimer1.Enabled = True
Car.Left = 0 - Car.Width
Helicopter.Left = 0 - Helicopter.Width
Explosion.Visible = False
ShowCaption.Visible = False
EnterGame.Visible = False

End Sub

Private Sub IntroTimer1_Timer()
Intro.Height = Intro.Height + 200
If Intro.Height >= 9000 Then
    IntroTimer1.Enabled = False
    IntroTimer2.Enabled = True
End If
End Sub

Private Sub IntroTimer2_Timer()
Car.Left = Car.Left + 200
If Car.Left >= 12000 - Car.Width Then
    IntroTimer2.Enabled = False
    IntroTimer3.Enabled = True
End If
End Sub

Private Sub IntroTimer3_Timer()
MakeSound = sndPlaySound("Media\Sound\gun.wav", &H1)
Helicopter.Left = Helicopter.Left + 200
If Helicopter.Left >= 12000 - Helicopter.Width Then
    MakeSound = sndPlaySound("Media\Sound\Explosion.wav", &H1)
    Explosion.Visible = True
    Car.Visible = False
    IntroTimer3.Enabled = False
    IntroTimer4.Enabled = True
End If
End Sub


Private Sub IntroTimer4_Timer()
Helicopter.Top = Helicopter.Top - 200
If Helicopter.Top <= 0 - Helicopter.Height Then
    ShowCaption.Visible = True
    EnterGame.Visible = True
    IntroTimer4.Enabled = False
End If
End Sub

Private Sub Skip_Click()
Me.Hide
IntroTimer1.Enabled = False
IntroTimer2.Enabled = False
IntroTimer3.Enabled = False
IntroTimer4.Enabled = False
Setup.Show
Setup.Timer1.Enabled = True

End Sub
