VERSION 5.00
Begin VB.Form GameOver 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ScoreAdder 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1440
   End
   Begin VB.Label TimeComment 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Back 
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Exit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Comment 
      BackStyle       =   0  'Transparent
      Caption         =   "Nice Try Marc"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Score1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1680
      TabIndex        =   2
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label KillType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   48
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "GameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MakeSound As Long
Private Sub Back_Click()
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
Me.Hide
Setup.Show
Setup.Timer1.Enabled = True
PlayerScore = 0
Score1.Caption = 0
FullTime = ""
MinTime = 0
SecTime = 0
Main.TimeTimer.Enabled = True
End Sub

Private Sub Exit_Click()
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End
End Sub

Private Sub Form_Load()
ScoreAdder.Enabled = True
End Sub

Private Sub ScoreAdder_Timer()
If Score1.Caption >= PlayerScore Then
    ScoreAdder.Enabled = False
Else
    Score1.Caption = Score1.Caption + 2
End If
Select Case PlayerScore
    Case Is < 300
        Comment.Caption = "You Stink " & PlayerName
    Case 300 To 400
        Comment.Caption = "Pretty Good " & PlayerName & ". NOT"
    Case 400 To 500
        Comment.Caption = "Umm No " & PlayerName
    Case 600 To 700
        Comment.Caption = "Ok Job, Watch The Helicopters " & PlayerName
    Case 700 To 900
        Comment.Caption = "Good Job " & PlayerName
    Case 1000 To 1200
        Comment.Caption = "Nice Job, Pretty Good " & PlayerName
    Case 1200 To 1600
        Comment.Caption = "Your Untouchable " & PlayerName
    Case 1700 To 1000000
        Comment.Caption = PlayerName & " Your Invincible!!"
End Select
TimeComment.Caption = "Your Time Was " & FullTime
Main.TimeTimer.Enabled = False

End Sub

