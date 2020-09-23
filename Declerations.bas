Attribute VB_Name = "Module1"
Public TopDirection As String
Public LeftDirection As String
Public HelicopterWidth As Integer
Public HelicopterHeight As Integer
Public PlayerWidth As Integer
Public PlayerHeight As Integer
Public PlayerSpeed As Integer
Public HelicopterSpeed As Integer
Public PlayerName As String
Public PlayerHealth As Integer
Public GotObjective As Boolean
Public PlayerLastTop As Integer
Public PlayerLastLeft As Integer
Public SoundOn As Integer
Public HelicopterAmount As Integer
Public GameType As String
Public PlayerScore As Long
Public PlayerOldSpeed As Integer
Public Nitro1 As Integer
Public Gun As Integer
Public SpeedUp As Boolean
Public SpeedNum As Integer
Public HelicopterLeft As Integer
Public ScoreAmount As Integer
Public OldHelicopterSpeed As Integer
Public PauseGame As Boolean
Public Unlim As Integer
Public FullTime As String
Public SecTime As Integer
Public MinTime As Integer

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

