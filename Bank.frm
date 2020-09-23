VERSION 5.00
Begin VB.Form Bank 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   Picture         =   "Bank.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer KillerMover 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   3840
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   5880
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   5640
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   4920
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   3960
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   4320
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   6360
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   5160
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   4320
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   4200
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   6000
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   5040
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   2880
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   5160
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   6480
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   5400
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   4680
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   3840
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   3480
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   3360
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   3120
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   3600
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   2760
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   2400
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2280
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2760
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   4920
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Killer 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   5640
      Top             =   3120
      Width           =   255
   End
   Begin VB.Image Player 
      Height          =   315
      Left            =   6720
      Picture         =   "Bank.frx":982BA
      Top             =   3120
      Width           =   315
   End
   Begin VB.Image Money 
      Height          =   1215
      Left            =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image Alarm 
      Height          =   1455
      Left            =   2040
      Top             =   2640
      Width           =   135
   End
End
Attribute VB_Name = "Bank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GoingDown(0 To 27)  As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        If Not Player.Top <= 0 Then
            Player.Top = Player.Top - 100
        End If
    Case vbKeyDown
        If Not Player.Top >= Bank.Height - Player.Height Then
            Player.Top = Player.Top + 100
        End If
    Case vbKeyRight
        If Not Player.Left >= Bank.Width - Player.Width Then
            Player.Left = Player.Left + 100
        End If
    Case vbKeyLeft
        If Not Player.Left <= 0 Then
            Player.Left = Player.Left - 100
        End If
End Select
If Player.Top < Alarm.Top + Alarm.Height And Player.Top > Alarm.Top - Player.Height Then
    If Player.Left < Alarm.Left + Alarm.Width And Player.Left > Alarm.Left - Player.Width Then
        KillerMover.Enabled = True
    End If
End If
For I = 1 To 17
    GoingDown(I) = False
Next I
End Sub

Private Sub KillerMover_Timer()
For I = 0 To Killer.Count
    If Not I = 19 Then
        If Killer(I).Top <= 0 Then
            GoingDown(I) = True
        ElseIf Killer(I).Top >= Bank.Height - Killer(I).Height Then
            GoingDown(I) = False
        End If
        If GoingDown(I) = True Then
            Killer(I).Top = Killer(I).Top + 100
        Else
            Killer(I).Top = Killer(I).Top - 100
        End If
    End If
Next I

End Sub
