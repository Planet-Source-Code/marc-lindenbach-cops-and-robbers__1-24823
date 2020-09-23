VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Bar 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
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
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
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
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image DownArrow 
      Height          =   270
      Left            =   4680
      Picture         =   "About.frx":0000
      Top             =   5520
      Width           =   270
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Label A8 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0432
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   360
      TabIndex        =   12
      Top             =   6000
      Width           =   4095
   End
   Begin VB.Label A9 
      BackStyle       =   0  'Transparent
      Caption         =   "Walls:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label A7 
      BackStyle       =   0  'Transparent
      Caption         =   "Wepons:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label A6 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":04C2
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   4095
   End
   Begin VB.Label A5 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image UpArrow 
      Height          =   270
      Left            =   4680
      Picture         =   "About.frx":0580
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label A4 
      BackStyle       =   0  'Transparent
      Caption         =   "The point of the robber game is to get to the bank, take the money and then make it back to your base alive."
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label A3 
      BackStyle       =   0  'Transparent
      Caption         =   "Robber Game:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label A2 
      BackStyle       =   0  'Transparent
      Caption         =   "The point of the survival is to elimanate the helicopters using nitros, bazokas and health."
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label A1 
      BackStyle       =   0  'Transparent
      Caption         =   "Survival Game:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label A10 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":09B2
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   360
      TabIndex        =   10
      Top             =   7320
      Width           =   4095
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MakeSound As Long

Private Sub DownArrow_Click()
A1.Top = A1.Top - 100
A2.Top = A2.Top - 100
A3.Top = A3.Top - 100
A4.Top = A4.Top - 100
A5.Top = A5.Top - 100
A6.Top = A6.Top - 100
A7.Top = A7.Top - 100
A8.Top = A8.Top - 100
A9.Top = A9.Top - 100
A10.Top = A10.Top - 100
Bar.Top = Bar.Top + 76

End Sub

Private Sub Label6_Click()
Me.Hide
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End Sub

Private Sub UpArrow_Click()

A1.Top = A1.Top + 100
A2.Top = A2.Top + 100
A3.Top = A3.Top + 100
A4.Top = A4.Top + 100
A5.Top = A5.Top + 100
A6.Top = A6.Top + 100
A7.Top = A7.Top + 100
A8.Top = A8.Top + 100
A9.Top = A9.Top + 100
A10.Top = A10.Top + 100
Bar.Top = Bar.Top - 76
End Sub
