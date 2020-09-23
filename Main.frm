VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   2280
   End
   Begin VB.Timer Survival 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   1800
   End
   Begin VB.Timer Mover 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1200
      Top             =   1320
   End
   Begin VB.Image Flowers 
      Height          =   915
      Index           =   5
      Left            =   0
      Picture         =   "Main.frx":0000
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label GameTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time 00:00"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   5760
      Picture         =   "Main.frx":2EF6
      Top             =   8640
      Width           =   315
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   6
      Left            =   5400
      Picture         =   "Main.frx":3378
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   5
      Left            =   5400
      Picture         =   "Main.frx":388D
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   4
      Left            =   0
      Picture         =   "Main.frx":3DA2
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   3
      Left            =   0
      Picture         =   "Main.frx":42B7
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   2
      Left            =   2160
      Picture         =   "Main.frx":47CC
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   1
      Left            =   0
      Picture         =   "Main.frx":4CE1
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Helicopter 
      Height          =   915
      Index           =   0
      Left            =   3240
      Picture         =   "Main.frx":51F6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label GamePause 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Paused"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   72
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   2693
      TabIndex        =   14
      Top             =   3833
      Width           =   6615
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   56
      Left            =   9720
      Picture         =   "Main.frx":570B
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nitro"
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
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   8640
      Width           =   975
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   4800
      Picture         =   "Main.frx":8601
      Top             =   8640
      Width           =   285
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mine"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   8640
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3480
      Picture         =   "Main.frx":8A3F
      Top             =   8640
      Width           =   600
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gun/Ammo"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   8640
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2040
      Picture         =   "Main.frx":9279
      Top             =   8640
      Width           =   285
   End
   Begin VB.Label f 
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   8640
      Width           =   975
   End
   Begin VB.Image CarStart 
      Height          =   495
      Left            =   8760
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Guns 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11400
      TabIndex        =   9
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Wepons:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10800
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Nitros 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   7
      Top             =   780
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nitros:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Score 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   5
      Top             =   420
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10800
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Health 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   11280
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   6
      Left            =   5400
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   5
      Left            =   0
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   4
      Left            =   5400
      Top             =   960
      Width           =   975
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   3
      Left            =   0
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   2
      Left            =   2160
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   1
      Left            =   3240
      Top             =   0
      Width           =   975
   End
   Begin VB.Image HelicopterPad 
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Quit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Back 
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10800
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   10800
      X2              =   12000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   10800
      X2              =   10800
      Y1              =   0
      Y2              =   1800
   End
   Begin VB.Image Bank 
      Height          =   915
      Left            =   1080
      Picture         =   "Main.frx":96B7
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   55
      Left            =   6480
      Top             =   8880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   54
      Left            =   4320
      Picture         =   "Main.frx":C5AD
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   53
      Left            =   5400
      Picture         =   "Main.frx":F4A3
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Base 
      Height          =   915
      Left            =   9720
      Picture         =   "Main.frx":12399
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   52
      Left            =   7560
      Picture         =   "Main.frx":1528F
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   51
      Left            =   11880
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   50
      Left            =   11880
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   49
      Left            =   10800
      Picture         =   "Main.frx":18185
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   48
      Left            =   10800
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   47
      Left            =   11880
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   46
      Left            =   11880
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   45
      Left            =   9720
      Picture         =   "Main.frx":1B07B
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   44
      Left            =   7560
      Picture         =   "Main.frx":1DF71
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   43
      Left            =   6480
      Picture         =   "Main.frx":20E67
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   42
      Left            =   4320
      Picture         =   "Main.frx":23D5D
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   41
      Left            =   9720
      Picture         =   "Main.frx":26C53
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   40
      Left            =   7560
      Picture         =   "Main.frx":29B49
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   39
      Left            =   8640
      Picture         =   "Main.frx":2CA3F
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   38
      Left            =   7560
      Picture         =   "Main.frx":2F935
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   37
      Left            =   5400
      Picture         =   "Main.frx":3282B
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   36
      Left            =   10800
      Picture         =   "Main.frx":35721
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   35
      Left            =   9720
      Picture         =   "Main.frx":38617
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   34
      Left            =   8640
      Picture         =   "Main.frx":3B50D
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   33
      Left            =   9720
      Picture         =   "Main.frx":3E403
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   32
      Left            =   2160
      Picture         =   "Main.frx":412F9
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   31
      Left            =   3240
      Picture         =   "Main.frx":441EF
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   30
      Left            =   7560
      Picture         =   "Main.frx":470E5
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   29
      Left            =   7560
      Picture         =   "Main.frx":49FDB
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   28
      Left            =   6480
      Picture         =   "Main.frx":4CED1
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   27
      Left            =   5400
      Picture         =   "Main.frx":4FDC7
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   25
      Left            =   4320
      Picture         =   "Main.frx":52CBD
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   24
      Left            =   5400
      Picture         =   "Main.frx":55BB3
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   23
      Left            =   10800
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   22
      Left            =   5400
      Picture         =   "Main.frx":58AA9
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   21
      Left            =   2160
      Picture         =   "Main.frx":5B99F
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   20
      Left            =   6480
      Picture         =   "Main.frx":5E895
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   19
      Left            =   2160
      Picture         =   "Main.frx":6178B
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Tunnel 
      Height          =   915
      Index           =   1
      Left            =   11760
      Picture         =   "Main.frx":64681
      Top             =   3840
      Width           =   270
   End
   Begin VB.Image Tunnel 
      Height          =   915
      Index           =   0
      Left            =   0
      Picture         =   "Main.frx":6541B
      Top             =   3840
      Width           =   270
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   7
      Left            =   0
      Top             =   6720
      Width           =   135
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   18
      Left            =   4320
      Picture         =   "Main.frx":661B5
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   17
      Left            =   4320
      Picture         =   "Main.frx":690AB
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   15
      Left            =   8640
      Picture         =   "Main.frx":6BFA1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   14
      Left            =   3240
      Picture         =   "Main.frx":6EE97
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   13
      Left            =   2160
      Picture         =   "Main.frx":71D8D
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   12
      Left            =   2160
      Picture         =   "Main.frx":74C83
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   11
      Left            =   3240
      Picture         =   "Main.frx":77B79
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   10
      Left            =   2160
      Picture         =   "Main.frx":7AA6F
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   9
      Left            =   1080
      Picture         =   "Main.frx":7D965
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   8
      Left            =   0
      Picture         =   "Main.frx":8085B
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   6
      Left            =   2160
      Picture         =   "Main.frx":83751
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   5
      Left            =   0
      Picture         =   "Main.frx":86647
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   4
      Left            =   0
      Picture         =   "Main.frx":8953D
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   3
      Left            =   0
      Picture         =   "Main.frx":8C433
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   2
      Left            =   0
      Picture         =   "Main.frx":8F329
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   1
      Left            =   0
      Picture         =   "Main.frx":9221F
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   0
      Left            =   0
      Picture         =   "Main.frx":95115
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   26
      Left            =   9720
      Picture         =   "Main.frx":9800B
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image Player 
      Height          =   465
      Left            =   9120
      Picture         =   "Main.frx":9AF01
      Top             =   1200
      Width           =   870
   End
   Begin VB.Image Flowers 
      Height          =   915
      Index           =   0
      Left            =   4320
      Picture         =   "Main.frx":9B3FE
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Flowers 
      Height          =   915
      Index           =   4
      Left            =   8640
      Picture         =   "Main.frx":9E2F4
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Flowers 
      Height          =   915
      Index           =   1
      Left            =   6480
      Picture         =   "Main.frx":A11EA
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Flowers 
      Height          =   915
      Index           =   2
      Left            =   7560
      Picture         =   "Main.frx":A40E0
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Flowers 
      Height          =   915
      Index           =   3
      Left            =   4320
      Picture         =   "Main.frx":A6FD6
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Wall 
      Height          =   915
      Index           =   16
      Left            =   4320
      Picture         =   "Main.frx":A9ECC
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Medic 
      Height          =   255
      Left            =   1680
      Picture         =   "Main.frx":ACDC2
      Top             =   1200
      Width           =   285
   End
   Begin VB.Image Nitro 
      Height          =   255
      Left            =   2040
      Picture         =   "Main.frx":AD200
      Top             =   1560
      Width           =   315
   End
   Begin VB.Image Mine 
      Height          =   255
      Index           =   2
      Left            =   2880
      Picture         =   "Main.frx":AD682
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image Mine 
      Height          =   255
      Index           =   1
      Left            =   2400
      Picture         =   "Main.frx":ADAC0
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image Mine 
      Height          =   255
      Index           =   0
      Left            =   1680
      Picture         =   "Main.frx":ADEFE
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image GunBox 
      Height          =   255
      Left            =   2040
      Picture         =   "Main.frx":AE33C
      Top             =   1200
      Width           =   600
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstLeft As Boolean
Dim FirstTop As Boolean
Dim MakeSound As Long
Private Sub Back_Click()
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
Me.Hide
Mover.Enabled = False
Setup.Show
Setup.Timer1.Enabled = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
LastLeft = Player.Left
LastTop = Player.Top
If PauseGame = True Then
    PlayerSpeed = 0
ElseIf PauseGame = False Then
    PlayerSpeed = PlayerOldSpeed
End If
    Select Case KeyCode
    Case vbKeyP
    FirstLeft = True
    FirstTop = True
    If PauseGame = True Then
        PauseGame = False
        GamePause.Caption = ""
        Mover.Enabled = True
        Survival.Enabled = True
    ElseIf PauseGame = False Then
        PauseGame = True
        GamePause.Caption = "Game Paused"
        Mover.Enabled = False
        Survival.Enabled = False
    End If
    Case vbKeyUp
        If FirstTop = True Then
            Player.Top = Player.Top - Player.Width / 3
            FirstTop = False
        End If
        FirstLeft = True
        Player.Top = Player.Top - PlayerSpeed
        Player.Picture = PicturePallet.CorvettUp.Picture
        PlayerHeight = 870
        PlayerWidth = 465
    Case vbKeyDown
        FirstLeft = True
        FirstTop = True
        FirstPush = 1
        Player.Top = Player.Top + PlayerSpeed
        Player.Picture = PicturePallet.CorvettDown.Picture
        PlayerHeight = 870
        PlayerWidth = 465
    Case vbKeyRight
        FirstLeft = True
        FirstTop = True
        FirstPushLeft = 1
        Player.Left = Player.Left + PlayerSpeed
        Player.Picture = PicturePallet.CorvettRight.Picture
        PlayerHeight = 465
        PlayerWidth = 870
    Case vbKeyLeft
        FirstTop = True
        If FirstLeft = True Then
            Player.Left = Player.Left - Player.Width / 3
            FirstLeft = False
        End If
        Player.Left = Player.Left - PlayerSpeed
        Player.Picture = PicturePallet.CorvettLeft.Picture
        PlayerHeight = 465
        PlayerWidth = 870
    Case vbKeySpace
        If SpeedUp = False Then
            If Nitro1 >= 1 Then
                Nitro1 = Nitro1 - 1
                SpeedUp = True
            End If
        End If
    Case vbKeyD
       Mover.Enabled = False
       Survival.Enabled = False
       CheatMode (InputBox("Enter A Desired Cheat", "Cheat Mode"))
       TimeTimer.Enabled = True
       Mover.Enabled = True
       Survival.Enabled = True
End Select
For I = 0 To 54
    If Player.Top < Wall(I).Top + 800 And Player.Top > Wall(I).Top - PlayerHeight Then
        If Player.Left < Wall(I).Left + 800 And Player.Left > Wall(I).Left - PlayerWidth Then
            Player.Top = LastTop
            Player.Left = LastLeft
        End If
    End If
Next I
If Player.Top < GunBox.Top + 255 And Player.Top > GunBox.Top - PlayerHeight Then
    If Player.Left < GunBox.Left + 200 And Player.Left > GunBox.Left - PlayerWidth Then
        Gun = Gun + 1
        Items GunBox
    End If
End If
If Player.Top < Medic.Top + 255 And Player.Top > Medic.Top - PlayerHeight Then
    If Player.Left < Medic.Left + 200 And Player.Left > Medic.Left - PlayerWidth Then
        PlayerHealth = PlayerHealth + 10
        Items Medic
    End If
End If
For I = 0 To Mine.Count - 1
    If Player.Top < Mine(I).Top + 255 And Player.Top > Mine(I).Top - PlayerHeight Then
        If Player.Left < Mine(I).Left + 200 And Player.Left > Mine(I).Left - PlayerWidth Then
            PlayerHealth = PlayerHealth - 10
            Items Mine(I)
                MakeSound = sndPlaySound("Media\Sound\Explosion2.wav", &H1)
        End If
    End If
Next I
If Player.Top < Nitro.Top + 255 And Player.Top > Nitro.Top - PlayerHeight Then
    If Player.Left < Nitro.Left + 200 And Player.Left > Nitro.Left - PlayerWidth Then
        Nitro1 = Nitro1 + 1
        Items Nitro
    End If
End If
If Player.Top < Base.Top + 900 And Player.Top > Base.Top - PlayerHeight Then
    If Player.Left < Base.Left + 900 And Player.Left > Base.Left - PlayerWidth Then
        If GotObjective = True Then
            MakeSound = sndPlaySound("Media\Sound\CLAP.wav", SND_SYNC)
            GameOver.ScoreAdder.Enabled = True
            Dead "You Won"
        End If
    End If
End If
If Player.Top < Bank.Top + 255 And Player.Top > Bank.Top - PlayerHeight Then
    If Player.Left < Bank.Left + 200 And Player.Left > Bank.Left - PlayerWidth Then
        If GameType = "Robber" Then
            If GotObjective = False Then
                PlayerScore = PlayerScore + 600
                MakeSound = sndPlaySound("Media\Sound\CASH.wav", &H1)
                GotObjective = True
            End If
        End If
    End If
End If
If Player.Left <= -870 Then
    Player.Left = 11900
ElseIf Player.Left >= 11970 Then
    Player.Left = -800
End If

End Sub
Private Sub Form_Load()
Items GunBox
Items Medic
For I = 0 To Mine.Count - 1
    Items Mine(I)
Next I
Items Nitro
End Sub

Public Sub Helicopter_Click(Index As Integer)
If PauseGame = False Then
    If Gun >= 1 Then
        Gun = Gun - 1
        MakeSound = sndPlaySound("Media\Sound\Explosion.wav", &H1)
        For I = 1 To 300
            Randomize
            RedOn = Int(Rnd * 2 + 1)
            Select Case RedOn
                Case 1
                    Main.Circle (Helicopter(Index).Left + Helicopter(Index).Width / 2, Helicopter(Index).Top + Helicopter(Index).Height / 2), I, vbRed
                Case 2
                    Main.Circle (Helicopter(Index).Left + Helicopter(Index).Width / 2, Helicopter(Index).Top + Helicopter(Index).Height / 2), I, vbYellow
            End Select
        Next I
        If Unlim = 0 Then
            HelicopterLeft = HelicopterLeft + 1
            Helicopter(Index).Visible = False
            If GameType = "Survival" Then
                If HelicopterLeft = HelicopterAmount Then
                    Dead ("You Won")
                    GameOver.ScoreAdder.Enabled = True
                    MakeSound = sndPlaySound("Media\Sound\CLAP.wav", SND_SYNC)
                End If
            End If
        ElseIf Unlim = 1 Then
            Helicopter(Index).Left = HelicopterPad(Index).Left
            Helicopter(Index).Top = HelicopterPad(Index).Top
        End If
        PlayerScore = PlayerScore + ScoreAmount

    End If
End If
End Sub
Private Sub Mover_Timer()
For O = 0 To HelicopterAmount - 1
    If Helicopter(O).Visible = False Then
        Helicopter(O).Left = -1500
    End If
    If Helicopter(O).Visible = True Then
        I = Player.Left - Helicopter(O).Left
        If I < 0 Then
            LeftDirection = "Left"
        ElseIf I > 0 Then
            LeftDirection = "Right"
        ElseIf I = 0 Then
            LeftDirection = "Middle"
        End If
        '!!!'!!!'!!!'!!!'!!!'!!!'!!!'
        'Find Top Position:
        I = Player.Top - Helicopter(O).Top
        If I < 0 Then
            Helicopter(O).Picture = PicturePallet.HelicopterUp.Picture
            TopDirection = "Above"
        ElseIf I > 0 Then
            TopDirection = "Below"
            Helicopter(O).Picture = PicturePallet.HelicopterDown.Picture
        ElseIf I = 0 Then
            TopDirection = "Middle"
        End If
        If LeftDirection = "Right" Then
            Helicopter(O).Left = Helicopter(O).Left + HelicopterSpeed
        ElseIf LeftDirection = "Middle" Then
            Helicopter(O).Left = Helicopter(O).Left
        ElseIf LeftDirection = "Left" Then
            Helicopter(O).Left = Helicopter(O).Left - HelicopterSpeed
        End If
    
        If TopDirection = "Below" Then
            Helicopter(O).Top = Helicopter(O).Top + HelicopterSpeed
        ElseIf TopDirection = "Middle" Then
            Helicopter(O).Top = Helicopter(O).Top
        ElseIf TopDirection = "Above" Then
            Helicopter(O).Top = Helicopter(O).Top - HelicopterSpeed
        End If
    End If
Next O
Health.Caption = PlayerHealth
Nitros.Caption = Nitro1
Guns.Caption = Gun
Score.Caption = PlayerScore
If SpeedUp = True Then
    SpeedNum = SpeedNum + 1
    PlayerSpeed = PlayerOldSpeed * 4
    If SpeedNum >= 300 Then
        SpeedUp = False
        PlayerSpeed = PlayerOldSpeed
        SpeedNum = 1
    End If
End If
End Sub



Private Sub Quit_Click()
MakeSound = sndPlaySound("Media\Sound\button.wav", &H1)
End
End Sub
Private Sub Survival_Timer()
'''''''''''''''''SURVIVAL MODE'''''''''''''''''''''
If GameType = "Survival" Then
    ScoreAmount = 25
    PlayerScore = PlayerScore + 1
    For I = 0 To HelicopterAmount - 1
        If Player.Top < Helicopter(I).Top + 600 And Player.Top > Helicopter(I).Top - PlayerHeight Then
            If Player.Left < Helicopter(I).Left + 600 And Player.Left > Helicopter(I).Left - PlayerWidth Then
                MakeSound = sndPlaySound("Media\Sound\gun.wav", &H1)
                PlayerHealth = PlayerHealth - 1
                If PlayerHealth <= 0 Then
                    Dead ("Game Over")
                    GameOver.ScoreAdder.Enabled = True
                End If
            End If
        End If
    Next I
'''''''''''''''''ROBBER MODE'''''''''''''''''''''
ElseIf GameType = "Robber" Then
    ScoreAmount = 100
    For I = 0 To HelicopterAmount - 1
        If Player.Top < Helicopter(I).Top + 600 And Player.Top > Helicopter(I).Top - PlayerHeight Then
            If Player.Left < Helicopter(I).Left + 600 And Player.Left > Helicopter(I).Left - PlayerWidth Then
                PlayerHealth = PlayerHealth - 1
                MakeSound = sndPlaySound("Media\Sound\gun.wav", &H1)
                If PlayerHealth <= 0 Then
                    Dead ("Game Over")
                    GameOver.ScoreAdder.Enabled = True
                End If
            End If
        End If
    Next I
End If
EndGame:

End Sub
Public Sub Dead(TypeDead As String)
Me.Hide
Survival.Enabled = False
Mover.Enabled = False
GameOver.Show
GameOver.KillType.Caption = TypeDead
End Sub
Public Function Items(LeftItem As Image)
Randomize
ItemNum = Int(14 * Rnd + 1)
Select Case ItemNum
    Case 1
        LeftItem.Top = 6000
        LeftItem.Left = 10440
    Case 2
        LeftItem.Top = 3240
        LeftItem.Left = 8880
    Case 3
        LeftItem.Top = 3120
        LeftItem.Left = 3720
    Case 4
        LeftItem.Top = 1440
        LeftItem.Left = 3240
    Case 5
        LeftItem.Top = 5528
        LeftItem.Left = 1320
    Case 6
        LeftItem.Top = 2166
        LeftItem.Left = 5640
    Case 7
        LeftItem.Top = 3000
        LeftItem.Left = 5520
    Case 8
        LeftItem.Top = 1200
        LeftItem.Left = 3600
    Case 9
        LeftItem.Top = 4200
        LeftItem.Left = 360
    Case 10
        LeftItem.Top = 6120
        LeftItem.Left = 4440
    Case 11
        LeftItem.Top = 5040
        LeftItem.Left = 8760
    Case 12
        LeftItem.Top = 3240
        LeftItem.Left = 10920
    Case 13
        LeftItem.Top = 4200
        LeftItem.Left = 3480
    Case 14
        LeftItem.Top = 7080
        LeftItem.Left = 1190
    Case 15
        LeftItem.Top = 6840
        LeftItem.Left = 6720
    Case 16
        LeftItem.Top = 4440
        LeftItem.Left = 6840
End Select
End Function

Private Sub TimeTimer_Timer()
SecTime = SecTime + 1
If SecTime = 60 Then
    SecTime = 0
    MinTime = MinTime + 1
End If
FullTime = "Time  " & MinTime & ":" & SecTime
GameTime.Caption = FullTime
End Sub

Public Sub CheatMode(Cheat As String)
LCase (Cheat)
Select Case Cheat
    Case "lotsofhealth"
        'MsgBox "You now have 300 more health"
        PlayerHealth = PlayerHealth + 300
    Case "byebye"
        'MsgBox "All The Helicopters Are Gone"
        For I = 0 To 6
        Gun = Gun + 1
            Helicopter_Click (I)
        Next I
    Case "megofast"
        'MsgBox "You now have 50 more nitros"
        Nitro1 = Nitro1 + 50
    Case "givegun"
        'MsgBox "You now have 20 more guns"
        Gun = Gun + 10
    Case "giveall"
        'MsgBox "You have enabled all cheats"
        Gun = Gun + 10
        Nitro1 = Nitro1 + 50
        PlayerHealth = PlayerHealth + 300
    Case "scoreup"
        'MsgBox "You now have 500 more points"
        PlayerScore = PlayerScore + 500
    End Select
MakeSound = sndPlaySound("Media\Sound\Button.wav", &H1)
End Sub
