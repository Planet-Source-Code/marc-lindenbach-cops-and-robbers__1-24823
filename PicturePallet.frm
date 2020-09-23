VERSION 5.00
Begin VB.Form PicturePallet 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Wound 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3000
      Picture         =   "PicturePallet.frx":0000
      Top             =   600
      Width           =   285
   End
   Begin VB.Image HelicopterDown 
      Height          =   915
      Left            =   1440
      Picture         =   "PicturePallet.frx":038B
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image HelicopterUp 
      Height          =   915
      Left            =   1440
      Picture         =   "PicturePallet.frx":08A0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image CorvettLeft 
      Height          =   465
      Left            =   360
      Picture         =   "PicturePallet.frx":0DB0
      Top             =   1080
      Width           =   870
   End
   Begin VB.Image CorvettRight 
      Height          =   465
      Left            =   360
      Picture         =   "PicturePallet.frx":12AD
      Top             =   1680
      Width           =   870
   End
   Begin VB.Image CorvettUp 
      Height          =   870
      Left            =   840
      Picture         =   "PicturePallet.frx":17AA
      Top             =   120
      Width           =   465
   End
   Begin VB.Image CorvettDown 
      Height          =   870
      Left            =   240
      Picture         =   "PicturePallet.frx":1C9F
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "PicturePallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
