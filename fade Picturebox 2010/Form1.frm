VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   16695
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option6 
      Caption         =   "Option6"
      Height          =   345
      Left            =   15240
      TabIndex        =   8
      Top             =   5130
      Width           =   945
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Height          =   435
      Left            =   15270
      TabIndex        =   7
      Top             =   4440
      Width           =   915
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   495
      Left            =   15240
      TabIndex        =   6
      Top             =   3630
      Width           =   885
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   435
      Left            =   15240
      TabIndex        =   5
      Top             =   2910
      Width           =   1125
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   345
      Left            =   15180
      TabIndex        =   4
      Top             =   2070
      Width           =   1065
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   345
      Left            =   15210
      TabIndex        =   3
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   15210
      TabIndex        =   2
      Top             =   30
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   660
      ScaleHeight     =   1965
      ScaleWidth      =   12975
      TabIndex        =   1
      Top             =   6570
      Visible         =   0   'False
      Width           =   12975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8745
      Left            =   1140
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   8745
      ScaleWidth      =   11490
      TabIndex        =   0
      Top             =   300
      Width           =   11490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)





Private Sub Command1_Click()
AddPicture 127
Dim Systex
'Set picMovie = LoadPictureResource(PicNu, "Custom")


Picture1.Picture = LoadPicture("C:/PPS.Jpg")

fade Picture1, 4, 25
Picture1.Picture = LoadPicture("C:/Designer.Jpg")
    
End Sub

Sub AddPicture(PicNu As Integer)
'picMovie.ScaleMode = 3
Set Picture1 = LoadPictureResource(PicNu, "Custom")
'picMovie.Left = Width / 2 - (picMovie.Width / 2)
'picMovie.Top = 1000
'picMovie.Left = (Me.Width - picMovie.Width) / 2

'picMovie.Left = 50
   ' picMovie.AutoRedraw = True
End Sub

