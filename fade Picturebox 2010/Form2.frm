VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "unload me"
      Height          =   495
      Left            =   13650
      TabIndex        =   3
      Top             =   9870
      Width           =   1605
   End
   Begin VB.PictureBox picMovie 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7725
      Left            =   12120
      ScaleHeight     =   7725
      ScaleWidth      =   9405
      TabIndex        =   2
      Top             =   240
      Width           =   9405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replace Picture"
      Height          =   555
      Left            =   11130
      TabIndex        =   1
      Top             =   9180
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7785
      Left            =   240
      ScaleHeight     =   7785
      ScaleWidth      =   11745
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   11745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F2,F3,F7,F8,F9,F11,F12 Also replaces the picture"
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   9900
      Width           =   7275
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Dim PicNumber As Integer
Private Const AC_SRC_OVER = &H0
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Declare Function AlphaBlend Lib "msimg32.dll" _
  (ByVal hdc As Long, ByVal lInt As Long, _
  ByVal lInt As Long, ByVal lInt As Long, _
  ByVal lInt As Long, ByVal hdc As Long, _
  ByVal lInt As Long, ByVal lInt As Long, _
  ByVal lInt As Long, ByVal lInt As Long, _
  ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
  (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" _
  (ByVal dwMilliseconds As Long)
Private BF As BLENDFUNCTION, lBF As Long, Fade As Byte
Private FadeInProgress As Boolean


Private Sub Command1_Click()
FadePic
End Sub


Private Sub Command2_Click()
'Picture1.Top = 0
'Picture1.Left = 0
'picMovie.Top = 0
'picMovie.Left = 0
Me.WindowState = vbNormal
Dim hRgn As Long, X As Long, Y As Long, Cx As Long, Cy As Long
X = ScaleX(picMovie.Left, Me.ScaleMode, vbPixels)
Y = ScaleY(picMovie.Top, Me.ScaleMode, vbPixels)
Cx = ScaleX(picMovie.Left + Picture1.Width, Me.ScaleMode, vbPixels)
Cy = ScaleY(picMovie.Top + Picture1.Height, Me.ScaleMode, vbPixels)

hRgn = CreateRectRgn(X, Y, Cx, Cy)
SetWindowRgn Me.hWnd, hRgn, True

' Me.ScaleMode = vbPixels
'
' Me.Width = Picture1.ScaleWidth '* 25
' Me.Height = Picture1.ScaleHeight '* 25

'Set picMovie = LoadPictureResource(Val(Text1), "Custom") 'main pic
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
     Select Case KeyCode
        Case 113 'F2
        Set Picture1 = LoadPictureResource(132, "Custom")
        FadePic
        
        Case 114 'F3
        Set Picture1 = LoadPictureResource(133, "Custom")
        FadePic
        
       Case 115 'F4
        Set Picture1 = LoadPictureResource(134, "Custom")
        FadePic
        Me.picMovie.Visible = False
     
     Case 118 'F7
     Set Picture1 = LoadPictureResource(127, "Custom")
     FadePic
     
     Case 119 'F8
     Set Picture1 = LoadPictureResource(128, "Custom")
     FadePic
     
     'skip F10
     Case 120 'F9
     Set Picture1 = LoadPictureResource(129, "Custom")
     FadePic
     
     
     Case 122 'F11
     Set Picture1 = LoadPictureResource(130, "Custom")
     FadePic
     
     Case 123 'F12
     Set Picture1 = LoadPictureResource(131, "Custom")
     FadePic
     
     End Select
    End If
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Me.Show
    Picture1.AutoRedraw = True
    Picture1.ScaleMode = vbPixels
    picMovie.AutoRedraw = True
    picMovie.ScaleMode = vbPixels
  PicNumber = 127
    AddPicture PicNumber '1st replacement picture
    
    Set picMovie = LoadPictureResource(PicNumber, "Custom") 'main pic
    
    picMovie.Width = Picture1.Width
    picMovie.Height = Picture1.Height
   Command1.Top = picMovie.Top + picMovie.Height + 200
   picMovie.Left = Me.ScaleWidth / 2 - picMovie.Width / 2
   Me.ScaleMode = vbPixels
   
   
End Sub

Sub FadePic()
    If FadeInProgress Then Exit Sub
    PicNumber = PicNumber + 1
    
  For Fade = 1 To 60 Step 2
      With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Fade
        .AlphaFormat = 0
      End With
      
      RtlMoveMemory lBF, BF, 4
      AlphaBlend picMovie.hdc, 0, 0, picMovie.ScaleWidth, _
      picMovie.ScaleHeight, Picture1.hdc, 0, 0, _
      Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
      picMovie.Refresh
      Sleep 25
  Next Fade
    DoEvents
    
    Set Picture1 = LoadPictureResource(PicNumber, "Custom") 'load a new pic
End Sub

Sub AddPicture(PicNu As Integer)
Set Picture1 = LoadPictureResource(PicNu, "Custom")
End Sub

Sub SetTrans(hWnd As Long, Trans As Integer)
    Dim Tcall As Long
    
    If Trans <= 0 Then
        Exit Sub
    Else
        Tcall = GetWindowLong(Picture1.hWnd, GWL_EXSTYLE)
        SetWindowLong hWnd, GWL_EXSTYLE, Tcall Or WS_EX_LAYERED
        SetLayeredWindowAttributes hWnd, RGB(255, 255, 0), Trans, LWA_ALPHA
    End If

End Sub

