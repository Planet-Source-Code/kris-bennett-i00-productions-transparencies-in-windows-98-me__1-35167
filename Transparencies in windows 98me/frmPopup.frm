VERSION 5.00
Begin VB.Form frmPopup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox Image1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   0
         Picture         =   "frmPopup.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C)Copyright 2000 i00

Dim fade As String



Private Sub Form_Load()

Me.Top = Screen.Height - 2055
Me.Left = Screen.Width - Me.Width

picDC.Width = ScaleWidth + 100
    picDC.Height = ScaleHeight + 100
    DeskHdc = GetDC(0)
    ret = BitBlt(picDC.hdc, 0, 0, ScaleWidth, ScaleHeight, DeskHdc, Left / Screen.TwipsPerPixelX, Top / Screen.TwipsPerPixelY, vbSrcCopy)
    ret = ReleaseDC(0&, DeskHdc)
'Picture = .picDC.Picture

fade = 0

Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    

        On Error Resume Next
        If fade > 255 Then Unload Me
        fade = fade + 15
        Me.Cls
        AlphaBlend Image1.hdc, 0, 0, picDC.ScaleWidth, picDC.ScaleHeight, picDC.hdc, 0, 0, picDC.ScaleWidth, picDC.ScaleHeight, CLng(fade) * (vbYellow + 1)
        Exit Sub


End Sub
