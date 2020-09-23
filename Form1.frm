VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Working..."
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   327
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2880
      Top             =   2760
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   3240
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picBall 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   780
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1800
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   60
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   60
      Width           =   2595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xP As Long
Dim iDir As Long
Private Const AC_SRC_OVER As Long = &H0&
Private Const ULW_COLORKEY As Long = &H1&
Private Const ULW_ALPHA As Long = &H2&
Private Const ULW_OPAQUE As Long = &H4&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TRANSPARENT  As Long = &H20&
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const SPI_GETSELECTIONFADE As Long = &H1014&

Private Sub Form_Load()
    picBlank.Move picDisplay.Left, picDisplay.Top, picDisplay.Width, picDisplay.Height
    picBuffer.Move picDisplay.Left, picDisplay.Top, picDisplay.Width, picDisplay.Height

    xP = 0
    iDir = 4
    Timer1.Enabled = True

End Sub
Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub

Sub BufferToScreen()
    BitBlt picDisplay.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBuffer.hdc, 0, 0, vbSrcCopy
End Sub
Sub ApplyBlend()
Dim Blend As BLENDFUNCTION
Dim BlendPtr As Long
    Blend.SourceConstantAlpha = 16
    
    CopyMemory BlendPtr, Blend, 4
    
    AlphaBlend picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, picBuffer.Width, picBuffer.Height, BlendPtr
End Sub

Private Sub Timer1_Timer()
    'ClearBuffer
    ApplyBlend
    'draw ball
    If iDir > 0 And (xP + picBall.Width + iDir) > picBuffer.Width Then
        iDir = iDir * -1
    ElseIf iDir < 0 And xP < 2 Then
        iDir = iDir * -1
    End If
    'BitBlt picBuffer.hdc, xP, 2, picBall.Width, picBall.Height, picBall.hdc, 0, 0, vbSrcCopy
    ApplyBall xP
    xP = xP + iDir
    
    BufferToScreen
    
End Sub
Sub ApplyBall(x As Long)
Dim BF As BLENDFUNCTION
Dim lBF As Long
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = WS_EX_TRANSPARENT
        .SourceConstantAlpha = 255
        .AlphaFormat = 1
    End With
    RtlMoveMemory lBF, BF, 4
    GdiAlphaBlend picBuffer.hdc, x, 1, picBall.Width, picBall.Height, picBall.hdc, 0, 0, picBall.Width, picBall.Height, lBF
End Sub

