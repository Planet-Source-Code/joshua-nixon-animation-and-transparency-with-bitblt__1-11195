VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   360
      Picture         =   "bitblt.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   4800
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   240
      Picture         =   "bitblt.frx":F042
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ANIMATE!"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1020
      Left            =   480
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Counter 'setup a counter for the x value
'Declare BitBlt and Constants
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCAND = &H8800C6
Const SRCPAINT = &HEE0086
Private Sub Command1_Click()
If Not Timer1.Enabled Then  'Check to see if timer is enabled or not
    Timer1.Enabled = True   'turn the timer on
    Command1.Caption = "STOP"
Else
    Timer1.Enabled = False  'turn the timer off
    Command1.Caption = "ANIMATE!"
    Picture1.Cls 'clear picture1
End If
End Sub

Private Sub Form_Load()
Counter = 0 'set the x value to 0
End Sub

Private Sub Timer1_Timer()

Dim ani As Long

'Paint the mask with SRCPAINT onto picture1
BitBlt Picture1.hDC, 0, 0, 64, 64, Picture3.hDC, Counter, 0, SRCPAINT

'Copy the picture with SRCAND onto picture1
BitBlt Picture1.hDC, 0, 0, 64, 64, Picture2.hDC, Counter, 0, SRCAND

Picture1.Refresh 'refresh picture1
Counter = Counter + 64 'add 64 to x
If Counter >= 320 Then Counter = 0 'checks to see if all the frames have been displayed. if they have then restart from the first frame
End Sub



