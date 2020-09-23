VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dancer & Grapher Ocx By Brainfusion"
   ClientHeight    =   2895
   ClientLeft      =   6285
   ClientTop       =   5490
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   2550
      ScaleHeight     =   1875
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   870
      Width           =   4515
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF80&
         Height          =   5055
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmabout.frx":0000
         Top             =   1650
         Width           =   4395
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1185
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   750
         Top             =   1140
      End
   End
   Begin Dancer_Ocx.GrapherX Grapher1 
      Height          =   1845
      Left            =   690
      Top             =   390
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   3254
      BackColor       =   0
      ForeColor       =   12648384
      Max             =   100
      BarWidth        =   1
      Flat            =   -1  'True
      Inverted        =   0   'False
      Bstyle          =   2
      Timer           =   1
   End
   Begin Dancer_Ocx.DancerX Dancer1 
      Height          =   1875
      Left            =   360
      TabIndex        =   0
      Top             =   390
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   3307
      Max             =   100
      Color           =   0
      Timer           =   50
   End
   Begin Dancer_Ocx.DancerX Dancer2 
      Height          =   1875
      Left            =   1920
      TabIndex        =   1
      Top             =   390
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   3307
      Max             =   100
      Color           =   0
      Timer           =   50
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BrainFusion "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   2520
      TabIndex        =   4
      Top             =   300
      Width           =   3585
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   90
      Top             =   120
      Width           =   2325
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   180
      Top             =   210
      Width           =   2145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Min"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Max"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1380
      TabIndex        =   2
      Top             =   2280
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   270
      Top             =   300
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dancer1_OnTimer(VolumeValue As Long, VolumeMax As Long)
If VolumeValue = VolumeMax Then Label1.ForeColor = &HFF& Else Label1.ForeColor = &H0&
If VolumeValue = 0 Then Label2.ForeColor = &HFF& Else Label2.ForeColor = &H0&
End Sub

Private Sub Dancer2_OnTimer(VolumeValue As Long, VolumeMax As Long)
Shape1.BorderColor = RGB((VolumeValue / VolumeMax) * 255, 0, 0)
Shape2.BorderColor = RGB(0, 0, 255 - (VolumeValue / VolumeMax) * 255)
Shape3.BorderColor = RGB(0, (VolumeValue / VolumeMax) * 255, 0)
End Sub

Private Sub Form_Load()
Dancer1.Dance = True
Dancer2.Dance = True
Grapher1.Dance = True
VScroll1.Max = Picture1.Height
VScroll1.Min = 0 - Text1.Height
VScroll1.Value = VScroll1.Max

End Sub


Private Sub Grapher1_OnTimer(VolumeValue As Long, VolumeMax As Long)
Grapher1.ForeColor = RGB(100 + (VolumeValue / VolumeMax) * 155, 0, 0)
End Sub

Private Sub Text1_Click()
Timer1.Enabled = False


End Sub

Private Sub Text1_DblClick()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
'Move the textbox here
If VScroll1.Value >= VScroll1.Min + 30 Then
  VScroll1.Value = VScroll1.Value - 20
Else
  VScroll1.Value = VScroll1.Max
  DoEvents
End If
If Text1.Visible = False Then
    Text1.Visible = True
End If
Text1.Top = VScroll1.Value
DoEvents
End Sub
