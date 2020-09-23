VERSION 5.00
Begin VB.UserControl DancerX 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ScaleHeight     =   3645
   ScaleWidth      =   1260
   ToolboxBitmap   =   "Dancer.ctx":0000
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   210
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30
      Top             =   1590
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   30
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   24
      Left            =   120
      Shape           =   1  'Square
      Top             =   315
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   30
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   23
      Left            =   120
      Shape           =   1  'Square
      Top             =   435
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   30
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   22
      Left            =   120
      Shape           =   1  'Square
      Top             =   555
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   30
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   21
      Left            =   120
      Shape           =   1  'Square
      Top             =   675
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   30
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   20
      Left            =   120
      Shape           =   1  'Square
      Top             =   795
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   19
      Left            =   120
      Shape           =   1  'Square
      Top             =   915
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   18
      Left            =   120
      Shape           =   1  'Square
      Top             =   1035
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   17
      Left            =   120
      Shape           =   1  'Square
      Top             =   1155
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   16
      Left            =   120
      Shape           =   1  'Square
      Top             =   1275
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   15
      Left            =   120
      Shape           =   1  'Square
      Top             =   1395
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderWidth     =   30
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   14
      Left            =   120
      Shape           =   1  'Square
      Top             =   1515
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderWidth     =   30
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   13
      Left            =   120
      Shape           =   1  'Square
      Top             =   1635
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderWidth     =   30
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   12
      Left            =   120
      Shape           =   1  'Square
      Top             =   1755
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderWidth     =   30
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   11
      Left            =   120
      Shape           =   1  'Square
      Top             =   1875
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderWidth     =   30
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   10
      Left            =   120
      Shape           =   1  'Square
      Top             =   1995
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      BorderWidth     =   30
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   9
      Left            =   120
      Shape           =   1  'Square
      Top             =   2115
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      BorderWidth     =   30
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   8
      Left            =   120
      Shape           =   1  'Square
      Top             =   2235
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      BorderWidth     =   30
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   7
      Left            =   120
      Shape           =   1  'Square
      Top             =   2355
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      BorderWidth     =   30
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   6
      Left            =   120
      Shape           =   1  'Square
      Top             =   2475
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      BorderWidth     =   30
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   5
      Left            =   120
      Shape           =   1  'Square
      Top             =   2595
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   30
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   4
      Left            =   120
      Shape           =   1  'Square
      Top             =   2715
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   30
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   3
      Left            =   120
      Shape           =   1  'Square
      Top             =   2835
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   30
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   2
      Left            =   120
      Shape           =   1  'Square
      Top             =   2955
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   30
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   1
      Left            =   120
      Shape           =   1  'Square
      Top             =   3075
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape Bar 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   30
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   0
      Left            =   120
      Shape           =   1  'Square
      Top             =   3195
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Menu mnuDancer 
      Caption         =   "&Simple"
      Enabled         =   0   'False
      Begin VB.Menu dsfs 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu sdfdf 
         Caption         =   "     Dancer And Grapher Ocx "
      End
      Begin VB.Menu qw 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep1 
         Caption         =   "    BrainFusion Tech Limited"
      End
      Begin VB.Menu erer44 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudance 
         Caption         =   "                   Dance"
      End
      Begin VB.Menu erer 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu icq 
         Caption         =   "          ICQ # 24294947"
      End
      Begin VB.Menu dfd445 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuabout 
         Caption         =   "                 About"
      End
      Begin VB.Menu fdgdf5 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "DancerX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const M_Dancer_Min = 1
Private Const M_Dancer_Max = 100
Private Const M_Dancer_Value = 100
Private Const M_Dancer_Mode = 1
Private Const M_Dancer_Orient = 1
Private Const M_Dancer_Size = 1
Private Const M_Dancer_Thick = 1
Private Const M_Dancer_Color = &H80000007
Private Const M_Dancer_Dance_Enable = False
Private Const M_Dancer_Dance_Timer = 1000

Dim Dancer_Min As Double
Dim Dancer_Max As Double
Dim Dancer_Value As Double
Dim Dancer_Mode As Integer
Dim Dancer_Orient As Integer
Dim Dancer_Size As Integer
Dim Dancer_LastValue As Double
Dim Dancer_Thick  As Integer
Dim Dancer_Color As OLE_COLOR
Dim Dancer_Dance_Enable As Boolean
Dim Dancer_Dance_Timer As Long

Dim hmixer As Long                  ' mixer handle
Dim inputVolCtrl As MIXERCONTROL    ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL   ' microphone volume control
Dim rc As Long                      ' return code
Dim ok As Boolean                   ' boolean return code

Dim temp As Double

Dim VolumeMax As Long

Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' handle to volume memory

Public Event MouseHover(Button As Integer, X As Single, Y As Single, Value As Double)
Public Event OnTimer(VolumeValue As Long, VolumeMax As Long)

Sub ThickChange(Thick As Integer)
Attribute ThickChange.VB_MemberFlags = "40"
Dim i As Integer
    For i = 0 To 24
        Bar(i).BorderWidth = Thick
    Next i
End Sub
Sub SizeChange(Size As Integer)
Attribute SizeChange.VB_MemberFlags = "40"
Dim i As Integer
    For i = 0 To 24
        Bar(i).Shape = Size
    Next i
End Sub
Sub Changemode(Mode As Integer, orr As Integer)
Attribute Changemode.VB_MemberFlags = "40"
Dim X As Integer
Dim Y As Integer
Dim Cy As Integer
Dim Z As Integer
Dim i As Integer
If Mode = 1 Then
    X = UserControl.ScaleWidth
    Y = UserControl.ScaleHeight
    If orr = 1 Then
        Cy = Int(Y / 49)
        For i = 0 To 24
            Bar(i).Left = Int(X / 10)
        Next i
        
        For i = 0 To 24
            Bar(i).Width = X - (2 * Int(X / 10))
        Next i
        Z = Cy
        'Stop
        For i = 24 To 0 Step -1
            Bar(i).Top = Z
            Z = Z + Cy + Cy * 0.9
        Next i
        
        For i = 0 To 24
            Bar(i).Height = Cy
        Next i
    End If
    If orr = 2 Then
        Cy = Int(Y / 49)
        For i = 0 To 24
            Bar(i).Left = Int(X / 10)
        Next i
        
        For i = 0 To 24
            Bar(i).Width = X - (2 * Int(X / 10))
        Next i
        Z = Cy
        'Stop
        For i = 0 To 24
            Bar(i).Top = Z
            Z = Z + Cy + Cy * 0.9
        Next i
        
        For i = 0 To 24
            Bar(i).Height = Cy
        Next i
    End If
End If
If Mode = 2 Then
    Y = UserControl.ScaleWidth
    X = UserControl.ScaleHeight
    If orr = 1 Then
        Cy = Int(Y / 49)
        
        For i = 0 To 24
            Bar(i).Top = Int(X / 10)
        Next i
        
        For i = 0 To 24
            Bar(i).Height = X - (2 * Int(X / 10))
        Next i
        
        Z = Cy
        
        'Stop
        For i = 24 To 0 Step -1
            Bar(i).Left = Z
            Z = Z + Cy + Cy * 0.9
        Next i
        
        For i = 0 To 24
            Bar(i).Width = Cy
        Next i
    End If
    If orr = 2 Then
        Cy = Int(Y / 49)
        For i = 0 To 24
            Bar(i).Top = Int(X / 10)
        Next i
        
        For i = 0 To 24
            Bar(i).Height = X - (2 * Int(X / 10))
        Next i
        
        Z = Cy
        
        'Stop
        For i = 0 To 24 Step 1
            Bar(i).Left = Z
            Z = Z + Cy + Cy * 0.9
        Next i
        
        For i = 0 To 24
            Bar(i).Width = Cy
        Next i
End If
End If
End Sub
Sub Dancerupdate(Max As Double, Min As Double, Value As Double)
Attribute Dancerupdate.VB_MemberFlags = "40"
Dim X As Double
Dim Lv As Double
Dim total As Double
X = Value
Lv = Dancer_LastValue
total = Max - 0
    per = Int((X / total) * 25)
    If per = 25 Then per = per - 1
    Bar(Dancer_LastValue).Visible = False
    Bar(per).Visible = True
    If per = 25 Then Stop
    If Lv > per Then
        For i = per + 1 To Lv
            Bar(i).Visible = False
        Next i
    ElseIf Lv < per Then
        For i = 0 To per - 1
            Bar(i).Visible = True
        Next i
    End If
    Dancer_LastValue = per
End Sub

Private Sub mnuabout_Click()
Form1.Show vbModal

End Sub

Private Sub mnudance_Click()
'  Dancer_Dance_Enable = Not (Dancer_Dance_Enable)
'  Timer1.Enabled = Dancer_Dance_Enable
'  PropertyChanged "Dance"
End Sub

Private Sub Timer1_Timer()
      mxcd.dwControlID = outputVolCtrl.dwControlID
      mxcd.item = outputVolCtrl.cMultipleItems
      rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
        
      temp = (Abs(volume) / VolumeMax) * Dancer_Max
      Dancerupdate Dancer_Max, Dancer_Min, temp
      RaiseEvent OnTimer(Abs(volume), Abs(VolumeMax))
End Sub

Private Sub Timer2_Timer()
  Timer2.Interval = 5000
  ThickChange Dancer_Thick
  SizeChange Dancer_Size - 1
  Changemode Dancer_Mode, Dancer_Orient
End Sub

Private Sub UserControl_DblClick()
  Form1.Show vbModal
End Sub

Private Sub UserControl_Initialize()
    Dancer_Min = M_Dancer_Min
    Dancer_Max = M_Dancer_Max
    Dancer_Value = M_Dancer_Value
    Dancer_Mode = M_Dancer_Mode
    Dancer_Orient = M_Dancer_Orient
    Dancer_Size = M_Dancer_Size
    Dancer_Thick = M_Dancer_Thick
    Dancer_LastValue = 24
    ThickChange Dancer_Thick
    SizeChange Dancer_Size
    Changemode Dancer_Mode, Dancer_Orient
    Dancerupdate Dancer_Max, Dancer_Min, Dancer_Value
    Dancer_Dance_Timer = M_Dancer_Dance_Timer
    Dancer_Dance_Enable = M_Dancer_Dance_Enable
    For i = 0 To 24
      Bar(i).Visible = True
    Next i
    ' Open the mixer specified by DEVICEID
    rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer."
        Exit Sub
    End If
        
    ' Get the output volume meter
    ok = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
    
    If (ok = True) Then
       VolumeMax = outputVolCtrl.lMaximum
    Else
       MsgBox "Couldn't get waveout meter"
       Timer1.Enabled = False
    End If
    
    ' Initialize mixercontrol structure
    mxcd.cbStruct = Len(mxcd)
    volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
    mxcd.paDetails = GlobalLock(volHmem)
    mxcd.cbDetails = Len(volume)
    mxcd.cChannels = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
'        PopupMenu mnuDancer
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseHover(Button, X, Y, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", M_Dancer_Value)
    Min = PropBag.ReadProperty("Min", M_Dancer_Min)
    Max = PropBag.ReadProperty("Max", M_Led_ColorOff)
    Mode = PropBag.ReadProperty("Mode", M_Dancer_Mode)
    Orientation = PropBag.ReadProperty("Orientation", M_Dancer_Orient)
    Size = PropBag.ReadProperty("Size", M_Dancer_Size)
    Thickness = PropBag.ReadProperty("Thickness", M_Dancer_Thick)
    Color = PropBag.ReadProperty("Color", M_Dancer_Color)
    Dance = PropBag.ReadProperty("Dance", M_Dancer_Dance_Enable)
    Timer = PropBag.ReadProperty("Timer", M_Dancer_Dance_Timer)
End Sub

Private Sub UserControl_Resize()
    Changemode Dancer_Mode, Dancer_Orient
End Sub
Public Property Get Value() As Double
Attribute Value.VB_Description = "Set The Value Dancer"
    Value = Dancer_Value
End Property

Public Property Let Value(ByVal New_Value As Double)
    If New_Value < Dancer_Max + 1 And New_Value > Dancer_Min - 1 Then
        Dancer_Value = New_Value
        Dancerupdate Dancer_Max, Dancer_Min, Dancer_Value
        PropertyChanged "Value"
    Else
        MsgBox "Value Out Of Range", vbOKOnly, "Dancer Ocx"
    End If
End Property

Public Property Get Mode() As Integer
Attribute Mode.VB_Description = "Set The Mode Horizontal Or Vertical"
    Mode = Dancer_Mode
End Property

Public Property Let Mode(ByVal New_Mode As Integer)
    If New_Mode = 1 Or New_Mode = 2 Then
        Dancer_Mode = New_Mode
        Changemode Dancer_Mode, Dancer_Orient
        PropertyChanged "Mode"
    Else
        MsgBox "Mode is Invalid Either 1 - Vertical or 2 - Horizontal ", vbOKOnly, "Dancer ocx"
    End If
End Property

Public Property Get Orientation() As Integer
Attribute Orientation.VB_Description = "Set Motion Forward Or Backward"
    Orientation = Dancer_Orient
End Property

Public Property Let Orientation(ByVal New_Orientation As Integer)
    If New_Orientation = 1 Or New_Orientation = 2 Then
        Dancer_Orient = New_Orientation
        Changemode Dancer_Mode, Dancer_Orient
    Else
        MsgBox "Orientation is Invalid Either 1 - Forward or 2 - Backwards ", vbOKOnly, "Dancer ocx"
    End If
End Property
Public Property Get Min() As Double
Attribute Min.VB_Description = "Set The Minimum Value"
    Min = Dancer_Min
End Property
Public Property Let Min(ByVal New_Min As Double)
    If New_Min < 0 Then
        MsgBox "No Negative Value Allowed", vbOKOnly, "Dancer Ocx"
        GoTo outy
    End If
    If New_Min > Dancer_Max Then
        MsgBox "Minimun Value is greater than Maximun Value", vbOKOnly, "Dancer Ocx"
        GoTo outy
    End If
    Dancer_Min = New_Min
    PropertyChanged "Min"
outy:
End Property
Public Property Get Max() As Double
Attribute Max.VB_Description = "Set The Maximum Value"
    Max = Dancer_Max
End Property
Public Property Let Max(ByVal New_Max As Double)
    If New_Max < 0 Then
        MsgBox "No Negative Value Allowed", vbOKOnly, "Dancer Ocx"
        GoTo outy
    End If
    If New_Max < Dancer_Min Then
        MsgBox "Minimun Value is less than Maximun Value", vbOKOnly, "Dancer Ocx"
        GoTo outy
    End If
    Dancer_Max = New_Max
    PropertyChanged "Max"
outy:
End Property
Public Property Get Size() As Integer
Attribute Size.VB_Description = "Set The Size Of The Dancer"
    Size = Dancer_Size
End Property
Public Property Let Size(ByVal New_Size As Integer)
    If New_Size > 0 And New_Size < 3 Then
        Dancer_Size = New_Size
        SizeChange Dancer_Size - 1
        PropertyChanged "Size"
    Else
       MsgBox "Invalid Size 1 - Big Size 2 - Small Size ", vbOKOnly, "Dancer Ocx"
    End If
End Property
Public Property Get Thickness() As Integer
Attribute Thickness.VB_Description = "Set The Thickness"
    Thickness = Dancer_Thick
End Property
Public Property Let Thickness(ByVal New_Thick As Integer)
    If New_Thick > 0 And New_Thick < 31 Then
        Dancer_Thick = New_Thick
        ThickChange Dancer_Thick
        PropertyChanged "Thickness"
    Else
        MsgBox "Invalid Thickness Range 0 to 30", vbOKCancel, "Dancer Ocx"
    End If
End Property
Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Sets The Background Color"
    Color = Dancer_Color
End Property
Public Property Let Color(ByVal New_Color As OLE_COLOR)
    Dancer_Color = New_Color
    UserControl.BackColor = Dancer_Color
    PropertyChanged "Color"
End Property
Public Property Get Dance() As Boolean
Attribute Dance.VB_Description = "Get Input From SoundCard by Interval You Specify [if enabled]"
    Dance = Dancer_Dance_Enable
End Property
Public Property Let Dance(ByVal New_Val As Boolean)
    Dancer_Dance_Enable = New_Val
    UserControl.Timer1.Enabled = Dancer_Dance_Enable
    PropertyChanged "Dance"
End Property

Public Property Get Timer() As Long
     Timer = Dancer_Dance_Timer
End Property
Public Property Let Timer(ByVal New_Val As Long)
    Dancer_Dance_Timer = New_Val
    UserControl.Timer1.Interval = Dancer_Dance_Timer
    PropertyChanged "Timer"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", Dancer_Value, M_Dancer_Value)
    Call PropBag.WriteProperty("Min", Dancer_Min, M_Dancer_Min)
    Call PropBag.WriteProperty("Max", Dancer_Max, M_Led_ColorOff)
    Call PropBag.WriteProperty("Mode", Dancer_Mode, M_Dancer_Mode)
    Call PropBag.WriteProperty("Orientation", Dancer_Orient, M_Dancer_Orient)
    Call PropBag.WriteProperty("Size", Dancer_Size, M_Dancer_Size)
    Call PropBag.WriteProperty("Thickness", Dancer_Thick, M_Dancer_Thick)
    Call PropBag.WriteProperty("Color", Dancer_Color, M_Dancer_Color)
    Call PropBag.WriteProperty("Dance", Dancer_Dance_Enable, M_Dancer_Dance_Enable)
    Call PropBag.WriteProperty("Timer", Dancer_Dance_Timer, M_Dancer_Dance_Timer)
End Sub
Public Sub DisplayAboutBox()
Attribute DisplayAboutBox.VB_UserMemId = -552
Attribute DisplayAboutBox.VB_MemberFlags = "40"
    Form1.Show vbModal
End Sub


