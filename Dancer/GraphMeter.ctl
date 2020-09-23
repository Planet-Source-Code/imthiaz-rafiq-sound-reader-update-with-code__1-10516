VERSION 5.00
Begin VB.UserControl GrapherX 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "GraphMeter.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1170
      Top             =   1410
   End
   Begin VB.Menu mnusimple 
      Caption         =   "&Simple"
      Visible         =   0   'False
      Begin VB.Menu mnustar32 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnus 
         Caption         =   "     Dancer And Grapher Ocx "
      End
      Begin VB.Menu mnustar6 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu er234 
         Caption         =   "    BrainFusion Tech Limited"
      End
      Begin VB.Menu fdfd 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudance 
         Caption         =   "                   Dance"
      End
      Begin VB.Menu mnustar1 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuprby 
         Caption         =   "          ICQ # 24294947"
      End
      Begin VB.Menu mnustar2 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuabout 
         Caption         =   "                 About"
      End
      Begin VB.Menu mnustar3 
         Caption         =   "####################"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "GrapherX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidthMax As Long, ByVal nHeightMax As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

Private Const M_Dancer_Dance_Enable = False
Private Const M_Dancer_Dance_Timer = 1000


Private nBarWidth As Long
Private nWidthMax As Long
Private nHeightMax As Long
Private nMax As Long
Private bInverted As Boolean

Dim Dancer_Dance_Enable As Boolean
Dim Dancer_Dance_Timer As Long

Public Enum Draw_Style
    Dots
    Lines
    Bar
End Enum

Dim hmixer As Long                  ' mixer handle
Dim inputVolCtrl As MIXERCONTROL    ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL   ' microphone volume control
Dim rc As Long                      ' return code
Dim ok As Boolean                   ' boolean return code

Dim VolumeMax As Long

Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' handle to volume memory

Dim temp As Long
Dim ColorF As Long
Private Bstyle As Draw_Style

Public Event OnTimer(VolumeValue As Long, VolumeMax As Long)

Public Property Get Drawstyle() As Draw_Style
    Drawstyle = Bstyle
End Property
Public Property Let Drawstyle(ByVal N_Style As Draw_Style)
    Bstyle = N_Style
    PropertyChanged "Drawstyle"
End Property

Public Property Get Max() As Long
    Max = nMax
End Property
Public Property Let Max(Val As Long)
    nMax = Val
    PropertyChanged "Max"
    UserControl_Resize
End Property

Public Property Get BarWidth() As Long
    BarWidth = nBarWidth
End Property
Public Property Let BarWidth(Val As Long)
    nBarWidth = Val
    PropertyChanged "BarWidth"
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(Val As OLE_COLOR)
    UserControl.BackColor = Val
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = ColorF
End Property
Public Property Let ForeColor(Val As OLE_COLOR)
    ColorF = Val
    PropertyChanged "ForeColor"
End Property

Public Property Get Flat() As Boolean
    Flat = (UserControl.BorderStyle = vbBSNone)
End Property
Public Property Let Flat(Val As Boolean)
    If Val Then
        UserControl.BorderStyle = vbBSNone
    Else
        UserControl.BorderStyle = vbFixedSingle
    End If
    PropertyChanged "Flat"
    UserControl_Resize
End Property

Public Property Get Inverted() As Boolean
    Inverted = bInverted
End Property
Public Property Let Inverted(Val As Boolean)
    bInverted = Val
    PropertyChanged "Inverted"
    UserControl_Resize
End Property

Private Sub mnuabout_Click()
    Form1.Show vbModal
End Sub

Private Sub mnudance_Click()
  Dancer_Dance_Enable = Not (Dancer_Dance_Enable)
  Timer1.Enabled = Dancer_Dance_Enable
  PropertyChanged "Dance"
End Sub

Private Sub Timer1_Timer()
      mxcd.dwControlID = outputVolCtrl.dwControlID
      mxcd.item = outputVolCtrl.cMultipleItems
      rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
        
      temp = (Abs(volume + 1) / VolumeMax) * Max
      Update temp
      RaiseEvent OnTimer(Abs(volume), Abs(VolumeMax))
End Sub

Private Sub UserControl_Initialize()
    ColorF = vbRed
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

'*********'*********'*********'*********'*********'*********'*********'*********
' Default property settings
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub UserControl_InitProperties()
    BackColor = vbButtonFace
    ForeColor = vbButtonText
    Bstyle = Bar
    Max = 100
    BarWidth = 2
    Flat = False
    Inverted = False
    Dancer_Dance_Timer = M_Dancer_Dance_Timer
    Dancer_Dance_Enable = M_Dancer_Dance_Enable
    ColorF = vbRed
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
'        PopupMenu mnusimple
    End If
End Sub

'*********'*********'*********'*********'*********'*********'*********'*********
' Read property settings
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'On Error Resume Next
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    ForeColor = PropBag.ReadProperty("ForeColor", ColorF)
    Max = PropBag.ReadProperty("Max", 100)
    BarWidth = PropBag.ReadProperty("BarWidth", 2)
    Flat = PropBag.ReadProperty("Flat", False)
    Inverted = PropBag.ReadProperty("Inverted", False)
    Bstyle = PropBag.ReadProperty("Bstyle", Bar)
    Dance = PropBag.ReadProperty("Dance", M_Dancer_Dance_Enable)
    Timer = PropBag.ReadProperty("Timer", M_Dancer_Dance_Timer)
End Sub

'*********'*********'*********'*********'*********'*********'*********'*********
' Write property settings
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'On Error Resume Next
    PropBag.WriteProperty "BackColor", BackColor
    PropBag.WriteProperty "ForeColor", ColorF
    PropBag.WriteProperty "Max", Max
    PropBag.WriteProperty "BarWidth", BarWidth
    PropBag.WriteProperty "Flat", Flat
    PropBag.WriteProperty "Inverted", Inverted
    PropBag.WriteProperty "Bstyle", Bstyle
    Call PropBag.WriteProperty("Dance", Dancer_Dance_Enable, M_Dancer_Dance_Enable)
    Call PropBag.WriteProperty("Timer", Dancer_Dance_Timer, M_Dancer_Dance_Timer)
End Sub

'*********'*********'*********'*********'*********'*********'*********'*********
' nWidthMax & nHeightMax are not height & width, they are
' they largest X & Y values on the control
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub UserControl_Resize()
    nWidthMax = UserControl.ScaleWidth
    nHeightMax = UserControl.ScaleHeight
    UserControl.Cls
End Sub

'*********'*********'*********'*********'*********'*********'*********'*********
' Draw the rightmost line & scroll left
'*********'*********'*********'*********'*********'*********'*********'*********
Public Sub Update(Val As Long)
    Dim temp As Long
    Dim temp1 As Long
    Dim x1 As Long
    Dim x2 As Long
    
    ' scroll left
    Call BitBlt(UserControl.hDC, 0 - nBarWidth, 0, nWidthMax + 1, nHeightMax + 1, UserControl.hDC, 0, 0, SRCCOPY)
    
    ' clear rightmost column
    Line (nWidthMax, nHeightMax)-(nWidthMax - nBarWidth, 0), UserControl.BackColor, BF
    
    ' draw line in rightmost column
    If Val > 0 Then
    
        temp = ((Val / nMax) * nHeightMax) - 1
        temp1 = ((Oldval / nMax) * nHeightMax) - 1
        
      If Bstyle = Dots Then
        If bInverted Then
            Line (nWidthMax - nBarWidth, temp)-(nWidthMax, temp - nBarWidth), ColorF, BF
        Else
            Line (nWidthMax - nBarWidth, nHeightMax - temp)-(nWidthMax, (nHeightMax - temp) + nBarWidth), ColorF, BF
        End If
      End If
      
      If Bstyle = Bar Then
        If bInverted Then
            Line (nWidthMax, 0)-(nWidthMax - nBarWidth, temp), ColorF, BF
        Else
            Line (nWidthMax, nHeightMax)-(nWidthMax - nBarWidth, nHeightMax - temp), ColorF, BF
        End If
      End If
      
      If Bstyle = Lines Then
        If bInverted Then
            If nBarWidth Mod 2 = 0 Then temp1 = 1 Else temp1 = 0
            If nBarWidth = 1 Then nBarWidth = 2
            x1 = nWidthMax - ((nBarWidth + temp1) / 2)
            x2 = nWidthMax - ((nBarWidth + temp1) / 2)
            Line (x1, 0)-(x2, temp), ColorF
        Else
            If nBarWidth Mod 2 = 0 Then temp1 = 1 Else temp1 = 0
            If nBarWidth = 1 Then nBarWidth = 2
            x1 = nWidthMax - ((nBarWidth + temp1) / 2)
            x2 = nWidthMax - ((nBarWidth + temp1) / 2)
            Line (x1, nHeightMax)-(x2, nHeightMax - temp), ColorF
        End If
      End If
    End If
End Sub
Public Property Get Dance() As Boolean
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
Public Sub DisplayAboutBox()
Attribute DisplayAboutBox.VB_UserMemId = -552
    frmTest.Show vbModal
End Sub

