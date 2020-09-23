VERSION 5.00
Begin VB.UserControl dmFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ControlContainer=   -1  'True
   MouseIcon       =   "dmFrame.ctx":0000
   ScaleHeight     =   91
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
End
Attribute VB_Name = "dmFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Original code by:-
' dreamvb PSC CodeId=58966

' FEATURES
' Change bar color
' Choice of using normal color for the bar or a Gradient color
' Change font properties of the bar caption
' Change the outline color of the frame
' Change the style of the frame
' Change the frames background color
' Added events to the frame and the bar for the frame
' Also support for moving the frames around on your forms
' O ok I just now added alignment for the caption also for you all

' RR added Moveable or not
'          Outline DrawWidth
'          Enable/Disable

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private m_Caption As String ' Stote the Caption for the frame
Private Const m_def_Caption As String = "CoolXPFrame" ' Default Caption for above
Private n_BarHeight As Integer   'Height of the bar along the top of the frame
Private m_GradEn As Boolean     ' used to turn on or off Gradient support for the bar
Private m_BarColor As OLE_COLOR ' Color of the bar
Private m_OutLineColor As OLE_COLOR ' Bordercolor that goes around the frame
Private m_OutLineStyle As DrawStyleConstants      ' Drawstyle for the outline
Private m_Alignment As dmAlignment  ' Caption Aligment

'RR
Private m_Moveable As Boolean
Private m_DrawWidth As Long   ' Draw width of outline
Private m_Enabled As Boolean

Private OldY As Single
Private OnBar As Boolean

Private Type dmRgb
    Red As Long
    Green As Long
    Blue As Long
End Type

Enum dmAlignment ' Alignment Enum
    dmLeft = 0
    dmCenter
    dmRight
End Enum

Dim dm_Rgb As dmRgb
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event BarMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Click()
Event BarClick()
Event DblClick()
Event BarDblClick()

Sub DrawdmFrame()
    Dim i As Long, culfac As Long
    ' This sub is in fact the main part of the project that draws the frame
    ' the rest of the project is just added features
    
    UserControl.Cls ' Clear the the usercontrol
    n_BarHeight = UserControl.TextHeight("Xy") + 1 ' Get the bars height from the fontsize
    'LongToRGB m_BarColor
    ' RR faster than CopyMemory method
    dm_Rgb.Red = (m_BarColor And &HFF&)
    dm_Rgb.Green = (m_BarColor And &HFF00&) / &H100&
    dm_Rgb.Blue = (m_BarColor And &HFF0000) / &H10000

    If m_GradEn Then
        For i = 0 To n_BarHeight  ' this bit of code just draws a Gradient bar very simple
            ' RR
            culfac = ((n_BarHeight - i) * 4)
            ' NB: the value for any argument to RGB that exceeds 255 is assumed to be 255
            UserControl.Line (0, i)-(UserControl.ScaleWidth, i), _
            RGB(culfac + dm_Rgb.Red, culfac + dm_Rgb.Green, culfac + dm_Rgb.Blue)
        Next i
    Else
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, n_BarHeight), RGB(dm_Rgb.Red, dm_Rgb.Green, dm_Rgb.Blue), BF
    End If
    
    UserControl.CurrentY = 1.3  ' Position caption Y
    
    'Alignment options for the caption
    Select Case m_Alignment
        Case dmLeft ' Left
            UserControl.CurrentX = 3
        Case dmCenter ' Center
            UserControl.CurrentX = (UserControl.ScaleWidth - TextWidth(m_Caption) + 3) / 2
        Case dmRight ' Right
            UserControl.CurrentX = (UserControl.ScaleWidth - TextWidth(m_Caption) - 3)
    End Select
    
    UserControl.Print m_Caption ' Print on the caption
        
    UserControl.DrawStyle = m_OutLineStyle ' Just for some effects
    UserControl.DrawWidth = m_DrawWidth
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), m_OutLineColor, B
    ' the line above just draws the outline of the frame
    UserControl.DrawStyle = vbSolid ' Restote the user control drawstyle back to Soild
    UserControl.DrawWidth = 1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)

    If Button <> vbLeftButton Then Exit Sub ' see if we using the left button. chnage if you like
    If m_Moveable Then  ' RR
      If Not (y > (n_BarHeight - 1)) Then ' check if we are not over the bars height
          ' Small usfull bit of code for moveing a window around
          ' TIP you chould use something like this to make your own titlebar
          Call ReleaseCapture
          Call SendMessage(UserControl.hwnd, &HA1, 2, 0&)
      End If
    End If
End Sub


Private Sub UserControl_Initialize()
    m_Caption = m_def_Caption
    m_Alignment = dmLeft
    m_OutLineStyle = vbSolid
    m_OutLineColor = &H80000010
    m_BarColor = vbBlue
    m_GradEn = True

    ' RR
    m_Moveable = False
    m_DrawWidth = 1
    m_Enabled = True

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_OutLineColor = PropBag.ReadProperty("OutLineColor", &H80000010)
    m_Alignment = PropBag.ReadProperty("Alignment", 0)
    m_BarColor = PropBag.ReadProperty("BarColor", vbBlue)
    m_GradEn = PropBag.ReadProperty("UseGradient", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_OutLineStyle = PropBag.ReadProperty("OutLineStyle", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    ' RR
    m_Moveable = PropBag.ReadProperty("RRMoveable", False)
    m_DrawWidth = PropBag.ReadProperty("RRDrawWidth", 1)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("OutLineColor", m_OutLineColor, &H80000010)
    Call PropBag.WriteProperty("BarColor", m_BarColor, vbBlue)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("UseGradient", m_GradEn, True)
    Call PropBag.WriteProperty("OutLineStyle", m_OutLineStyle, 0)
    Call PropBag.WriteProperty("Alignment", m_Alignment, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    ' RR
    Call PropBag.WriteProperty("RRMoveable", m_Moveable, False)
    Call PropBag.WriteProperty("RRDrawWidth", m_DrawWidth, 1)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub


' Moveable RR
Public Property Get RRMoveable() As Boolean
   RRMoveable = m_Moveable
End Property
Public Property Let RRMoveable(ByVal Visible As Boolean)
   m_Moveable = Visible
   Call UserControl.PropertyChanged("RRMoveable")
   Call UserControl.Refresh
   DrawdmFrame
End Property

' DrawWidth RR
Public Property Get RRDrawWidth() As Long
   RRDrawWidth = m_DrawWidth
End Property
Public Property Let RRDrawWidth(ByVal DW As Long)
   m_DrawWidth = DW
   Call UserControl.PropertyChanged("RRDrawWidth")
   Call UserControl.Refresh
   DrawdmFrame
End Property

' Enabled RR
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
  m_Enabled = vNewValue
  UserControl.Enabled = m_Enabled
  DrawdmFrame
  PropertyChanged "Enabled"
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Property Get Alignment() As dmAlignment
Attribute Alignment.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As dmAlignment)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    DrawdmFrame
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawdmFrame
End Property


Private Sub UserControl_Resize()
    DrawdmFrame
End Sub

Private Sub UserControl_Show()
    DrawdmFrame
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawdmFrame
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    DrawdmFrame
End Property

Public Property Get OutLineColor() As OLE_COLOR
    OutLineColor = m_OutLineColor
End Property
Public Property Let OutLineColor(ByVal New_OutlineColor As OLE_COLOR)
    m_OutLineColor = New_OutlineColor
    PropertyChanged "OutLineColor"
    DrawdmFrame
End Property


Public Property Get BarColor() As OLE_COLOR
    BarColor = m_BarColor
End Property
Public Property Let BarColor(ByVal New_BarColor As OLE_COLOR)
    m_BarColor = New_BarColor
    
    ' RR allow use of system colors
    If m_BarColor And &H80000000 Then
      m_BarColor = GetSysColor(m_BarColor And &HFFFFFF)
    End If
    
    PropertyChanged "BarColor"
    DrawdmFrame
End Property

Public Property Get UseGradient() As Boolean
    UseGradient = m_GradEn
End Property
Public Property Let UseGradient(ByVal vNewValue As Boolean)
    m_GradEn = vNewValue
    DrawdmFrame
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawdmFrame
End Property

Public Property Get OutLineStyle() As DrawStyleConstants
Attribute OutLineStyle.VB_Description = "Determines the line style for output from graphics methods."
    OutLineStyle = m_OutLineStyle
End Property
Public Property Let OutLineStyle(ByVal New_OutLineStyle As DrawStyleConstants)
    m_OutLineStyle = New_OutLineStyle
    PropertyChanged "OutLineStyle"
    DrawdmFrame
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    OnBar = False
    
    If Not (y > (n_BarHeight - 1)) Then
        OldY = y: OnBar = True
        RaiseEvent BarMouseDown(Button, Shift, x, y)
    End If
     
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Click()

    If Not OnBar Then RaiseEvent Click
    
    If Not (0 > (n_BarHeight - 1)) And OnBar Then
        RaiseEvent BarClick
    End If
    
End Sub

Private Sub UserControl_DblClick()
    If Not OnBar Then RaiseEvent DblClick
    
    If Not (0 > (n_BarHeight - 1)) And OnBar Then
        RaiseEvent BarDblClick
    End If
    
End Sub

