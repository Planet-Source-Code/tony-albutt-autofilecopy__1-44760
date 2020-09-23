VERSION 5.00
Begin VB.UserControl cpvProgressBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2700
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Picture         =   "cpvProgressBar.ctx":0000
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   ToolboxBitmap   =   "cpvProgressBar.ctx":07A7
End
Attribute VB_Name = "cpvProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'-------------------------------------------------------------------------------------------
' cpvProgressBar OCX 1.1
'
' Carles P.V.
' carles_pv@terra.es
'-------------------------------------------------------------------------------------------

' 08/04/01: Added:     - Custom caption
'                      - Precise value selection (by Right mouse button) +/- 1
'                      - MouseIcon/MousePointer properties

' 08/10/01: Modified:  - Long values accepted*
'                        *: [- 2.147.483.648 , 2.147.483.647]
'                           AbsCount can't exceed max value: 2.147.483.647!
'           Fixed:     - Selection of out of range values (See MouseMove/Down events)





Option Explicit

Public Enum pbCaptionFormats
    [fNothing]     'No caption
    [fValue]       'Value
    [f0%]          'Relative % ±0
    [f0.0%]        'Relative % ±0.0
    [f0.00%]       'Relative % ±0.00
    [fAbs0%]       'Absolute % +0
    [fAbs0.0%]     'Absolute % +0.0
    [fAbs0.00%]    'Absolute % +0.00
    [fCustom] = 99 'Custom format
End Enum

Public Enum pbOrientationConstants
    [Horizontal]
    [Vertical]
End Enum

Private MouseSelect As Boolean 'Mouse scrolling flag

Private tmpX      As Single    'Temp params. of mouse click (See DblClick event)
Private tmpY      As Single
Private tmpShift  As Integer
Private tmpButton As Integer

Private AbsCount  As Long      'AbsCount = Max - Min
Private LastValue As Long      'LastValue (def. = Min)

'-- Default Property Values:
Private Const m_def_CaptionCustom = ""
Private Const m_def_CaptionFormat = 0
Private Const m_def_Max = 100
Private Const m_def_Min = 0
Private Const m_def_Orientation = 0
Private Const m_def_Value = 0

'-- Property Variables:
Private m_CaptionCustom  As String
Private m_CaptionFormat  As pbCaptionFormats
Private m_BarPictureBack As StdPicture
Private m_BarPicture     As StdPicture
Private m_Max            As Long
Private m_Min            As Long
Private m_Orientation    As pbOrientationConstants
Private m_Value          As Long

'-- Event Declarations:
Public Event Click()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event ValueChanged()

'-------------------------------------------------------------------------------------------
' InitProperties/ReadProperties/WriteProperties
'-------------------------------------------------------------------------------------------

Private Sub UserControl_InitProperties()

    m_CaptionCustom = m_def_CaptionCustom
    m_CaptionFormat = m_def_CaptionFormat
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Orientation = m_def_Orientation
    m_Value = m_def_Value

    Set UserControl.Font = Ambient.Font
    Set m_BarPicture = LoadPicture("")
    Set m_BarPictureBack = LoadPicture("")

    AbsCount = 100
    LastValue = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
    
        m_CaptionCustom = .ReadProperty("CaptionCustom", m_def_CaptionCustom)
        m_CaptionFormat = .ReadProperty("CaptionFormat", m_def_CaptionFormat)
        m_Max = .ReadProperty("Max", m_def_Max)
        m_Min = .ReadProperty("Min", m_def_Min)
        m_Orientation = .ReadProperty("Orientation", m_def_Orientation)
        m_Value = .ReadProperty("Value", m_def_Value)
    
        UserControl.Enabled = .ReadProperty("Enabled", True)
        UserControl.ForeColor = .ReadProperty("FontColor", &H80000012)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        Set m_BarPicture = .ReadProperty("BarPicture", Nothing)
        Set m_BarPictureBack = .ReadProperty("BarPictureBack", Nothing)
    
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
    End With
    
    AbsCount = m_Max - m_Min
    LastValue = m_Min
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BarPicture", m_BarPicture, Nothing)
        Call .WriteProperty("BarPictureBack", m_BarPictureBack, Nothing)
        Call .WriteProperty("CaptionCustom", m_CaptionCustom, m_def_CaptionCustom)
        Call .WriteProperty("CaptionFormat", m_CaptionFormat, m_def_CaptionFormat)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("FontColor", UserControl.ForeColor, &H80000012)
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Min", m_Min, m_def_Min)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("Orientation", m_Orientation, m_def_Orientation)
        Call .WriteProperty("Value", m_Value, m_def_Value)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
    End With
End Sub

'-------------------------------------------------------------------------------------------
' Resizing/Refreshing
'-------------------------------------------------------------------------------------------

Private Sub UserControl_Resize()

    On Error Resume Next

    If (m_BarPictureBack.Type = vbPicTypeNone) Then
        Size ScaleX(Picture.Width) * Screen.TwipsPerPixelX, ScaleY(Picture.Height) * Screen.TwipsPerPixelY
      Else
        Size ScaleX(m_BarPictureBack.Width) * Screen.TwipsPerPixelX, ScaleY(m_BarPictureBack.Height) * Screen.TwipsPerPixelY
        Set Picture = m_BarPictureBack
    End If
    Refresh
      
    On Error GoTo 0
End Sub

Public Sub Refresh()

  Dim AbsValue As Long
  Dim Caption As String
  Dim PicPos As Integer

    Cls
    AbsValue = m_Value - m_Min

    On Error Resume Next ' Width/Height=0
    
    Select Case m_Orientation

      Case 0 ' [Horizontal]
        PicPos = AbsValue / AbsCount * ScaleWidth
        PaintPicture m_BarPicture, 0, 0, PicPos, ScaleHeight, 0, 0, PicPos, ScaleHeight

      Case 1 ' [Vertical]
        PicPos = ScaleHeight - AbsValue / AbsCount * ScaleHeight
        PaintPicture m_BarPicture, 0, PicPos, ScaleWidth, ScaleHeight - PicPos, 0, PicPos, ScaleWidth, ScaleHeight - PicPos
    End Select

    Select Case m_CaptionFormat

      Case 0  ' fNothing
      Case 1  ' fValue
        Caption = m_Value
      Case 2  ' f0%
        Caption = Format(m_Value / AbsCount, "0%")
      Case 3  ' f0.0%
        Caption = Format(m_Value / AbsCount, "0.0%")
      Case 4  ' f0.00%
        Caption = Format(m_Value / AbsCount, "0.00%")
      Case 5  ' fAbs0%
        Caption = Format(AbsValue / AbsCount, "0%")
      Case 6  ' fAbs0.0%
        Caption = Format(AbsValue / AbsCount, "0.0%")
      Case 7  ' fAbs0.00%
        Caption = Format(AbsValue / AbsCount, "0.00%")
      Case 99 ' Custom
        Caption = m_CaptionCustom
    End Select

    CurrentX = (ScaleWidth - TextWidth(Caption)) \ 2
    CurrentY = (ScaleHeight - TextHeight("")) \ 2

    Print Caption;
      
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------------
' Events
'-------------------------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    '-- This enables 2nd. click of DblClick event (Incr./Decr. by one):
    UserControl_MouseDown tmpButton, tmpShift, tmpX, tmpY
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (x >= 0 And x < ScaleWidth And y >= 0 And y < ScaleHeight) Then

        tmpButton = Button
        tmpShift = Shift
        tmpX = x
        tmpY = y
    
        If (Button = vbLeftButton) Then
    
            MouseSelect = True
            UserControl_MouseMove Button, Shift, x, y
    
          ElseIf (Button = vbRightButton) Then
    
            If (GetValue(x, y) > Value) Then
                Value = Value + 1
              ElseIf (GetValue(x, y) < Value) Then
                Value = Value - 1
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Prevents selection of out of range x values
    If (x < 0) Then x = 0 Else If (x >= ScaleWidth) Then x = ScaleWidth
    '-- Prevents selection of out of range y values
    If (y < 0) Then y = 0 Else If (y >= ScaleHeight) Then y = ScaleHeight

    If (MouseSelect) Then Value = GetValue(x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseSelect = False
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'-------------------------------------------------------------------------------------------
' Properties
'-------------------------------------------------------------------------------------------

Public Property Get CaptionCustom() As String
    CaptionCustom = m_CaptionCustom
End Property

Public Property Let CaptionCustom(ByVal New_CaptionCustom As String)
    m_CaptionCustom = New_CaptionCustom
    Refresh
End Property

Public Property Get CaptionFormat() As pbCaptionFormats
    CaptionFormat = m_CaptionFormat
End Property

Public Property Let CaptionFormat(ByVal New_CaptionFormat As pbCaptionFormats)
    m_CaptionFormat = New_CaptionFormat
    Refresh
End Property

Public Property Get BarPicture() As StdPicture
    Set BarPicture = m_BarPicture
End Property

Public Property Set BarPicture(ByVal New_BarPicture As StdPicture)
    Set m_BarPicture = New_BarPicture
    PropertyChanged "BarPicture"
    UserControl_Resize
End Property

Public Property Get BarPictureBack() As StdPicture
    Set BarPictureBack = m_BarPictureBack
End Property

Public Property Set BarPictureBack(ByVal New_BarPictureBack As StdPicture)
    Set m_BarPictureBack = New_BarPictureBack
    UserControl_Resize
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Refresh
End Property

Public Property Get FontColor() As OLE_COLOR
    FontColor = UserControl.ForeColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    UserControl.ForeColor() = New_FontColor
    Refresh
End Property

Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = "General"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    
    If (New_Max <= m_Min) Then Err.Raise 380
    
    m_Max = New_Max
    AbsCount = m_Max - m_Min
    Refresh
End Property

Public Property Get Min() As Long
Attribute Min.VB_ProcData.VB_Invoke_Property = "General"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)

    If (New_Min >= m_Max) Then Err.Raise 380
    
    m_Min = New_Min
    Value = New_Min
    AbsCount = m_Max - m_Min
End Property

Public Property Get Orientation() As pbOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As pbOrientationConstants)
    m_Orientation = New_Orientation
End Property

Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = "General"
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)

    If (New_Value < m_Min Or New_Value > m_Max) Then Err.Raise 380
    
    m_Value = New_Value
    Refresh

    If (m_Value <> LastValue) Then

        LastValue = m_Value

        RaiseEvent ValueChanged
        If (m_Value = m_Max) Then RaiseEvent ArrivedLast
        If (m_Value = m_Min) Then RaiseEvent ArrivedFirst
    End If
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Devuelve o establece el tipo de puntero del mouse mostrado al pasar por encima de un objeto."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
End Property

'-------------------------------------------------------------------------------------------
' Private
'-------------------------------------------------------------------------------------------

Private Function GetValue(x As Single, y As Single) As Long

    Select Case m_Orientation
      Case 0 ' [Horizontal]
        GetValue = (x) / ScaleWidth * AbsCount + m_Min
      Case 1 ' [Vertical]
        GetValue = (ScaleHeight - y) / ScaleHeight * AbsCount + m_Min
    End Select
    
    If (GetValue < m_Min) Then
        GetValue = m_Min
      ElseIf (GetValue > m_Max) Then
        GetValue = m_Max
    End If
End Function
