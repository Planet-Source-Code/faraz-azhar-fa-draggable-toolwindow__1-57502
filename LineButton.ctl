VERSION 5.00
Begin VB.UserControl LineButton 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   70
   ToolboxBitmap   =   "LineButton.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Label lblMask 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Line lLine 
      Index           =   0
      X1              =   18
      X2              =   36
      Y1              =   6
      Y2              =   6
   End
   Begin VB.Line lLine 
      Index           =   1
      X1              =   6
      X2              =   6
      Y1              =   30
      Y2              =   18
   End
   Begin VB.Line lLine 
      Index           =   3
      X1              =   42
      X2              =   18
      Y1              =   42
      Y2              =   42
   End
   Begin VB.Line lLine 
      Index           =   2
      X1              =   48
      X2              =   48
      Y1              =   30
      Y2              =   12
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   105
   End
End
Attribute VB_Name = "LineButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Created by Faraz Azhar [http://www.geocities.com/farazazhar_net/]
'
Dim m_Light As OLE_COLOR, m_Normal As OLE_COLOR, m_Dark As OLE_COLOR
Dim m_Enabled As Boolean

Public Event Click()
Attribute Click.VB_UserMemId = -600

Private Sub DrawWindow(Optional bDown As Boolean)
    '
    Cls
    '
    UserControl.BackColor = m_Normal
    '
    Dim CLR1 As Long, CLR2 As Long
    CLR1 = IIf(bDown, m_Dark, m_Light)
    CLR2 = IIf(bDown, m_Light, m_Dark)
    '
    ' NOTE: The following coordinates are for vbPixel scalemode only.
    SetLine 0, 0, 0, ScaleWidth, 0, CLR1
    SetLine 1, 0, 0, 0, ScaleHeight - 1, CLR1
    SetLine 2, ScaleWidth - 1, 1, ScaleWidth - 1, ScaleHeight - 1, CLR2
    SetLine 3, 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, CLR2
    '
    lblCaption.Left = Int((ScaleWidth - lblCaption.Width) / 2)
    lblCaption.Top = Int((ScaleHeight - lblCaption.Height) / 2)
    '
End Sub

Private Sub SetLine(Index, X1, Y1, X2, Y2, Clr)
    With lLine(Index)
        .X1 = X1
        .Y1 = Y1
        .X2 = X2
        .Y2 = Y2
        .BorderColor = Clr
    End With
End Sub

Private Sub lblMask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
End Sub

Private Sub lblMask_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Initialize()
    m_Normal = vbButtonFace
    m_Light = vb3DLight
    m_Dark = vb3DShadow
    m_Enabled = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled = False Then Exit Sub
    DrawWindow True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled = False Then Exit Sub
    DrawWindow
    '
    If ((X >= 0) And (X <= ScaleWidth)) And ((Y >= 0) And (Y <= ScaleHeight)) Then
        RaiseEvent Click
    End If
    '
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ColorNormal = .ReadProperty("ColorNormal", vbButtonFace)
        ColorLight = .ReadProperty("ColorLight", vb3DLight)
        ColorDark = .ReadProperty("ColorDark", vb3DShadow)
        Caption = .ReadProperty("Caption", "x")
        Set lblCaption.Font = .ReadProperty("Font", UserControl.Font)
        Enabled = .ReadProperty("Enabled", True)
    End With
End Sub

Private Sub UserControl_Resize()
    lblMask.Move 0, 0, ScaleWidth, ScaleHeight
    DrawWindow
End Sub

Private Sub UserControl_Show()
    DrawWindow
End Sub

Public Property Get ColorNormal() As OLE_COLOR
Attribute ColorNormal.VB_UserMemId = -501
    ColorNormal = m_Normal
End Property

Public Property Let ColorNormal(ByVal vNewValue As OLE_COLOR)
    m_Normal = vNewValue
    PropertyChanged "ColorNormal"
    DrawWindow
End Property

Public Property Get ColorLight() As OLE_COLOR
Attribute ColorLight.VB_UserMemId = -513
    ColorLight = m_Light
End Property

Public Property Let ColorLight(ByVal vNewValue As OLE_COLOR)
    m_Light = vNewValue
    PropertyChanged "ColorLight"
    DrawWindow
End Property

Public Property Get ColorDark() As OLE_COLOR
Attribute ColorDark.VB_UserMemId = -510
    ColorDark = m_Dark
End Property

Public Property Let ColorDark(ByVal vNewValue As OLE_COLOR)
    m_Dark = vNewValue
    PropertyChanged "ColorDark"
    DrawWindow
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ColorNormal", ColorNormal, vbButtonFace
        .WriteProperty "ColorDark", ColorDark, vb3DShadow
        .WriteProperty "ColorLight", ColorLight, vb3DLight
        .WriteProperty "Caption", Caption, "x"
        .WriteProperty "Font", Font, UserControl.Font
        .WriteProperty "Enabled", Enabled, True
    End With
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = lblCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    lblCaption = vNewValue
    PropertyChanged "Caption"
    DrawWindow
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    Set lblCaption.Font = vNewValue
    PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_Enabled = vNewValue
    PropertyChanged "Enabled"
    lblCaption.ForeColor = IIf(m_Enabled, vbBlack, vb3DDKShadow)
End Property
