VERSION 5.00
Begin VB.UserControl ToolWindow 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1125
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ToolWindow.ctx":0000
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ ToolWindow ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   4560
   End
   Begin prjToolWindow.LineButton cmdClose 
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ ToolWindow ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   15
      Width           =   4560
   End
End
Attribute VB_Name = "ToolWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Created by Faraz Azhar [http://www.geocities.com/farazazhar_net/]
'
Public Enum twStartupMode
    twNone
    twManual
End Enum

Dim mX As Single, mY As Single
Dim m_DefaultX As Single, m_DefaultY As Single, m_Validate As Boolean
Dim m_Startup As twStartupMode, m_Enabled As Boolean, m_BringOnTop As Boolean
Dim m_Light As OLE_COLOR, m_Normal As OLE_COLOR, m_Dark As OLE_COLOR
'
Public Event WindowDragged()
Public Event Closed()

Private Sub cmdClose_Click()
    UserControl.Extender.Visible = False
    RaiseEvent Closed
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BringOnTop Then UserControl.Extender.ZOrder 0
    '
    If Index = 0 Then
        If Button = vbLeftButton Then
            mX = X
            mY = Y
        End If
    End If
End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lLeft As Single, lTop As Single, bDragged As Boolean
    If Index = 0 Then
        If Button = vbLeftButton Then
            '
            lLeft = UserControl.Extender.Left - (mX - X)
            lTop = UserControl.Extender.Top - (mY - Y)
            ' validate
            If m_Validate Then
                If (lLeft >= 0) And ((lLeft + ScaleWidth) <= UserControl.Parent.ScaleWidth) Then
                    UserControl.Extender.Left = lLeft
                    bDragged = True
                End If
                If (lTop >= 0) And ((lTop + cmdClose.Height) <= UserControl.Parent.ScaleHeight) Then
                    UserControl.Extender.Top = lTop
                    bDragged = True
                End If
                If bDragged Then RaiseEvent WindowDragged
            Else
                ' no validation
                UserControl.Extender.Left = lLeft
                UserControl.Extender.Top = lTop
                RaiseEvent WindowDragged
            End If
        End If
    End If
End Sub

Private Sub UserControl_GotFocus()
    If BringOnTop Then UserControl.Extender.ZOrder 0
End Sub

Private Sub UserControl_Initialize()
    m_DefaultX = 150
    m_DefaultY = 150
    m_Normal = vbButtonFace
    m_Light = vb3DLight
    m_Dark = vb3DShadow
    m_BringOnTop = True
    m_Enabled = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BringOnTop Then UserControl.Extender.ZOrder 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ColorNormal = .ReadProperty("ColorNormal", vbButtonFace)
        ColorLight = .ReadProperty("ColorLight", vb3DLight)
        ColorDark = .ReadProperty("ColorDark", vb3DShadow)
        Caption = .ReadProperty("Caption", "Message")
        DefaultX = .ReadProperty("DefaultX", 150)
        DefaultY = .ReadProperty("DefaultY", 150)
        ValidateMovements = .ReadProperty("ValidateMovements ", True)
        StartupMode = .ReadProperty("StartupMode", twNone)
        Enabled = .ReadProperty("Enabled", True)
        BringOnTop = .ReadProperty("BringOnTop", True)
    End With
    '
    If (UserControl.Ambient.UserMode) And (StartupMode = twManual) Then
        UserControl.Extender.Left = DefaultX
        UserControl.Extender.Top = DefaultY
        'RaiseEvent WindowDragged   ' no need for this.
    End If
    '
End Sub

Private Sub UserControl_Resize()
    lblCaption(0).Width = ScaleWidth
    lblCaption(1).Width = ScaleWidth
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ColorNormal", ColorNormal, vbButtonFace
        .WriteProperty "ColorDark", ColorDark, vb3DShadow
        .WriteProperty "ColorLight", ColorLight, vb3DLight
        .WriteProperty "Caption", Caption, "ToolWindow"
        .WriteProperty "DefaultX", DefaultX, 150
        .WriteProperty "DefaultY", DefaultY, 150
        .WriteProperty "ValidateMovements", ValidateMovements, True
        .WriteProperty "StartupMode", StartupMode, twNone
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "BringOnTop", BringOnTop, True
    End With
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Specifies the title for the toolwindow."
Attribute Caption.VB_UserMemId = -518
    Caption = lblCaption(0)
End Property

Public Property Let Caption(ByVal vNewValue As String)
    lblCaption(0) = vNewValue
    lblCaption(1) = vNewValue
    PropertyChanged "Caption"
End Property

Public Property Get DefaultX() As Single
    DefaultX = m_DefaultX
End Property

Public Property Let DefaultX(ByVal vNewValue As Single)
    m_DefaultX = vNewValue
    PropertyChanged "DefaultX"
End Property

Public Property Get DefaultY() As Single
    DefaultY = m_DefaultY
End Property

Public Property Let DefaultY(ByVal vNewValue As Single)
    m_DefaultY = vNewValue
    PropertyChanged "DefaultY"
End Property

Public Property Get ValidateMovements() As Boolean
    ValidateMovements = m_Validate
End Property

Public Property Let ValidateMovements(ByVal vNewValue As Boolean)
    m_Validate = vNewValue
    PropertyChanged "ValidateMovements"
End Property

Public Property Get StartupMode() As twStartupMode
    StartupMode = m_Startup
End Property

Public Property Let StartupMode(ByVal vNewValue As twStartupMode)
    m_Startup = vNewValue
    PropertyChanged "StartupMode"
End Property

Public Property Get ColorNormal() As OLE_COLOR
Attribute ColorNormal.VB_UserMemId = -501
    ColorNormal = m_Normal
End Property

Public Property Let ColorNormal(ByVal vNewValue As OLE_COLOR)
    m_Normal = vNewValue
    cmdClose.ColorNormal = vNewValue
    BackColor = vNewValue
    PropertyChanged "ColorNormal"
End Property

Public Property Get ColorLight() As OLE_COLOR
    ColorLight = m_Light
End Property

Public Property Let ColorLight(ByVal vNewValue As OLE_COLOR)
    m_Light = vNewValue
    cmdClose.ColorLight = vNewValue
    PropertyChanged "ColorLight"
End Property

Public Property Get ColorDark() As OLE_COLOR
Attribute ColorDark.VB_UserMemId = -513
    ColorDark = m_Dark
End Property

Public Property Let ColorDark(ByVal vNewValue As OLE_COLOR)
    m_Dark = vNewValue
    cmdClose.ColorDark = vNewValue
    lblCaption(0).BackColor = vNewValue
    lblCaption(1).BackColor = vNewValue
    PropertyChanged "ColorDark"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_Enabled = vNewValue
    cmdClose.Enabled = vNewValue
    PropertyChanged "Enabled"
End Property

Public Property Get BringOnTop() As Boolean
    BringOnTop = m_BringOnTop
End Property

Public Property Let BringOnTop(ByVal vNewValue As Boolean)
    m_BringOnTop = vNewValue
    PropertyChanged "BringOnTop"
End Property
