VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00CC8142&
   Caption         =   "ToolWindow Demonstration"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBringOnTop 
      Caption         =   "BringOnTop Demonstration"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   5490
      Width           =   2265
   End
   Begin VB.CommandButton cmdShowHide2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show/Hide Window 2"
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   5490
      Width           =   1905
   End
   Begin VB.CommandButton cmdShowHide1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show/Hide Window 1"
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   5040
      Width           =   1905
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   9180
      TabIndex        =   3
      Top             =   5490
      Width           =   1275
   End
   Begin prjToolWindow.ToolWindow ToolWindow1 
      Height          =   1275
      Left            =   4320
      TabIndex        =   0
      Top             =   630
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2249
      Caption         =   "ToolWindow 1"
      DefaultX        =   20
      DefaultY        =   20
      BringOnTop      =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   2085
      Left            =   180
      ScaleHeight     =   2025
      ScaleWidth      =   3285
      TabIndex        =   2
      Top             =   2070
      Width           =   3345
   End
   Begin prjToolWindow.ToolWindow ToolWindow2 
      Height          =   1275
      Left            =   4320
      TabIndex        =   1
      Top             =   2160
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2249
      Caption         =   "ToolWindow 2"
      BringOnTop      =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Try dragging the toolwindows over the picturebox."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2070
      TabIndex        =   7
      Top             =   5580
      Width           =   4320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBringOnTop_Click()
    frmBringOnTop.Show vbModal
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdShowHide1_Click()
    ToolWindow1.Visible = Not ToolWindow1.Visible
End Sub

Private Sub cmdShowHide2_Click()
    ToolWindow2.Visible = Not ToolWindow2.Visible
End Sub
