VERSION 5.00
Begin VB.Form frmBringOnTop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ToolWindow: BringOnTop Demonstration"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin prjToolWindow.ToolWindow ToolWindow2 
      Height          =   1275
      Left            =   5220
      TabIndex        =   4
      Top             =   3060
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   2249
      ColorNormal     =   13402434
      ColorDark       =   8388608
      ColorLight      =   16761024
      Caption         =   "Window 2"
   End
   Begin prjToolWindow.ToolWindow ToolWindow1 
      Height          =   1185
      Left            =   450
      TabIndex        =   3
      Top             =   360
      Width           =   3885
      _ExtentX        =   6694
      _ExtentY        =   2090
      ColorNormal     =   4699390
      ColorDark       =   33023
      ColorLight      =   903679
      Caption         =   "Window 1"
   End
   Begin VB.CommandButton cmdShowHide1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show/Hide Window 1"
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   4950
      Width           =   1905
   End
   Begin VB.CommandButton cmdShowHide2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show/Hide Window 2"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   5400
      Width           =   1905
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   8190
      TabIndex        =   0
      Top             =   5400
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drag the toolwindows over each other to see the BringOnTop effect."
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
      Left            =   2160
      TabIndex        =   5
      Top             =   5580
      Width           =   5850
   End
End
Attribute VB_Name = "frmBringOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdShowHide1_Click()
    ToolWindow1.Visible = Not ToolWindow1.Visible
End Sub

Private Sub cmdShowHide2_Click()
    ToolWindow2.Visible = Not ToolWindow2.Visible
End Sub
