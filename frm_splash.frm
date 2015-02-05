VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   3360
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   3870
      Left            =   0
      Picture         =   "frm_splash.frx":0000
      Top             =   0
      Width           =   7950
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim efek As Integer

Private Sub Timer1_Timer()
On Error Resume Next
efek = efek + 5
ProgressBar1.Value = ProgressBar1.Value + 400 / 400

If efek > 500 Then
    Timer1.Enabled = False
    Screen.MousePointer = vbNormal
    Me.WindowState = 0
    Do
    Me.Left = Me.Left + 20
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Left > Screen.Width
    Load frm_depan
    frm_depan.Show
    Unload Me
End If
End Sub

