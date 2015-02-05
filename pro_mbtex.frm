VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm_depan 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Mbtex"
   ClientHeight    =   4710
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7185
   Icon            =   "pro_mbtex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "pro_mbtex.frx":1CCA
   ScaleHeight     =   4710
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   4200
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2566
      _cy             =   1085
   End
   Begin VB.Menu Info 
      Caption         =   "&INFO"
      Begin VB.Menu inf_tentang 
         Caption         =   "Tentang"
      End
      Begin VB.Menu inf_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu input_data 
      Caption         =   "&INPUT DATA"
      Begin VB.Menu inp_pembelian 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu inp_mkrajut 
         Caption         =   "MK Rajut"
      End
      Begin VB.Menu inp_maklooncelup 
         Caption         =   "Makloon Celup"
      End
      Begin VB.Menu inp_stokkaincelupan 
         Caption         =   "Stok Kain Celupan"
      End
      Begin VB.Menu inp_so 
         Caption         =   "SO"
      End
   End
   Begin VB.Menu lihatdata 
      Caption         =   "&LIHAT DATA"
      Begin VB.Menu lih_pembelian 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu lih_mkrajut 
         Caption         =   "MK Rajut"
      End
      Begin VB.Menu lih_maklooncelup 
         Caption         =   "Makloon Celup"
      End
      Begin VB.Menu lih_stokkaincelupan 
         Caption         =   "Stok Kain Celupan"
      End
      Begin VB.Menu lih_so 
         Caption         =   "SO"
      End
      Begin VB.Menu lih_stokbenang 
         Caption         =   "Stok Benang"
      End
   End
End
Attribute VB_Name = "frm_depan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
 
Const LWA_BOTH = 3
Const LWA_ALPHA = 2
Const LWA_COLORKEY = 1
Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000
 
Dim iTransparant As Integer
 
Public Sub MakeTransparan(hWndBro As Long, iTransp As Integer)
    On Error Resume Next
 
    Dim ret As Long
    ret = GetWindowLong(hWndBro, GWL_EXSTYLE)
 
    SetWindowLong hWndBro, GWL_EXSTYLE, ret Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWndBro, RGB(255, 255, 0), iTransp, LWA_ALPHA
    Exit Sub
End Sub
'-------------------------------

Private Sub Form_Load()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer1.Interval = 1
    Timer2.Interval = 1
    Me.Visible = False
    Timer1.Enabled = True
    
'Play a Music
'Welcome
Dim strBuff As String
Dim strFile As String
    'Membuat nama temp file
    strFile = App.Path & "\01f83fc36192e824017ef7a27443-orig.wav"
    
    'Extrak File dari Resource File
    strBuff = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
    
    'Menghapus attribut Read-Only sebelum membuka file untuk output
    If Len(Dir(strFile, vbHidden)) > 0 Then SetAttr strFile, vbNormal
    
    'Save the string as a file
    Open strFile For Output As #1
        Print #1, strBuff
    Close #1
    
    'Menempatkan atrribut lagi setelah menutupnya
    SetAttr strFile, vbArchive + vbHidden
    
    wmp.URL = App.Path & "\01f83fc36192e824017ef7a27443-orig.wav" 'Load a Music
    wmp.Controls.Play 'Mainkan
End Sub

Private Sub Form_Resize()
   If Me.WindowState <> 1 Then Me.WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub inf_exit_Click()
Unload Me
End Sub

Private Sub inf_tentang_Click()
frm_tentang.Show
End Sub

Private Sub inp_maklooncelup_Click()
frm_maklooncelup.Show
End Sub

Private Sub inp_mkrajut_Click()
frm_mkrajut.Show
End Sub

Private Sub inp_pembelian_Click()
frm_pembelian.Show
End Sub

Private Sub inp_so_Click()
frm_so.Show
End Sub

Private Sub inp_stokkaincelupan_Click()
frm_stokkaincelup.Show
End Sub

Private Sub lih_maklooncelup_Click()
tampil_maklooncelup.Show
End Sub

Private Sub lih_mkrajut_Click()
tampil_mkrajut.Show
End Sub

Private Sub lih_pembelian_Click()
tampil_pembelian.Show
End Sub

Private Sub lih_so_Click()
tampil_so.Show
End Sub

Private Sub lih_stokbenang_Click()
tampil_stokbenang.Show
End Sub

Private Sub lih_stokkaincelupan_Click()
tampil_stokkaincelupan.Show
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    iTransparant = iTransparant + 5
    If iTransparant > 255 Then
        iTransparant = 255
        Timer1.Enabled = False
    End If
      MakeTransparan Me.hWnd, iTransparant
    Me.Show
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    iTransparant = iTransparant - 5
    If iTransparant < 0 Then
        iTransparant = 0
        Timer2.Enabled = False
        End
    End If
    MakeTransparan Me.hWnd, iTransparant
End Sub
