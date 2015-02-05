VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form tampil_stokbenang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Stok Benang"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   Icon            =   "tampil_stokbenang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H tamstok_kembali 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Kembali"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12632256
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView LvStokbenang 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Jumlah Stok Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Harga Stok Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Total Stok Benang"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "tampil_stokbenang.frx":076A
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Data Stok Benang"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "tampil_stokbenang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAMPILIN DATA STOKBENANG
Private Sub Form_Load()
koneksi
Set stokbenang = New ADODB.Recordset
    stokbenang.Open "select * from tbl_stokbenang", dbkoneksi, adOpenKeyset, adLockOptimistic
LvStokbenang.ListItems.Clear
LvStokbenang.View = lvwReport
        While Not stokbenang.EOF
            Set iList = LvStokbenang.ListItems.Add(, , stokbenang.Fields(0).Value & "")
                iList.SubItems(1) = stokbenang.Fields(1).Value & ""
                iList.SubItems(2) = stokbenang.Fields(2).Value & ""
            stokbenang.MoveNext
        Wend
End Sub

Private Sub tamstok_kembali_Click()
Unload Me
End Sub
