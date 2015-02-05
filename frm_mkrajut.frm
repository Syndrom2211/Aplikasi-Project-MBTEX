VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_mkrajut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input MK Rajut"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   Icon            =   "frm_mkrajut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H mkrajut_kembali 
      Height          =   495
      Left            =   11400
      TabIndex        =   54
      Top             =   7440
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
   Begin lvButton.lvButtons_H mkrajut_batal 
      Height          =   495
      Left            =   9960
      TabIndex        =   53
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Batal"
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
   Begin lvButton.lvButtons_H mkrajut_editsimpan 
      Height          =   615
      Left            =   11160
      TabIndex        =   52
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Simpan"
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
      ImgAlign        =   1
      Image           =   "frm_mkrajut.frx":1CCA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H mkrajut_edit 
      Height          =   615
      Left            =   9480
      TabIndex        =   51
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Edit"
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
      ImgAlign        =   1
      Image           =   "frm_mkrajut.frx":2064
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H mkrajut_simpan 
      Height          =   615
      Left            =   11160
      TabIndex        =   50
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Simpan"
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
      ImgAlign        =   1
      Image           =   "frm_mkrajut.frx":23FE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H mkrajut_input 
      Height          =   615
      Left            =   9480
      TabIndex        =   49
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Input"
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
      ImgAlign        =   1
      Image           =   "frm_mkrajut.frx":2798
      cBack           =   -2147483633
   End
   Begin VB.TextBox total_stokbenang 
      Height          =   285
      Left            =   4560
      TabIndex        =   48
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox harga_stokbenang 
      Height          =   285
      Left            =   1800
      TabIndex        =   47
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox jumlah_stokbenang 
      Height          =   285
      Left            =   1800
      TabIndex        =   46
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Stok Benang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      TabIndex        =   42
      Top             =   6360
      Width           =   8535
      Begin VB.Label Label22 
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Jumlah "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.TextBox mkrajut_id 
      Height          =   285
      Left            =   11400
      TabIndex        =   41
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Main Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   40
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Main Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   39
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox mkrajut_keterangan 
      Height          =   405
      Left            =   9240
      TabIndex        =   38
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox mkrajut_saldo 
      Height          =   405
      Left            =   11040
      TabIndex        =   37
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox mkrajut_bayar 
      Height          =   375
      Left            =   9240
      TabIndex        =   36
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dari Kain :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   5040
      TabIndex        =   20
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox harga_kain 
         Height          =   285
         Left            =   1560
         TabIndex        =   35
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox jumlah_kain 
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox rol_kain 
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox jenis_kain 
         Height          =   285
         Left            =   1560
         TabIndex        =   32
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox nosj_kain 
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox tgl_kain 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Rol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Jenis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "No SJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dari Benang :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox jumlah_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox krg_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox merk_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox lot_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox jenis_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox dari_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox nopo_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox nosj_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox tgl_benang 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Krg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Merk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Lot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Jenis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Dari "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "No PO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No SJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   480
      Picture         =   "frm_mkrajut.frx":2B32
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   29
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   28
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   27
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input MK Rajut"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frm_mkrajut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
tgl_benang.Enabled = False
nosj_benang.Enabled = False
nopo_benang.Enabled = False
dari_benang.Enabled = False
jenis_benang.Enabled = False
lot_benang.Enabled = False
merk_benang.Enabled = False
krg_benang.Enabled = False
jumlah_benang.Enabled = False
tgl_kain.Enabled = False
nosj_kain.Enabled = False
jenis_kain.Enabled = False
rol_kain.Enabled = False
jumlah_kain.Enabled = False
harga_kain.Enabled = False
mkrajut_bayar.Enabled = False
mkrajut_saldo.Enabled = False
mkrajut_keterangan.Enabled = False
mkrajut_simpan.Enabled = False
mkrajut_batal.Enabled = False
mkrajut_edit.Enabled = False
mkrajut_editsimpan.Enabled = False
jumlah_stokbenang.Enabled = False
total_stokbenang.Enabled = False
harga_stokbenang.Enabled = False
End Sub

Private Sub mkrajut_batal_Click()
tgl_benang.Text = ""
nosj_benang.Text = ""
nopo_benang.Text = ""
dari_benang.Text = ""
jenis_benang.Text = ""
lot_benang.Text = ""
merk_benang.Text = ""
krg_benang.Text = ""
jumlah_benang.Text = ""
tgl_kain.Text = ""
nosj_kain.Text = ""
jenis_kain.Text = ""
rol_kain.Text = ""
jumlah_kain.Text = ""
harga_kain.Text = ""
mkrajut_bayar.Text = ""
mkrajut_saldo.Text = ""
mkrajut_keterangan.Text = ""
jumlah_stokbenang.Text = ""
total_stokbenang.Text = ""
harga_stokbenang.Text = ""
mkrajut_input.Enabled = True
mkrajut_simpan.Enabled = False
mkrajut_batal.Enabled = False
tgl_benang.Enabled = False
nosj_benang.Enabled = False
nopo_benang.Enabled = False
dari_benang.Enabled = False
jenis_benang.Enabled = False
lot_benang.Enabled = False
merk_benang.Enabled = False
krg_benang.Enabled = False
jumlah_benang.Enabled = False
tgl_kain.Enabled = False
nosj_kain.Enabled = False
jenis_kain.Enabled = False
rol_kain.Enabled = False
jumlah_kain.Enabled = False
harga_kain.Enabled = False
mkrajut_bayar.Enabled = False
mkrajut_saldo.Enabled = False
mkrajut_keterangan.Enabled = False
jumlah_stokbenang.Enabled = False
total_stokbenang.Enabled = False
harga_stokbenang.Enabled = False
mkrajut_simpan.Enabled = False
mkrajut_batal.Enabled = False
mkrajut_edit.Enabled = False
mkrajut_editsimpan.Enabled = False
End Sub

Private Sub mkrajut_edit_Click()
tgl_benang.Enabled = True
nosj_benang.Enabled = True
nopo_benang.Enabled = True
dari_benang.Enabled = True
jenis_benang.Enabled = True
lot_benang.Enabled = True
merk_benang.Enabled = True
krg_benang.Enabled = True
jumlah_benang.Enabled = True
tgl_kain.Enabled = True
nosj_kain.Enabled = True
jenis_kain.Enabled = True
rol_kain.Enabled = True
jumlah_kain.Enabled = True
harga_kain.Enabled = True
mkrajut_bayar.Enabled = True
mkrajut_saldo.Enabled = True
mkrajut_keterangan.Enabled = True
mkrajut_editsimpan.Enabled = True
mkrajut_edit.Enabled = False
End Sub

'BAGIAN EDIT UNTUK MKRAJUT
Private Sub mkrajut_editsimpan_Click()
koneksi
selek = "SELECT * FROM tbl_mkrajut WHERE id_mkrajut = " & mkrajut_id.Text
Set mkrajut = New ADODB.Recordset
    mkrajut.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
If Not mkrajut.EOF Then
    With mkrajut
        !tgl_bng = tgl_benang
        !no_sj_bng = nosj_benang
        !no_po_bng = nopo_benang
        !dari_bng = dari_benang
        !jenis_bng = jenis_benang
        !lot_bng = lot_benang
        !merk_bng = merk_benang
        !krg_bng = krg_benang
        !jumlah_bng = jumlah_benang
        !tgl_kain = tgl_kain
        !no_sj_kain = nosj_kain
        !jenis_kain = jenis_kain
        !rol_kain = rol_kain
        !jumlah_kain = jumlah_kain
        !harga_kain = harga_kain
        !bayar = mkrajut_bayar
        !saldo = mkrajut_saldo
        !keterangan = mkrajut_keterangan
        .Update
    End With
End If

tgl_benang.Text = ""
nosj_benang.Text = ""
nopo_benang.Text = ""
dari_benang.Text = ""
jenis_benang.Text = ""
lot_benang.Text = ""
merk_benang.Text = ""
krg_benang.Text = ""
jumlah_benang.Text = ""
tgl_kain.Text = ""
nosj_kain.Text = ""
jenis_kain.Text = ""
rol_kain.Text = ""
jumlah_kain.Text = ""
harga_kain.Text = ""
mkrajut_bayar.Text = ""
mkrajut_saldo.Text = ""
mkrajut_keterangan.Text = ""
tgl_benang.Enabled = False
nosj_benang.Enabled = False
nopo_benang.Enabled = False
dari_benang.Enabled = False
jenis_benang.Enabled = False
lot_benang.Enabled = False
merk_benang.Enabled = False
krg_benang.Enabled = False
jumlah_benang.Enabled = False
tgl_kain.Enabled = False
nosj_kain.Enabled = False
jenis_kain.Enabled = False
rol_kain.Enabled = False
jumlah_kain.Enabled = False
harga_kain.Enabled = False
mkrajut_bayar.Enabled = False
mkrajut_saldo.Enabled = False
mkrajut_keterangan.Enabled = False
mkrajut_editsimpan.Enabled = False
MsgBox ("Data Berhasil di Edit..."), vbInformation, "Success"
End Sub

Private Sub mkrajut_input_Click()
tgl_benang.Enabled = True
nosj_benang.Enabled = True
nopo_benang.Enabled = True
dari_benang.Enabled = True
jenis_benang.Enabled = True
lot_benang.Enabled = True
merk_benang.Enabled = True
krg_benang.Enabled = True
jumlah_benang.Enabled = True
tgl_kain.Enabled = True
nosj_kain.Enabled = True
jenis_kain.Enabled = True
rol_kain.Enabled = True
jumlah_kain.Enabled = True
harga_kain.Enabled = True
mkrajut_bayar.Enabled = True
mkrajut_saldo.Enabled = True
mkrajut_keterangan.Enabled = True
mkrajut_simpan.Enabled = True
mkrajut_batal.Enabled = True
jumlah_stokbenang.Enabled = True
harga_stokbenang.Enabled = True
total_stokbenang.Enabled = True
mkrajut_input.Enabled = False
mkrajut_editsimpan.Enabled = False
tgl_benang.Text = ""
nosj_benang.Text = ""
nopo_benang.Text = ""
dari_benang.Text = ""
jenis_benang.Text = ""
lot_benang.Text = ""
merk_benang.Text = ""
krg_benang.Text = ""
jumlah_benang.Text = ""
tgl_kain.Text = ""
nosj_kain.Text = ""
jenis_kain.Text = ""
rol_kain.Text = ""
jumlah_kain.Text = ""
harga_kain.Text = ""
mkrajut_bayar.Text = ""
mkrajut_saldo.Text = ""
mkrajut_keterangan.Text = ""
mkrajut_edit.Enabled = False
End Sub

Private Sub mkrajut_kembali_Click()
Unload Me
End Sub

'TAMBAH DATA KE MKRAJUT
Private Sub mkrajut_simpan_Click()
koneksi
selek = "SELECT * FROM tbl_mkrajut"
Set mkrajut = New ADODB.Recordset
    mkrajut.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With mkrajut
    .AddNew
        !tgl_bng = tgl_benang
        !no_sj_bng = nosj_benang
        !no_po_bng = nopo_benang
        !dari_bng = dari_benang
        !jenis_bng = jenis_benang
        !lot_bng = lot_benang
        !merk_bng = merk_benang
        !krg_bng = krg_benang
        !jumlah_bng = jumlah_benang
        !tgl_kain = tgl_kain
        !no_sj_kain = nosj_kain
        !jenis_kain = jenis_kain
        !rol_kain = rol_kain
        !jumlah_kain = jumlah_kain
        !harga_kain = harga_kain
        !bayar = mkrajut_bayar
        !saldo = mkrajut_saldo
        !keterangan = mkrajut_keterangan
    .Update
End With

'TAMBAH DATA KE STOK BENANG
selek = "SELECT * FROM tbl_stokbenang"
Set stokbenang = New ADODB.Recordset
    stokbenang.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With stokbenang
Dim jmlnya As String
jmlnya = (Val(!jumlah_stokbenang) - Val(jumlah_stokbenang.Text))
    .AddNew
        !jumlah_stokbenang = jmlnya
        !harga_stokbenang = harga_stokbenang
        !total_stokbenang = total_stokbenang
    .Update
End With

tgl_benang.Text = ""
nosj_benang.Text = ""
nopo_benang.Text = ""
dari_benang.Text = ""
jenis_benang.Text = ""
lot_benang.Text = ""
merk_benang.Text = ""
krg_benang.Text = ""
jumlah_benang.Text = ""
tgl_kain.Text = ""
nosj_kain.Text = ""
jenis_kain.Text = ""
rol_kain.Text = ""
jumlah_kain.Text = ""
harga_kain.Text = ""
mkrajut_bayar.Text = ""
mkrajut_saldo.Text = ""
mkrajut_keterangan.Text = ""
jumlah_stokbenang.Text = ""
total_stokbenang.Text = ""
harga_stokbenang.Text = ""
mkrajut_simpan.Enabled = False
mkrajut_input.Enabled = True
mkrajut_batal.Enabled = False
tgl_benang.Enabled = False
nosj_benang.Enabled = False
nopo_benang.Enabled = False
dari_benang.Enabled = False
jenis_benang.Enabled = False
lot_benang.Enabled = False
merk_benang.Enabled = False
krg_benang.Enabled = False
jumlah_benang.Enabled = False
tgl_kain.Enabled = False
nosj_kain.Enabled = False
jenis_kain.Enabled = False
rol_kain.Enabled = False
jumlah_kain.Enabled = False
harga_kain.Enabled = False
mkrajut_bayar.Enabled = False
mkrajut_saldo.Enabled = False
mkrajut_keterangan.Enabled = False
jumlah_stokbenang.Enabled = False
total_stokbenang.Enabled = False
harga_stokbenang.Enabled = False
MsgBox ("Data Berhasil di Simpan..."), vbInformation, "Success"
End Sub
