VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_pembelian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Pembelian"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8820
   Icon            =   "frm_pembelian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H pem_cmd_kembali 
      Height          =   495
      Left            =   4680
      TabIndex        =   37
      Top             =   7680
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
   Begin lvButton.lvButtons_H pem_cmd_batal 
      Height          =   495
      Left            =   3120
      TabIndex        =   36
      Top             =   7680
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
   Begin lvButton.lvButtons_H pem_cmd_editsimpan 
      Height          =   615
      Left            =   6600
      TabIndex        =   35
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
      Image           =   "frm_pembelian.frx":1CCA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H pem_cmd_edit 
      Height          =   615
      Left            =   5280
      TabIndex        =   34
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
      Image           =   "frm_pembelian.frx":2064
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H pem_cmd_input 
      Height          =   615
      Left            =   1200
      TabIndex        =   33
      Top             =   6360
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
      Image           =   "frm_pembelian.frx":23FE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H pem_cmd_simpan 
      Height          =   615
      Left            =   2520
      TabIndex        =   32
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
      Image           =   "frm_pembelian.frx":2798
      cBack           =   -2147483633
   End
   Begin VB.TextBox pem_id 
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
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
      Height          =   1575
      Left            =   720
      TabIndex        =   30
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Frame Frame1 
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
      Height          =   1575
      Left            =   4800
      TabIndex        =   29
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox pem_keterangan 
      Height          =   375
      Left            =   6360
      TabIndex        =   28
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox pem_saldo 
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox pem_retur 
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox pem_harga 
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox pem_nolot 
      Height          =   405
      Left            =   6360
      TabIndex        =   24
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox pem_jenisbarang 
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox pem_nama 
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox pem_banknogaji 
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox pem_bayar 
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox pem_total 
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox pem_jumlah 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox pem_nosj 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox pem_nopo 
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox pem_tgl 
      Height          =   405
      Left            =   2040
      TabIndex        =   15
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   720
      Picture         =   "frm_pembelian.frx":2B32
      Top             =   240
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1560
      X2              =   3240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label15 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank No. Gaji"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label13 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Retur"
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
      Left            =   4920
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "No Lot"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Barang"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
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
      Left            =   4920
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembelian"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frm_pembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
pem_tgl.Enabled = False
pem_nopo.Enabled = False
pem_nosj.Enabled = False
pem_jumlah.Enabled = False
pem_total.Enabled = False
pem_bayar.Enabled = False
pem_banknogaji.Enabled = False
pem_nama.Enabled = False
pem_jenisbarang.Enabled = False
pem_nolot.Enabled = False
pem_harga.Enabled = False
pem_retur.Enabled = False
pem_saldo.Enabled = False
pem_keterangan.Enabled = False
pem_cmd_simpan.Enabled = False
pem_cmd_batal.Enabled = False
pem_cmd_edit.Enabled = False
pem_cmd_editsimpan.Enabled = False
End Sub

Private Sub pem_cmd_batal_Click()
pem_tgl.Text = ""
pem_nopo.Text = ""
pem_nosj.Text = ""
pem_jumlah.Text = ""
pem_total.Text = ""
pem_bayar.Text = ""
pem_banknogaji.Text = ""
pem_nama.Text = ""
pem_jenisbarang.Text = ""
pem_nolot.Text = ""
pem_harga.Text = ""
pem_retur.Text = ""
pem_saldo.Text = ""
pem_keterangan.Text = ""
pem_cmd_input.Enabled = True
pem_cmd_simpan.Enabled = False
pem_cmd_batal.Enabled = False
pem_tgl.Enabled = False
pem_nopo.Enabled = False
pem_nosj.Enabled = False
pem_jumlah.Enabled = False
pem_total.Enabled = False
pem_bayar.Enabled = False
pem_banknogaji.Enabled = False
pem_nama.Enabled = False
pem_jenisbarang.Enabled = False
pem_nolot.Enabled = False
pem_harga.Enabled = False
pem_retur.Enabled = False
pem_saldo.Enabled = False
pem_keterangan.Enabled = False
pem_cmd_simpan.Enabled = False
pem_cmd_batal.Enabled = False
pem_cmd_editsimpan.Enabled = False
End Sub

Private Sub pem_cmd_edit_Click()
pem_tgl.Enabled = True
pem_nopo.Enabled = True
pem_nosj.Enabled = True
pem_jumlah.Enabled = True
pem_total.Enabled = True
pem_bayar.Enabled = True
pem_banknogaji.Enabled = True
pem_nama.Enabled = True
pem_jenisbarang.Enabled = True
pem_nolot.Enabled = True
pem_harga.Enabled = True
pem_retur.Enabled = True
pem_saldo.Enabled = True
pem_keterangan.Enabled = True
pem_cmd_edit.Enabled = False
pem_cmd_editsimpan.Enabled = True
pem_cmd_batal.Enabled = True
End Sub

'BAGIAN EDIT UNTUK PEMBELIAN
Private Sub pem_cmd_editsimpan_Click()
koneksi
selek = "SELECT * FROM tbl_pembelian WHERE id_pbl = " & pem_id.Text
Set pembelian = New ADODB.Recordset
    pembelian.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
If Not pembelian.EOF Then
    With pembelian
        !tgl_pbl = pem_tgl
        !no_po_pbl = pem_nopo
        !no_sj_pbl = pem_nosj
        !nama_pbl = pem_nama
        !jnsbar_pbl = pem_jenisbarang
        !no_lot_pbl = pem_nolot
        !jumlah_pbl = pem_jumlah
        !harga_pbl = pem_harga
        !total_pbl = pem_total
        !retur_pbl = pem_retur
        !bayar_pbl = pem_bayar
        !saldo_pbl = pem_saldo
        !bank_no_giro_pbl = pem_banknogaji
        !keterangan_pbl = pem_keterangan
        .Update
    End With
End If
pem_tgl.Text = ""
pem_nopo.Text = ""
pem_nosj.Text = ""
pem_jumlah.Text = ""
pem_total.Text = ""
pem_bayar.Text = ""
pem_banknogaji.Text = ""
pem_nama.Text = ""
pem_jenisbarang.Text = ""
pem_nolot.Text = ""
pem_harga.Text = ""
pem_retur.Text = ""
pem_saldo.Text = ""
pem_keterangan.Text = ""
pem_cmd_input.Enabled = True
pem_cmd_editsimpan.Enabled = False
pem_cmd_input.Enabled = True
pem_cmd_simpan.Enabled = False
pem_cmd_batal.Enabled = False
pem_tgl.Enabled = False
pem_nopo.Enabled = False
pem_nosj.Enabled = False
pem_jumlah.Enabled = False
pem_total.Enabled = False
pem_bayar.Enabled = False
pem_banknogaji.Enabled = False
pem_nama.Enabled = False
pem_jenisbarang.Enabled = False
pem_nolot.Enabled = False
pem_harga.Enabled = False
pem_retur.Enabled = False
pem_saldo.Enabled = False
pem_keterangan.Enabled = False
MsgBox ("Data Berhasil di Edit..."), vbInformation, "Success"
End Sub

Private Sub pem_cmd_input_Click()
pem_tgl.Enabled = True
pem_nopo.Enabled = True
pem_nosj.Enabled = True
pem_jumlah.Enabled = True
pem_total.Enabled = True
pem_bayar.Enabled = True
pem_banknogaji.Enabled = True
pem_nama.Enabled = True
pem_jenisbarang.Enabled = True
pem_nolot.Enabled = True
pem_harga.Enabled = True
pem_retur.Enabled = True
pem_saldo.Enabled = True
pem_keterangan.Enabled = True
pem_cmd_simpan.Enabled = True
pem_cmd_batal.Enabled = True
pem_cmd_input.Enabled = False
pem_cmd_edit.Enabled = False
pem_cmd_editsimpan.Enabled = False
pem_tgl.Text = ""
pem_nopo.Text = ""
pem_nosj.Text = ""
pem_jumlah.Text = ""
pem_total.Text = ""
pem_bayar.Text = ""
pem_banknogaji.Text = ""
pem_nama.Text = ""
pem_jenisbarang.Text = ""
pem_nolot.Text = ""
pem_harga.Text = ""
pem_retur.Text = ""
pem_saldo.Text = ""
pem_keterangan.Text = ""
End Sub

Private Sub pem_cmd_kembali_Click()
Unload Me
End Sub

'TAMBAH DATA KE PEMBELIAN
Private Sub pem_cmd_simpan_Click()
koneksi
selek = "SELECT * FROM tbl_pembelian"
Set pembelian = New ADODB.Recordset
    pembelian.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With pembelian
    .AddNew
        !tgl_pbl = pem_tgl
        !no_po_pbl = pem_nopo
        !no_sj_pbl = pem_nosj
        !nama_pbl = pem_nama
        !jnsbar_pbl = pem_jenisbarang
        !no_lot_pbl = pem_nolot
        !jumlah_pbl = pem_jumlah
        !harga_pbl = pem_harga
        !total_pbl = pem_total
        !retur_pbl = pem_retur
        !bayar_pbl = pem_bayar
        !saldo_pbl = pem_saldo
        !bank_no_giro_pbl = pem_banknogaji
        !keterangan_pbl = pem_keterangan
    .Update
End With

'TAMBAH DATA KE STOK BENANG
selek = "SELECT * FROM tbl_stokbenang"
Set stokbenang = New ADODB.Recordset
    stokbenang.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With stokbenang
    .AddNew
        !jumlah_stokbenang = pem_jumlah
        !harga_stokbenang = pem_harga
        !total_stokbenang = pem_total
    .Update
End With

pem_tgl.Text = ""
pem_nopo.Text = ""
pem_nosj.Text = ""
pem_jumlah.Text = ""
pem_total.Text = ""
pem_bayar.Text = ""
pem_banknogaji.Text = ""
pem_nama.Text = ""
pem_jenisbarang.Text = ""
pem_nolot.Text = ""
pem_harga.Text = ""
pem_retur.Text = ""
pem_saldo.Text = ""
pem_keterangan.Text = ""
pem_cmd_simpan.Enabled = False
pem_cmd_input.Enabled = True
pem_cmd_edit.Enabled = False
pem_tgl.Enabled = False
pem_nopo.Enabled = False
pem_nosj.Enabled = False
pem_jumlah.Enabled = False
pem_total.Enabled = False
pem_bayar.Enabled = False
pem_banknogaji.Enabled = False
pem_nama.Enabled = False
pem_jenisbarang.Enabled = False
pem_nolot.Enabled = False
pem_harga.Enabled = False
pem_retur.Enabled = False
pem_saldo.Enabled = False
pem_keterangan.Enabled = False
pem_cmd_batal.Enabled = False
MsgBox ("Data Berhasil di Simpan..."), vbInformation, "Success"
End Sub
