VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_stokkaincelup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Stok Kain Celupan"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7605
   Icon            =   "frm_stokkaincelup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H kembali_stokkain 
      Height          =   495
      Left            =   5760
      TabIndex        =   37
      Top             =   6000
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
   Begin lvButton.lvButtons_H batal_stokkain 
      Height          =   495
      Left            =   4440
      TabIndex        =   36
      Top             =   6000
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
   Begin lvButton.lvButtons_H editsimpan_stokkain 
      Height          =   615
      Left            =   5640
      TabIndex        =   35
      Top             =   4920
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
      Image           =   "frm_stokkaincelup.frx":1CCA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H edit_stokkain 
      Height          =   615
      Left            =   4200
      TabIndex        =   34
      Top             =   4920
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
      Image           =   "frm_stokkaincelup.frx":2064
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H simpan_stokkain 
      Height          =   615
      Left            =   2280
      TabIndex        =   33
      Top             =   4920
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
      Image           =   "frm_stokkaincelup.frx":23FE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H input_stokkain 
      Height          =   615
      Left            =   840
      TabIndex        =   32
      Top             =   4920
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
      Image           =   "frm_stokkaincelup.frx":2798
      cBack           =   -2147483633
   End
   Begin VB.TextBox id_stokkain 
      Height          =   285
      Left            =   6120
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
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
      Left            =   3960
      TabIndex        =   30
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Frame Frame1 
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
      Left            =   600
      TabIndex        =   29
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox keterangan_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   28
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox nogiro_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   27
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox saldo_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   26
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox bayar_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   25
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox total_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   24
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox harga_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   23
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox kg_stokkain 
      Height          =   285
      Left            =   5400
      TabIndex        =   22
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox rol_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox nowarna_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox warna_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox jeniskain_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox nofaktur_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox nopo_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox tgl_stokkain 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   6120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   720
      Picture         =   "frm_stokkaincelup.frx":2B32
      Top             =   120
      Width           =   720
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
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank / No Giro"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   3600
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
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   3120
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
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
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
      Left            =   3960
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      Left            =   600
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "No Warna"
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
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Warna"
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
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kain"
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
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur"
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
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
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
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1680
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
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Stok Kain Celupan"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frm_stokkaincelup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batal_stokkain_Click()
tgl_stokkain.Text = ""
nopo_stokkain.Text = ""
nofaktur_stokkain.Text = ""
jeniskain_stokkain.Text = ""
warna_stokkain.Text = ""
nowarna_stokkain.Text = ""
rol_stokkain.Text = ""
kg_stokkain.Text = ""
harga_stokkain.Text = ""
total_stokkain.Text = ""
bayar_stokkain.Text = ""
saldo_stokkain.Text = ""
nogiro_stokkain.Text = ""
keterangan_stokkain.Text = ""
tgl_stokkain.Enabled = False
nopo_stokkain.Enabled = False
nofaktur_stokkain.Enabled = False
jeniskain_stokkain.Enabled = False
warna_stokkain.Enabled = False
nowarna_stokkain.Enabled = False
rol_stokkain.Enabled = False
kg_stokkain.Enabled = False
harga_stokkain.Enabled = False
total_stokkain.Enabled = False
bayar_stokkain.Enabled = False
saldo_stokkain.Enabled = False
nogiro_stokkain.Enabled = False
keterangan_stokkain.Enabled = False
input_stokkain.Enabled = True
simpan_stokkain.Enabled = False
batal_stokkain.Enabled = False
End Sub

Private Sub edit_stokkain_Click()
tgl_stokkain.Enabled = True
nopo_stokkain.Enabled = True
nofaktur_stokkain.Enabled = True
jeniskain_stokkain.Enabled = True
warna_stokkain.Enabled = True
nowarna_stokkain.Enabled = True
rol_stokkain.Enabled = True
kg_stokkain.Enabled = True
harga_stokkain.Enabled = True
total_stokkain.Enabled = True
bayar_stokkain.Enabled = True
saldo_stokkain.Enabled = True
nogiro_stokkain.Enabled = True
keterangan_stokkain.Enabled = True
editsimpan_stokkain.Enabled = True
input_stokkain.Enabled = False
simpan_stokkain.Enabled = False
edit_stokkain.Enabled = False
batal_stokkain.Enabled = True
End Sub

'BAGIAN EDIT UNTUK STOK KAIN CELUPAN
Private Sub editsimpan_stokkain_Click()
koneksi
selek = "SELECT * FROM tbl_stokkaincelupan WHERE id_skc = " & id_stokkain.Text
Set stokkaincelupan = New ADODB.Recordset
    stokkaincelupan.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
If Not stokkaincelupan.EOF Then
    With stokkaincelupan
        !tgl_skc = tgl_stokkain
        !no_po_skc = nopo_stokkain
        !no_faktur_skc = nofaktur_stokkain
        !jenis_kain_skc = jeniskain_stokkain
        !warna_skc = warna_stokkain
        !no_warna_skc = nowarna_stokkain
        !rol_skc = rol_stokkain
        !kg_skc = kg_stokkain
        !harga_skc = harga_stokkain
        !total_skc = total_stokkain
        !bayar_skc = bayar_stokkain
        !saldo_skc = saldo_stokkain
        !bank_nogiro_skc = nogiro_stokkain
        !keterangan_skc = keterangan_stokkain
        .Update
    End With
End If

tgl_stokkain.Text = ""
nopo_stokkain.Text = ""
nofaktur_stokkain.Text = ""
jeniskain_stokkain.Text = ""
warna_stokkain.Text = ""
nowarna_stokkain.Text = ""
rol_stokkain.Text = ""
kg_stokkain.Text = ""
harga_stokkain.Text = ""
total_stokkain.Text = ""
bayar_stokkain.Text = ""
saldo_stokkain.Text = ""
nogiro_stokkain.Text = ""
keterangan_stokkain.Text = ""
tgl_stokkain.Enabled = False
nopo_stokkain.Enabled = False
nofaktur_stokkain.Enabled = False
jeniskain_stokkain.Enabled = False
warna_stokkain.Enabled = False
nowarna_stokkain.Enabled = False
rol_stokkain.Enabled = False
kg_stokkain.Enabled = False
harga_stokkain.Enabled = False
total_stokkain.Enabled = False
bayar_stokkain.Enabled = False
saldo_stokkain.Enabled = False
nogiro_stokkain.Enabled = False
keterangan_stokkain.Enabled = False
editsimpan_stokkain.Enabled = False
input_stokkain.Enabled = True
batal_stokkain.Enabled = False
MsgBox ("Data Berhasil di Edit..."), vbInformation, "Success"
End Sub

Private Sub Form_Load()
tgl_stokkain.Enabled = False
nopo_stokkain.Enabled = False
nofaktur_stokkain.Enabled = False
jeniskain_stokkain.Enabled = False
warna_stokkain.Enabled = False
nowarna_stokkain.Enabled = False
rol_stokkain.Enabled = False
kg_stokkain.Enabled = False
harga_stokkain.Enabled = False
total_stokkain.Enabled = False
bayar_stokkain.Enabled = False
saldo_stokkain.Enabled = False
nogiro_stokkain.Enabled = False
keterangan_stokkain.Enabled = False
simpan_stokkain.Enabled = False
batal_stokkain.Enabled = False
edit_stokkain.Enabled = False
editsimpan_stokkain.Enabled = False
End Sub

Private Sub input_stokkain_Click()
tgl_stokkain.Enabled = True
nopo_stokkain.Enabled = True
nofaktur_stokkain.Enabled = True
jeniskain_stokkain.Enabled = True
warna_stokkain.Enabled = True
nowarna_stokkain.Enabled = True
rol_stokkain.Enabled = True
kg_stokkain.Enabled = True
harga_stokkain.Enabled = True
total_stokkain.Enabled = True
bayar_stokkain.Enabled = True
saldo_stokkain.Enabled = True
nogiro_stokkain.Enabled = True
keterangan_stokkain.Enabled = True
input_stokkain.Enabled = False
simpan_stokkain.Enabled = True
batal_stokkain.Enabled = True
tgl_stokkain.Text = ""
nopo_stokkain.Text = ""
nofaktur_stokkain.Text = ""
jeniskain_stokkain.Text = ""
warna_stokkain.Text = ""
nowarna_stokkain.Text = ""
rol_stokkain.Text = ""
kg_stokkain.Text = ""
harga_stokkain.Text = ""
total_stokkain.Text = ""
bayar_stokkain.Text = ""
saldo_stokkain.Text = ""
nogiro_stokkain.Text = ""
keterangan_stokkain.Text = ""
edit_stokkain.Enabled = False
editsimpan_stokkain.Enabled = False
End Sub

Private Sub kembali_stokkain_Click()
Unload Me
End Sub

'TAMBAH DATA KE STOK KAIN CELUPAN
Private Sub simpan_stokkain_Click()
koneksi
selek = "SELECT * FROM tbl_stokkaincelupan"
Set stokkaincelupan = New ADODB.Recordset
    stokkaincelupan.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With stokkaincelupan
    .AddNew
        !tgl_skc = tgl_stokkain
        !no_po_skc = nopo_stokkain
        !no_faktur_skc = nofaktur_stokkain
        !jenis_kain_skc = jeniskain_stokkain
        !warna_skc = warna_stokkain
        !no_warna_skc = nowarna_stokkain
        !rol_skc = rol_stokkain
        !kg_skc = kg_stokkain
        !harga_skc = harga_stokkain
        !total_skc = total_stokkain
        !bayar_skc = bayar_stokkain
        !saldo_skc = saldo_stokkain
        !bank_nogiro_skc = nogiro_stokkain
        !keterangan_skc = keterangan_stokkain
    .Update
End With

tgl_stokkain.Text = ""
nopo_stokkain.Text = ""
nofaktur_stokkain.Text = ""
jeniskain_stokkain.Text = ""
warna_stokkain.Text = ""
nowarna_stokkain.Text = ""
rol_stokkain.Text = ""
kg_stokkain.Text = ""
harga_stokkain.Text = ""
total_stokkain.Text = ""
bayar_stokkain.Text = ""
saldo_stokkain.Text = ""
nogiro_stokkain.Text = ""
keterangan_stokkain.Text = ""
tgl_stokkain.Enabled = False
nopo_stokkain.Enabled = False
nofaktur_stokkain.Enabled = False
jeniskain_stokkain.Enabled = False
warna_stokkain.Enabled = False
nowarna_stokkain.Enabled = False
rol_stokkain.Enabled = False
kg_stokkain.Enabled = False
harga_stokkain.Enabled = False
total_stokkain.Enabled = False
bayar_stokkain.Enabled = False
saldo_stokkain.Enabled = False
nogiro_stokkain.Enabled = False
keterangan_stokkain.Enabled = False
simpan_stokkain.Enabled = False
batal_stokkain.Enabled = False
input_stokkain.Enabled = True
MsgBox ("Data Berhasil di Simpan..."), vbInformation, "Success"
End Sub
