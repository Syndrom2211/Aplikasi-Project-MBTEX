VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form tampil_pembelian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Pembelian"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16185
   Icon            =   "tampil_pembelian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   16185
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H tampem_kembali 
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
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
   Begin lvButton.lvButtons_H tampem_hapus 
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Hapus"
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
      Image           =   "tampil_pembelian.frx":076A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tampem_edit 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   7200
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
      Image           =   "tampil_pembelian.frx":0B04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tampem_refresh 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Refresh"
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
      Image           =   "tampil_pembelian.frx":0E9E
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Menu "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   3255
   End
   Begin MSComctlLib.ListView LvPembelian 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   9551
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1606
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tanggal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "No PO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No SJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nama"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Jenis Barang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "No Lot"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Jumlah"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Harga"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Retur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Bayar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Bank / No Giro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Keterangan"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "tampil_pembelian.frx":1238
      Top             =   240
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   4800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Data Pembelian"
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
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "tampil_pembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAMPILIN DATA PEMBELIAN
Private Sub Form_Load()
koneksi
Set pembelian = New ADODB.Recordset
    pembelian.Open "select * from tbl_pembelian", dbkoneksi, adOpenKeyset, adLockOptimistic
LvPembelian.ListItems.Clear
LvPembelian.View = lvwReport
        While Not pembelian.EOF
            Set iList = LvPembelian.ListItems.Add(, , pembelian.Fields(0).Value & "")
                iList.SubItems(1) = pembelian.Fields(1).Value & ""
                iList.SubItems(2) = pembelian.Fields(2).Value & ""
                iList.SubItems(3) = pembelian.Fields(3).Value & ""
                iList.SubItems(4) = pembelian.Fields(4).Value & ""
                iList.SubItems(5) = pembelian.Fields(5).Value & ""
                iList.SubItems(6) = pembelian.Fields(6).Value & ""
                iList.SubItems(7) = pembelian.Fields(7).Value & ""
                iList.SubItems(8) = pembelian.Fields(8).Value & ""
                iList.SubItems(9) = pembelian.Fields(9).Value & ""
                iList.SubItems(10) = pembelian.Fields(10).Value & ""
                iList.SubItems(11) = pembelian.Fields(11).Value & ""
                iList.SubItems(12) = pembelian.Fields(12).Value & ""
                iList.SubItems(13) = pembelian.Fields(13).Value & ""
                iList.SubItems(14) = pembelian.Fields(14).Value & ""
            pembelian.MoveNext
        Wend
        
End Sub

'EDIT PEMBELIAN
Private Sub tampem_edit_Click()
On Error GoTo salahedit
koneksi
edit = "select * from tbl_pembelian where tgl_pbl = '" & LvPembelian.SelectedItem.Text & "'"
dbkoneksi.Execute (edit)
frm_pembelian.pem_id.Text = LvPembelian.SelectedItem.Text
frm_pembelian.pem_tgl.Text = LvPembelian.SelectedItem.SubItems(1)
frm_pembelian.pem_nopo.Text = LvPembelian.SelectedItem.SubItems(2)
frm_pembelian.pem_nosj.Text = LvPembelian.SelectedItem.SubItems(3)
frm_pembelian.pem_nama.Text = LvPembelian.SelectedItem.SubItems(4)
frm_pembelian.pem_jenisbarang.Text = LvPembelian.SelectedItem.SubItems(5)
frm_pembelian.pem_nolot.Text = LvPembelian.SelectedItem.SubItems(6)
frm_pembelian.pem_jumlah.Text = LvPembelian.SelectedItem.SubItems(7)
frm_pembelian.pem_harga.Text = LvPembelian.SelectedItem.SubItems(8)
frm_pembelian.pem_total.Text = LvPembelian.SelectedItem.SubItems(9)
frm_pembelian.pem_retur.Text = LvPembelian.SelectedItem.SubItems(10)
frm_pembelian.pem_bayar.Text = LvPembelian.SelectedItem.SubItems(11)
frm_pembelian.pem_saldo.Text = LvPembelian.SelectedItem.SubItems(12)
frm_pembelian.pem_banknogaji.Text = LvPembelian.SelectedItem.SubItems(13)
frm_pembelian.pem_keterangan.Text = LvPembelian.SelectedItem.SubItems(14)
frm_pembelian.Show
frm_pembelian.pem_cmd_edit.Enabled = True
frm_pembelian.pem_cmd_simpan.Enabled = False
frm_pembelian.pem_cmd_editsimpan.Enabled = False
frm_pembelian.pem_tgl.Enabled = False
frm_pembelian.pem_nopo.Enabled = False
frm_pembelian.pem_nosj.Enabled = False
frm_pembelian.pem_jumlah.Enabled = False
frm_pembelian.pem_total.Enabled = False
frm_pembelian.pem_bayar.Enabled = False
frm_pembelian.pem_banknogaji.Enabled = False
frm_pembelian.pem_nama.Enabled = False
frm_pembelian.pem_jenisbarang.Enabled = False
frm_pembelian.pem_nolot.Enabled = False
frm_pembelian.pem_harga.Enabled = False
frm_pembelian.pem_retur.Enabled = False
frm_pembelian.pem_saldo.Enabled = False
frm_pembelian.pem_keterangan.Enabled = False
Exit Sub
salahedit:
MsgBox ("Harap pilih data yang mau di edit ...."), vbInformation, "Pesan"
End Sub

'HAPUS PEMBELIAN
Private Sub tampem_hapus_Click()
If MsgBox("Yakin mau dihapus ?", vbYesNo, "Info") = vbYes Then
koneksi
On Error GoTo salahhapus
delete = "delete from tbl_pembelian where id_pbl = " & LvPembelian.SelectedItem.Text
dbkoneksi.Execute (delete)
MsgBox ("Data Berhasil di Hapus..."), vbInformation, "Success"
Unload Me
tampil_pembelian.Show
End If
Exit Sub
salahhapus:
MsgBox ("Harap pilih data yang mau di hapus ...."), vbInformation, "Pesan"
End Sub

Private Sub tampem_kembali_Click()
Unload Me
frm_depan.Show
End Sub

Private Sub tampem_refresh_Click()
Unload Me
tampil_pembelian.Show
End Sub
