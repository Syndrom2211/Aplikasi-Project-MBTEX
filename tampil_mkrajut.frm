VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form tampil_mkrajut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data MK Rajut"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16035
   Icon            =   "tampil_mkrajut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H tammkrajut_kembali 
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   7920
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
   Begin lvButton.lvButtons_H tammkrajut_hapus 
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   7080
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
      Image           =   "tampil_mkrajut.frx":076A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tammkrajut_refresh 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   7920
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
      Image           =   "tampil_mkrajut.frx":0B04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tammkrajut_edit 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   7080
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
      Image           =   "tampil_mkrajut.frx":0E9E
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Menu"
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
      Top             =   6720
      Width           =   3015
   End
   Begin MSComctlLib.ListView LvMkrajut 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1080
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1606
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tanggal Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "No SJ Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No PO Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dari Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Jenis Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Lot Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Merk Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Krg Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Jumlah Benang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Tanggal Kain"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "No SJ Kain"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Jenis Kain"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Rol Kain"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Jumlah Kain"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Harga Kain"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Bayar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Keterangan"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "tampil_mkrajut.frx":1238
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Data MK Rajut"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "tampil_mkrajut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAMPILIN DATA MKRAJUT
Private Sub Form_Load()
koneksi
Set mkrajut = New ADODB.Recordset
    mkrajut.Open "select * from tbl_mkrajut", dbkoneksi, adOpenKeyset, adLockOptimistic
LvMkrajut.ListItems.Clear
LvMkrajut.View = lvwReport
        While Not mkrajut.EOF
            Set iList = LvMkrajut.ListItems.Add(, , mkrajut.Fields(0).Value & "")
                iList.SubItems(1) = mkrajut.Fields(1).Value & ""
                iList.SubItems(2) = mkrajut.Fields(2).Value & ""
                iList.SubItems(3) = mkrajut.Fields(3).Value & ""
                iList.SubItems(4) = mkrajut.Fields(4).Value & ""
                iList.SubItems(5) = mkrajut.Fields(5).Value & ""
                iList.SubItems(6) = mkrajut.Fields(6).Value & ""
                iList.SubItems(7) = mkrajut.Fields(7).Value & ""
                iList.SubItems(8) = mkrajut.Fields(8).Value & ""
                iList.SubItems(9) = mkrajut.Fields(9).Value & ""
                iList.SubItems(10) = mkrajut.Fields(10).Value & ""
                iList.SubItems(11) = mkrajut.Fields(11).Value & ""
                iList.SubItems(12) = mkrajut.Fields(12).Value & ""
                iList.SubItems(13) = mkrajut.Fields(13).Value & ""
                iList.SubItems(14) = mkrajut.Fields(14).Value & ""
                iList.SubItems(15) = mkrajut.Fields(15).Value & ""
                iList.SubItems(16) = mkrajut.Fields(16).Value & ""
                iList.SubItems(17) = mkrajut.Fields(17).Value & ""
                iList.SubItems(18) = mkrajut.Fields(18).Value & ""
            mkrajut.MoveNext
        Wend
End Sub

'EDIT PEMBELIAN
Private Sub tammkrajut_edit_Click()
On Error GoTo salahedit
koneksi
edit = "select * from tbl_mkrajut where tgl_bng = '" & LvMkrajut.SelectedItem.Text & "'"
dbkoneksi.Execute (edit)
frm_mkrajut.mkrajut_id.Text = LvMkrajut.SelectedItem.Text
frm_mkrajut.tgl_benang.Text = LvMkrajut.SelectedItem.SubItems(1)
frm_mkrajut.nosj_benang.Text = LvMkrajut.SelectedItem.SubItems(2)
frm_mkrajut.nopo_benang.Text = LvMkrajut.SelectedItem.SubItems(3)
frm_mkrajut.dari_benang.Text = LvMkrajut.SelectedItem.SubItems(4)
frm_mkrajut.jenis_benang.Text = LvMkrajut.SelectedItem.SubItems(5)
frm_mkrajut.lot_benang.Text = LvMkrajut.SelectedItem.SubItems(6)
frm_mkrajut.merk_benang.Text = LvMkrajut.SelectedItem.SubItems(7)
frm_mkrajut.krg_benang.Text = LvMkrajut.SelectedItem.SubItems(8)
frm_mkrajut.jumlah_benang.Text = LvMkrajut.SelectedItem.SubItems(9)
frm_mkrajut.tgl_kain.Text = LvMkrajut.SelectedItem.SubItems(10)
frm_mkrajut.nosj_kain.Text = LvMkrajut.SelectedItem.SubItems(11)
frm_mkrajut.jenis_kain.Text = LvMkrajut.SelectedItem.SubItems(12)
frm_mkrajut.rol_kain.Text = LvMkrajut.SelectedItem.SubItems(13)
frm_mkrajut.jumlah_kain.Text = LvMkrajut.SelectedItem.SubItems(14)
frm_mkrajut.harga_kain.Text = LvMkrajut.SelectedItem.SubItems(15)
frm_mkrajut.mkrajut_bayar.Text = LvMkrajut.SelectedItem.SubItems(16)
frm_mkrajut.mkrajut_saldo.Text = LvMkrajut.SelectedItem.SubItems(17)
frm_mkrajut.mkrajut_keterangan.Text = LvMkrajut.SelectedItem.SubItems(18)
frm_mkrajut.Show
frm_mkrajut.mkrajut_edit.Enabled = True
frm_mkrajut.tgl_benang.Enabled = False
frm_mkrajut.nosj_benang.Enabled = False
frm_mkrajut.nopo_benang.Enabled = False
frm_mkrajut.dari_benang.Enabled = False
frm_mkrajut.jenis_benang.Enabled = False
frm_mkrajut.lot_benang.Enabled = False
frm_mkrajut.merk_benang.Enabled = False
frm_mkrajut.krg_benang.Enabled = False
frm_mkrajut.jumlah_benang.Enabled = False
frm_mkrajut.tgl_kain.Enabled = False
frm_mkrajut.nosj_kain.Enabled = False
frm_mkrajut.jenis_kain.Enabled = False
frm_mkrajut.rol_kain.Enabled = False
frm_mkrajut.jumlah_kain.Enabled = False
frm_mkrajut.harga_kain.Enabled = False
frm_mkrajut.mkrajut_bayar.Enabled = False
frm_mkrajut.mkrajut_saldo.Enabled = False
frm_mkrajut.mkrajut_keterangan.Enabled = False
frm_mkrajut.jumlah_stokbenang.Enabled = False
frm_mkrajut.total_stokbenang.Enabled = False
frm_mkrajut.harga_stokbenang.Enabled = False
frm_mkrajut.mkrajut_simpan.Enabled = False
Exit Sub
salahedit:
MsgBox ("Harap pilih data yang mau di edit ...."), vbInformation, "Pesan"
End Sub

'HAPUS DATA MKRAJUT
Private Sub tammkrajut_hapus_Click()
If MsgBox("Yakin mau dihapus ?", vbYesNo, "Info") = vbYes Then
koneksi
On Error GoTo salahhapus
delete = "delete from tbl_mkrajut where id_mkrajut = " & LvMkrajut.SelectedItem.Text
dbkoneksi.Execute (delete)
MsgBox ("Data Berhasil di Hapus..."), vbInformation, "Success"
Unload Me
tampil_mkrajut.Show
End If
Exit Sub
salahhapus:
MsgBox ("Harap pilih data yang mau di hapus ...."), vbInformation, "Pesan"
End Sub

Private Sub tammkrajut_kembali_Click()
Unload Me
End Sub

Private Sub tammkrajut_refresh_Click()
Unload Me
tampil_mkrajut.Show
End Sub
