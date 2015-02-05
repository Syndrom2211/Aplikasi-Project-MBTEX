VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form tampil_stokkaincelupan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Stok Kain Celupan"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16035
   Icon            =   "tampil_stokkaincelupan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H tamstok_kembali 
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   8280
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
   Begin lvButton.lvButtons_H tamstok_hapus 
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   7440
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
      Image           =   "tampil_stokkaincelupan.frx":076A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tamstok_refresh 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   8280
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
      Image           =   "tampil_stokkaincelupan.frx":0B04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tamstok_edit 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   7440
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
      Image           =   "tampil_stokkaincelupan.frx":0E9E
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
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   3375
   End
   Begin MSComctlLib.ListView LvStokkain 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1440
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
         Text            =   "No Faktur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Jenis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Warna"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "No Warna"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Kg"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Harga"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Total"
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
         Text            =   "Bank No Giro"
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
      Left            =   240
      Picture         =   "tampil_stokkaincelupan.frx":1238
      Top             =   360
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   6840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Data Stok Kain Celupan"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "tampil_stokkaincelupan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAMPILIN DATA STOK KAIN CELUPAN
Private Sub Form_Load()
koneksi
Set stokkaincelupan = New ADODB.Recordset
    stokkaincelupan.Open "select * from tbl_stokkaincelupan", dbkoneksi, adOpenKeyset, adLockOptimistic
LvStokkain.ListItems.Clear
LvStokkain.View = lvwReport
        While Not stokkaincelupan.EOF
            Set iList = LvStokkain.ListItems.Add(, , stokkaincelupan.Fields(0).Value & "")
                iList.SubItems(1) = stokkaincelupan.Fields(1).Value & ""
                iList.SubItems(2) = stokkaincelupan.Fields(2).Value & ""
                iList.SubItems(3) = stokkaincelupan.Fields(3).Value & ""
                iList.SubItems(4) = stokkaincelupan.Fields(4).Value & ""
                iList.SubItems(5) = stokkaincelupan.Fields(5).Value & ""
                iList.SubItems(6) = stokkaincelupan.Fields(6).Value & ""
                iList.SubItems(7) = stokkaincelupan.Fields(7).Value & ""
            stokkaincelupan.MoveNext
        Wend
End Sub

'EDIT DATA STOK KAIN CELUPAN
Private Sub tamstok_edit_Click()
On Error GoTo salahedit
koneksi
edit = "select * from tbl_stokkaincelupan where tgl_skc = '" & LvStokkain.SelectedItem.Text & "'"
dbkoneksi.Execute (edit)
frm_stokkaincelup.id_stokkain.Text = LvStokkain.SelectedItem.Text
frm_stokkaincelup.tgl_stokkain.Text = LvStokkain.SelectedItem.SubItems(1)
frm_stokkaincelup.nopo_stokkain.Text = LvStokkain.SelectedItem.SubItems(2)
frm_stokkaincelup.nofaktur_stokkain.Text = LvStokkain.SelectedItem.SubItems(3)
frm_stokkaincelup.jeniskain_stokkain.Text = LvStokkain.SelectedItem.SubItems(4)
frm_stokkaincelup.warna_stokkain.Text = LvStokkain.SelectedItem.SubItems(5)
frm_stokkaincelup.nowarna_stokkain.Text = LvStokkain.SelectedItem.SubItems(6)
frm_stokkaincelup.rol_stokkain.Text = LvStokkain.SelectedItem.SubItems(7)
frm_stokkaincelup.kg_stokkain.Text = LvStokkain.SelectedItem.SubItems(8)
frm_stokkaincelup.harga_stokkain.Text = LvStokkain.SelectedItem.SubItems(9)
frm_stokkaincelup.total_stokkain.Text = LvStokkain.SelectedItem.SubItems(10)
frm_stokkaincelup.bayar_stokkain.Text = LvStokkain.SelectedItem.SubItems(11)
frm_stokkaincelup.saldo_stokkain.Text = LvStokkain.SelectedItem.SubItems(12)
frm_stokkaincelup.nogiro_stokkain.Text = LvStokkain.SelectedItem.SubItems(13)
frm_stokkaincelup.keterangan_stokkain.Text = LvStokkain.SelectedItem.SubItems(14)
frm_stokkaincelup.Show
frm_stokkaincelup.edit_stokkain.Enabled = True
frm_stokkaincelup.tgl_stokkain.Enabled = False
frm_stokkaincelup.nopo_stokkain.Enabled = False
frm_stokkaincelup.nofaktur_stokkain.Enabled = False
frm_stokkaincelup.jeniskain_stokkain.Enabled = False
frm_stokkaincelup.warna_stokkain.Enabled = False
frm_stokkaincelup.nowarna_stokkain.Enabled = False
frm_stokkaincelup.rol_stokkain.Enabled = False
frm_stokkaincelup.kg_stokkain.Enabled = False
frm_stokkaincelup.harga_stokkain.Enabled = False
frm_stokkaincelup.total_stokkain.Enabled = False
frm_stokkaincelup.bayar_stokkain.Enabled = False
frm_stokkaincelup.saldo_stokkain.Enabled = False
frm_stokkaincelup.nogiro_stokkain.Enabled = False
frm_stokkaincelup.keterangan_stokkain.Enabled = False
frm_stokkaincelup.simpan_stokkain.Enabled = False
Exit Sub
salahedit:
MsgBox ("Harap pilih data yang mau di edit ...."), vbInformation, "Pesan"
End Sub

'HAPUS PEMBELIAN
Private Sub tamstok_hapus_Click()
If MsgBox("Yakin mau dihapus ?", vbYesNo, "Info") = vbYes Then
koneksi
On Error GoTo salahhapus
delete = "delete from tbl_stokkaincelupan where id_skc = " & LvStokkain.SelectedItem.Text
dbkoneksi.Execute (delete)
MsgBox ("Data Berhasil di Hapus..."), vbInformation, "Success"
Unload Me
tampil_stokkaincelupan.Show
End If
Exit Sub
salahhapus:
MsgBox ("Harap pilih data yang mau di hapus ...."), vbInformation, "Pesan"
End Sub

Private Sub tamstok_kembali_Click()
Unload Me
End Sub

Private Sub tamstok_refresh_Click()
Unload Me
tampil_stokkaincelupan.Show
End Sub
