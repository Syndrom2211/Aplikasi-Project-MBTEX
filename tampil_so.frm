VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form tampil_so 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data SO"
   ClientHeight    =   9225
   ClientLeft      =   165
   ClientTop       =   375
   ClientWidth     =   15855
   Icon            =   "tampil_so.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15855
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H tamso_kembali 
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   8160
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
   Begin lvButton.lvButtons_H tamso_hapus 
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   7320
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
      Image           =   "tampil_so.frx":076A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tamso_refresh 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Refresh"
      CapAlign        =   2
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
      Image           =   "tampil_so.frx":0B04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tamso_edit 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   7320
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
      Image           =   "tampil_so.frx":0E9E
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
      Top             =   6960
      Width           =   3015
   End
   Begin MSComctlLib.ListView LvSO 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   15615
      _ExtentX        =   27543
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1606
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No SO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tgl Kirim Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No SJ Kirim Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Untuk Kirim Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Jenis Kirim Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rol Kirim Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Kg Kirim Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Tgl Hasil Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "No SJ Hasil Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Warna Hasil Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Rol Hasil Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Kg Hasil Celupan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Tgl Kirim Langganan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "No SJ Kirim Langganan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Warna Kirim Langganan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Rol Kirim Langganan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Kg Kirim Langganan"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   4080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "tampil_so.frx":1238
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Data SO"
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
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "tampil_so"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAMPILIN DATA SO
Private Sub Form_Load()
koneksi
Set so = New ADODB.Recordset
    so.Open "select * from tbl_SO", dbkoneksi, adOpenKeyset, adLockOptimistic
LvSO.ListItems.Clear
LvSO.View = lvwReport
        While Not so.EOF
            Set iList = LvSO.ListItems.Add(, , so.Fields(0).Value & "")
                iList.SubItems(1) = so.Fields(1).Value & ""
                iList.SubItems(2) = so.Fields(2).Value & ""
                iList.SubItems(3) = so.Fields(3).Value & ""
                iList.SubItems(4) = so.Fields(4).Value & ""
                iList.SubItems(5) = so.Fields(5).Value & ""
                iList.SubItems(6) = so.Fields(6).Value & ""
                iList.SubItems(7) = so.Fields(7).Value & ""
                iList.SubItems(8) = so.Fields(8).Value & ""
                iList.SubItems(9) = so.Fields(9).Value & ""
                iList.SubItems(10) = so.Fields(10).Value & ""
                iList.SubItems(11) = so.Fields(11).Value & ""
                iList.SubItems(12) = so.Fields(12).Value & ""
                iList.SubItems(13) = so.Fields(13).Value & ""
                iList.SubItems(14) = so.Fields(14).Value & ""
                iList.SubItems(15) = so.Fields(15).Value & ""
                iList.SubItems(16) = so.Fields(16).Value & ""
                iList.SubItems(17) = so.Fields(17).Value & ""
            so.MoveNext
        Wend
End Sub

'EDIT DATA SO
Private Sub tamso_edit_Click()
On Error GoTo salahedit
koneksi
edit = "select * from tbl_SO where no_so = '" & LvSO.SelectedItem.Text & "'"
dbkoneksi.Execute (edit)
frm_so.id_so.Text = LvSO.SelectedItem.Text
frm_so.no_so.Text = LvSO.SelectedItem.SubItems(1)
frm_so.tgl_krmclpn.Text = LvSO.SelectedItem.SubItems(2)
frm_so.nosj_krmclpn.Text = LvSO.SelectedItem.SubItems(3)
frm_so.untuk_krmclpn.Text = LvSO.SelectedItem.SubItems(4)
frm_so.jenis_krmclpn.Text = LvSO.SelectedItem.SubItems(5)
frm_so.rol_krmclpn.Text = LvSO.SelectedItem.SubItems(6)
frm_so.kg_krmclpn.Text = LvSO.SelectedItem.SubItems(7)
frm_so.tgl_hslclpn.Text = LvSO.SelectedItem.SubItems(8)
frm_so.nosj_hslclpn.Text = LvSO.SelectedItem.SubItems(9)
frm_so.warna_hslclpn.Text = LvSO.SelectedItem.SubItems(10)
frm_so.rol_hslclpn.Text = LvSO.SelectedItem.SubItems(11)
frm_so.kg_hslclpn.Text = LvSO.SelectedItem.SubItems(12)
frm_so.tgl_krmlgn.Text = LvSO.SelectedItem.SubItems(13)
frm_so.nosj_krmlgn.Text = LvSO.SelectedItem.SubItems(14)
frm_so.warna_krmlgn.Text = LvSO.SelectedItem.SubItems(15)
frm_so.rol_krmlgn.Text = LvSO.SelectedItem.SubItems(16)
frm_so.kg_krmlgn.Text = LvSO.SelectedItem.SubItems(17)
frm_so.Show
frm_so.edit_so.Enabled = True
frm_so.no_so.Enabled = False
frm_so.tgl_krmclpn.Enabled = False
frm_so.nosj_krmclpn.Enabled = False
frm_so.untuk_krmclpn.Enabled = False
frm_so.jenis_krmclpn.Enabled = False
frm_so.rol_krmclpn.Enabled = False
frm_so.kg_krmclpn.Enabled = False
frm_so.tgl_hslclpn.Enabled = False
frm_so.nosj_hslclpn.Enabled = False
frm_so.warna_hslclpn.Enabled = False
frm_so.rol_hslclpn.Enabled = False
frm_so.kg_hslclpn.Enabled = False
frm_so.tgl_krmlgn.Enabled = False
frm_so.nosj_krmlgn.Enabled = False
frm_so.warna_krmlgn.Enabled = False
frm_so.rol_krmlgn.Enabled = False
frm_so.kg_krmlgn.Enabled = False
frm_so.simpan_so.Enabled = False
frm_so.edit_so.Enabled = True
frm_so.simpan_so.Enabled = False
frm_so.input_so.Enabled = True
frm_so.batal_so.Enabled = False
Exit Sub
salahedit:
MsgBox ("Harap pilih data yang mau di edit ...."), vbInformation, "Pesan"
End Sub

'HAPUS SO
Private Sub tamso_hapus_Click()
If MsgBox("Yakin mau dihapus ?", vbYesNo, "Info") = vbYes Then
koneksi
On Error GoTo salahhapus
delete = "delete from tbl_SO where id_so = " & LvSO.SelectedItem.Text
dbkoneksi.Execute (delete)
MsgBox ("Data Berhasil di Hapus..."), vbInformation, "Success"
Unload Me
tampil_so.Show
End If
Exit Sub
salahhapus:
MsgBox ("Harap pilih data yang mau di hapus ...."), vbInformation, "Pesan"
End Sub

Private Sub tamso_kembali_Click()
Unload Me
End Sub

Private Sub tamso_refresh_Click()
Unload Me
tampil_so.Show
End Sub
