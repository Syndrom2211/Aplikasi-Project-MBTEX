VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form tampil_maklooncelup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Makloon Celup"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   Icon            =   "tampil_maklooncelup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H tammakloon_kembali 
      Height          =   615
      Left            =   1680
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
   Begin lvButton.lvButtons_H tammakloon_hapus 
      Height          =   615
      Left            =   1680
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
      Image           =   "tampil_maklooncelup.frx":076A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tammakloon_refresh 
      Height          =   615
      Left            =   240
      TabIndex        =   4
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
      Image           =   "tampil_maklooncelup.frx":0B04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H tammakloon_edit 
      Height          =   615
      Left            =   240
      TabIndex        =   3
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
      Image           =   "tampil_maklooncelup.frx":0E9E
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
      Top             =   6840
      Width           =   3015
   End
   Begin MSComctlLib.ListView LvMakloon 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
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
      NumItems        =   8
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
         Text            =   "Jenis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dari"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Kg"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "tampil_maklooncelup.frx":1238
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Data Makloon Celup"
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
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "tampil_maklooncelup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAMPILIN DATA MAKLOON CELUP
Private Sub Form_Load()
koneksi
Set maklooncelup = New ADODB.Recordset
    maklooncelup.Open "select * from tbl_maklooncelup", dbkoneksi, adOpenKeyset, adLockOptimistic
LvMakloon.ListItems.Clear
LvMakloon.View = lvwReport
        While Not maklooncelup.EOF
            Set iList = LvMakloon.ListItems.Add(, , maklooncelup.Fields(0).Value & "")
                iList.SubItems(1) = maklooncelup.Fields(1).Value & ""
                iList.SubItems(2) = maklooncelup.Fields(2).Value & ""
                iList.SubItems(3) = maklooncelup.Fields(3).Value & ""
                iList.SubItems(4) = maklooncelup.Fields(4).Value & ""
                iList.SubItems(5) = maklooncelup.Fields(5).Value & ""
                iList.SubItems(6) = maklooncelup.Fields(6).Value & ""
                iList.SubItems(7) = maklooncelup.Fields(7).Value & ""
            maklooncelup.MoveNext
        Wend
End Sub

'EDIT DATA MAKLOON CELUP
Private Sub tammakloon_edit_Click()
On Error GoTo salahedit
koneksi
edit = "select * from tbl_maklooncelup where tgl_mknclp = '" & LvMakloon.SelectedItem.Text & "'"
dbkoneksi.Execute (edit)
frm_maklooncelup.id_makloon.Text = LvMakloon.SelectedItem.Text
frm_maklooncelup.tgl_makloon.Text = LvMakloon.SelectedItem.SubItems(1)
frm_maklooncelup.nopo_makloon.Text = LvMakloon.SelectedItem.SubItems(2)
frm_maklooncelup.nosj_makloon.Text = LvMakloon.SelectedItem.SubItems(3)
frm_maklooncelup.jenis_makloon.Text = LvMakloon.SelectedItem.SubItems(4)
frm_maklooncelup.dari_makloon.Text = LvMakloon.SelectedItem.SubItems(5)
frm_maklooncelup.rol_makloon.Text = LvMakloon.SelectedItem.SubItems(6)
frm_maklooncelup.kg_makloon.Text = LvMakloon.SelectedItem.SubItems(7)
frm_maklooncelup.Show
frm_maklooncelup.edit_makloon.Enabled = True
frm_maklooncelup.tgl_makloon.Enabled = False
frm_maklooncelup.nosj_makloon.Enabled = False
frm_maklooncelup.dari_makloon.Enabled = False
frm_maklooncelup.kg_makloon.Enabled = False
frm_maklooncelup.nopo_makloon.Enabled = False
frm_maklooncelup.jenis_makloon.Enabled = False
frm_maklooncelup.rol_makloon.Enabled = False
frm_maklooncelup.simpan_makloon.Enabled = False
frm_maklooncelup.editsimpan_makloon = False
Exit Sub
salahedit:
MsgBox ("Harap pilih data yang mau di edit ...."), vbInformation, "Pesan"
End Sub

'HAPUS DATA MAKLOON CELUP
Private Sub tammakloon_hapus_Click()
If MsgBox("Yakin mau dihapus ?", vbYesNo, "Info") = vbYes Then
koneksi
On Error GoTo salahhapus
delete = "delete from tbl_maklooncelup where id_mknclp = " & LvMakloon.SelectedItem.Text
dbkoneksi.Execute (delete)
MsgBox ("Data Berhasil di Hapus..."), vbInformation, "Success"
Unload Me
tampil_maklooncelup.Show
End If
Exit Sub
salahhapus:
MsgBox ("Harap pilih data yang mau di hapus ...."), vbInformation, "Pesan"
End Sub

Private Sub tammakloon_kembali_Click()
Unload Me
End Sub

Private Sub tammakloon_refresh_Click()
Unload Me
tampil_maklooncelup.Show
End Sub
