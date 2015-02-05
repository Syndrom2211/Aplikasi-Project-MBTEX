VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_maklooncelup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Makloon Celup"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   Icon            =   "frm_maklooncelup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H kembali_makloon 
      Height          =   495
      Left            =   5280
      TabIndex        =   23
      Top             =   4800
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
   Begin lvButton.lvButtons_H batal_makloon 
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      Top             =   4800
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
   Begin lvButton.lvButtons_H editsimpan_makloon 
      Height          =   615
      Left            =   5160
      TabIndex        =   21
      Top             =   3720
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
      Image           =   "frm_maklooncelup.frx":1CCA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H edit_makloon 
      Height          =   615
      Left            =   3840
      TabIndex        =   20
      Top             =   3720
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
      Image           =   "frm_maklooncelup.frx":2064
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H simpan_makloon 
      Height          =   615
      Left            =   2040
      TabIndex        =   19
      Top             =   3720
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
      Image           =   "frm_maklooncelup.frx":23FE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H input_makloon 
      Height          =   615
      Left            =   720
      TabIndex        =   18
      Top             =   3720
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
      Image           =   "frm_maklooncelup.frx":2798
      cBack           =   -2147483633
   End
   Begin VB.TextBox id_makloon 
      Height          =   285
      Left            =   5640
      TabIndex        =   17
      Top             =   120
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
      Height          =   1335
      Left            =   3720
      TabIndex        =   16
      Top             =   3240
      Width           =   2775
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
      Height          =   1335
      Left            =   600
      TabIndex        =   15
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox rol_makloon 
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox jenis_makloon 
      Height          =   285
      Left            =   4920
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox nopo_makloon 
      Height          =   285
      Left            =   4920
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox kg_makloon 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox dari_makloon 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox nosj_makloon 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox tgl_makloon 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   5520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   720
      Picture         =   "frm_maklooncelup.frx":2B32
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dari"
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
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
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
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   855
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
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   975
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
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Makloon Celup"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frm_maklooncelup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batal_makloon_Click()
tgl_makloon.Text = ""
nosj_makloon.Text = ""
dari_makloon.Text = ""
kg_makloon.Text = ""
nopo_makloon.Text = ""
jenis_makloon.Text = ""
rol_makloon.Text = ""
tgl_makloon.Enabled = False
nosj_makloon.Enabled = False
dari_makloon.Enabled = False
kg_makloon.Enabled = False
nopo_makloon.Enabled = False
jenis_makloon.Enabled = False
rol_makloon.Enabled = False
input_makloon.Enabled = True
simpan_makloon.Enabled = False
batal_makloon.Enabled = False
editsimpan_makloon.Enabled = False
End Sub

Private Sub edit_makloon_Click()
tgl_makloon.Enabled = True
nosj_makloon.Enabled = True
dari_makloon.Enabled = True
kg_makloon.Enabled = True
nopo_makloon.Enabled = True
jenis_makloon.Enabled = True
rol_makloon.Enabled = True
input_makloon.Enabled = False
editsimpan_makloon.Enabled = True
edit_makloon.Enabled = False
batal_makloon.Enabled = True
End Sub

'BAGIAN EDIT UNTUK MAKLOON CELUP
Private Sub editsimpan_makloon_Click()
koneksi
selek = "SELECT * FROM tbl_maklooncelup WHERE id_mknclp = " & id_makloon.Text
Set maklooncelup = New ADODB.Recordset
    maklooncelup.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
If Not maklooncelup.EOF Then
    With maklooncelup
        !tgl_mknclp = tgl_makloon
        !no_po_mknclp = nopo_makloon
        !no_sj_mknclp = nosj_makloon
        !jenis_mknclp = jenis_makloon
        !dari_mknclp = dari_makloon
        !rol_mknclp = rol_makloon
        !kg_mknclp = kg_makloon
        .Update
    End With
End If

tgl_makloon.Text = ""
nosj_makloon.Text = ""
dari_makloon.Text = ""
kg_makloon.Text = ""
nopo_makloon.Text = ""
jenis_makloon.Text = ""
rol_makloon.Text = ""
tgl_makloon.Enabled = False
nosj_makloon.Enabled = False
dari_makloon.Enabled = False
kg_makloon.Enabled = False
nopo_makloon.Enabled = False
jenis_makloon.Enabled = False
rol_makloon.Enabled = False
input_makloon.Enabled = True
editsimpan_makloon.Enabled = False
batal_makloon.Enabled = False
MsgBox ("Data Berhasil di Edit..."), vbInformation, "Success"
End Sub

Private Sub Form_Load()
tgl_makloon.Enabled = False
nosj_makloon.Enabled = False
dari_makloon.Enabled = False
kg_makloon.Enabled = False
nopo_makloon.Enabled = False
jenis_makloon.Enabled = False
rol_makloon.Enabled = False
simpan_makloon.Enabled = False
batal_makloon.Enabled = False
edit_makloon.Enabled = False
editsimpan_makloon.Enabled = False
End Sub

Private Sub input_makloon_Click()
tgl_makloon.Enabled = True
nosj_makloon.Enabled = True
dari_makloon.Enabled = True
kg_makloon.Enabled = True
nopo_makloon.Enabled = True
jenis_makloon.Enabled = True
rol_makloon.Enabled = True
simpan_makloon.Enabled = True
batal_makloon.Enabled = True
input_makloon.Enabled = False
End Sub

Private Sub keluar_makloon_Click()
Unload Me
frm_depan.Show
End Sub

Private Sub kembali_makloon_Click()
Unload Me
End Sub

'TAMBAH DATA KE MAKLOON CELUP
Private Sub simpan_makloon_Click()
koneksi
selek = "SELECT * FROM tbl_maklooncelup"
Set maklooncelup = New ADODB.Recordset
    maklooncelup.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With maklooncelup
    .AddNew
        !tgl_mknclp = tgl_makloon
        !no_po_mknclp = nopo_makloon
        !no_sj_mknclp = nosj_makloon
        !jenis_mknclp = jenis_makloon
        !dari_mknclp = dari_makloon
        !rol_mknclp = rol_makloon
        !kg_mknclp = kg_makloon
    .Update
End With
tgl_makloon.Text = ""
nosj_makloon.Text = ""
dari_makloon.Text = ""
kg_makloon.Text = ""
nopo_makloon.Text = ""
jenis_makloon.Text = ""
rol_makloon.Text = ""
simpan_makloon.Enabled = False
batal_makloon.Enabled = False
input_makloon.Enabled = True
tgl_makloon.Enabled = False
nosj_makloon.Enabled = False
dari_makloon.Enabled = False
kg_makloon.Enabled = False
nopo_makloon.Enabled = False
jenis_makloon.Enabled = False
rol_makloon.Enabled = False
MsgBox ("Data Berhasil di Simpan..."), vbInformation, "Success"
End Sub
