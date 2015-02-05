VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_so 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input SO"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11235
   Icon            =   "frm_so.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H kembali_so 
      Height          =   495
      Left            =   6600
      TabIndex        =   46
      Top             =   6600
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
   Begin lvButton.lvButtons_H batal_so 
      Height          =   495
      Left            =   6600
      TabIndex        =   45
      Top             =   5880
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
   Begin lvButton.lvButtons_H editsimpan_so 
      Height          =   615
      Left            =   5040
      TabIndex        =   44
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frm_so.frx":1CCA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H edit_so 
      Height          =   615
      Left            =   3840
      TabIndex        =   43
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frm_so.frx":2064
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H simpan_so 
      Height          =   615
      Left            =   1920
      TabIndex        =   42
      Top             =   6240
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
      Image           =   "frm_so.frx":23FE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H input_so 
      Height          =   615
      Left            =   720
      TabIndex        =   41
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frm_so.frx":2798
      cBack           =   -2147483633
   End
   Begin VB.TextBox id_so 
      Height          =   285
      Left            =   3720
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame5 
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
      Left            =   3600
      TabIndex        =   39
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Frame Frame4 
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
      Left            =   480
      TabIndex        =   38
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Kirim ke Langganan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   7440
      TabIndex        =   5
      Top             =   1800
      Width           =   3255
      Begin VB.TextBox kg_krmlgn 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox rol_krmlgn 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox warna_krmlgn 
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox nosj_krmlgn 
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox tgl_krmlgn 
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label18 
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
         Left            =   240
         TabIndex        =   32
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label17 
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
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label16 
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
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label15 
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
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
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
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hasil dari Celupan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
      Begin VB.TextBox kg_hslclpn 
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox rol_hslclpn 
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox warna_hslclpn 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox nosj_hslclpn 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox tgl_hslclpn 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label12 
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
         TabIndex        =   21
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label11 
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
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label10 
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
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
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
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kirim ke Celupan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
      Begin VB.TextBox kg_krmclpn 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox rol_krmclpn 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox jenis_krmclpn 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox untuk_krmclpn 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox nosj_krmclpn 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox tgl_krmclpn 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label7 
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
         TabIndex        =   10
         Top             =   2760
         Width           =   615
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Untuk"
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
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
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
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
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
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox no_so 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   480
      Picture         =   "frm_so.frx":2B32
      Top             =   240
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No SO "
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
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input SO"
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
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frm_so"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batal_so_Click()
no_so.Text = ""
tgl_krmclpn.Text = ""
nosj_krmclpn.Text = ""
untuk_krmclpn.Text = ""
jenis_krmclpn.Text = ""
rol_krmclpn.Text = ""
kg_krmclpn.Text = ""
tgl_hslclpn.Text = ""
nosj_hslclpn.Text = ""
warna_hslclpn.Text = ""
rol_hslclpn.Text = ""
kg_hslclpn.Text = ""
tgl_krmlgn.Text = ""
nosj_krmlgn.Text = ""
warna_krmlgn.Text = ""
rol_krmlgn.Text = ""
kg_krmlgn.Text = ""
no_so.Enabled = False
tgl_krmclpn.Enabled = False
nosj_krmclpn.Enabled = False
untuk_krmclpn.Enabled = False
editsimpan_so.Enabled = False
edit_so.Enabled = False
jenis_krmclpn.Enabled = False
rol_krmclpn.Enabled = False
kg_krmclpn.Enabled = False
tgl_hslclpn.Enabled = False
nosj_hslclpn.Enabled = False
warna_hslclpn.Enabled = False
rol_hslclpn.Enabled = False
kg_hslclpn.Enabled = False
tgl_krmlgn.Enabled = False
nosj_krmlgn.Enabled = False
warna_krmlgn.Enabled = False
rol_krmlgn.Enabled = False
kg_krmlgn.Enabled = False
simpan_so.Enabled = False
batal_so.Enabled = False
input_so.Enabled = True
End Sub

Private Sub edit_so_Click()
edit_so.Enabled = False
input_so.Enabled = False
simpan_so.Enabled = False
editsimpan_so.Enabled = True
batal_so.Enabled = True
no_so.Enabled = True
tgl_krmclpn.Enabled = True
nosj_krmclpn.Enabled = True
untuk_krmclpn.Enabled = True
jenis_krmclpn.Enabled = True
rol_krmclpn.Enabled = True
kg_krmclpn.Enabled = True
tgl_hslclpn.Enabled = True
nosj_hslclpn.Enabled = True
warna_hslclpn.Enabled = True
rol_hslclpn.Enabled = True
kg_hslclpn.Enabled = True
tgl_krmlgn.Enabled = True
nosj_krmlgn.Enabled = True
warna_krmlgn.Enabled = True
rol_krmlgn.Enabled = True
kg_krmlgn.Enabled = True
End Sub

'BAGIAN EDIT SO
Private Sub editsimpan_so_Click()
koneksi
selek = "SELECT * FROM tbl_so WHERE id_so = " & id_so.Text
Set so = New ADODB.Recordset
    so.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
If Not so.EOF Then
    With so
        !no_so = no_so
        !tgl_so_krmclpn = tgl_krmclpn
        !no_sj_so_krmclpn = nosj_krmclpn
        !untuk_so_krmclpn = untuk_krmclpn
        !jenis_so_krmclpn = jenis_krmclpn
        !rol_so_krmclpn = rol_krmclpn
        !kg_so_krmclpn = kg_krmclpn
        !tgl_so_hslclpn = tgl_hslclpn
        !no_sj_so_hslclpn = nosj_hslclpn
        !warna_so_hslclpn = warna_hslclpn
        !rol_so_hslclpn = rol_hslclpn
        !kg_so_hslclpn = kg_hslclpn
        !tgl_so_langganan = tgl_krmlgn
        !no_sj_so_langganan = nosj_krmlgn
        !warna_so_langganan = warna_krmlgn
        !rol_so_langganan = rol_krmlgn
        !kg_so_langganan = kg_krmlgn
        .Update
    End With
End If

no_so.Text = ""
tgl_krmclpn.Text = ""
nosj_krmclpn.Text = ""
untuk_krmclpn.Text = ""
jenis_krmclpn.Text = ""
rol_krmclpn.Text = ""
kg_krmclpn.Text = ""
tgl_hslclpn.Text = ""
nosj_hslclpn.Text = ""
warna_hslclpn.Text = ""
rol_hslclpn.Text = ""
kg_hslclpn.Text = ""
tgl_krmlgn.Text = ""
nosj_krmlgn.Text = ""
warna_krmlgn.Text = ""
rol_krmlgn.Text = ""
kg_krmlgn.Text = ""
no_so.Enabled = False
tgl_krmclpn.Enabled = False
nosj_krmclpn.Enabled = False
untuk_krmclpn.Enabled = False
editsimpan_so.Enabled = False
edit_so.Enabled = False
jenis_krmclpn.Enabled = False
rol_krmclpn.Enabled = False
kg_krmclpn.Enabled = False
tgl_hslclpn.Enabled = False
nosj_hslclpn.Enabled = False
warna_hslclpn.Enabled = False
rol_hslclpn.Enabled = False
kg_hslclpn.Enabled = False
tgl_krmlgn.Enabled = False
nosj_krmlgn.Enabled = False
warna_krmlgn.Enabled = False
rol_krmlgn.Enabled = False
kg_krmlgn.Enabled = False
editsimpan_so.Enabled = False
input_so.Enabled = True
MsgBox ("Data Berhasil di Edit..."), vbInformation, "Success"
End Sub

Private Sub Form_Load()
no_so.Enabled = False
tgl_krmclpn.Enabled = False
nosj_krmclpn.Enabled = False
untuk_krmclpn.Enabled = False
jenis_krmclpn.Enabled = False
rol_krmclpn.Enabled = False
kg_krmclpn.Enabled = False
tgl_hslclpn.Enabled = False
nosj_hslclpn.Enabled = False
warna_hslclpn.Enabled = False
rol_hslclpn.Enabled = False
kg_hslclpn.Enabled = False
tgl_krmlgn.Enabled = False
nosj_krmlgn.Enabled = False
warna_krmlgn.Enabled = False
rol_krmlgn.Enabled = False
kg_krmlgn.Enabled = False
simpan_so.Enabled = False
batal_so.Enabled = False
edit_so.Enabled = False
editsimpan_so.Enabled = False
End Sub

Private Sub input_so_Click()
no_so.Enabled = True
tgl_krmclpn.Enabled = True
nosj_krmclpn.Enabled = True
untuk_krmclpn.Enabled = True
jenis_krmclpn.Enabled = True
rol_krmclpn.Enabled = True
kg_krmclpn.Enabled = True
tgl_hslclpn.Enabled = True
nosj_hslclpn.Enabled = True
warna_hslclpn.Enabled = True
rol_hslclpn.Enabled = True
kg_hslclpn.Enabled = True
tgl_krmlgn.Enabled = True
nosj_krmlgn.Enabled = True
warna_krmlgn.Enabled = True
rol_krmlgn.Enabled = True
kg_krmlgn.Enabled = True
simpan_so.Enabled = True
batal_so.Enabled = True
input_so.Enabled = False
edit_so.Enabled = False
no_so.Text = ""
tgl_krmclpn.Text = ""
nosj_krmclpn.Text = ""
untuk_krmclpn.Text = ""
jenis_krmclpn.Text = ""
rol_krmclpn.Text = ""
kg_krmclpn.Text = ""
tgl_hslclpn.Text = ""
nosj_hslclpn.Text = ""
warna_hslclpn.Text = ""
rol_hslclpn.Text = ""
kg_hslclpn.Text = ""
tgl_krmlgn.Text = ""
nosj_krmlgn.Text = ""
warna_krmlgn.Text = ""
rol_krmlgn.Text = ""
kg_krmlgn.Text = ""
End Sub

Private Sub kembali_so_Click()
Unload Me
End Sub

'TAMBAH DATA KE SO
Private Sub simpan_so_Click()
koneksi
selek = "SELECT * FROM tbl_SO"
Set so = New ADODB.Recordset
    so.Open selek, dbkoneksi, adOpenDynamic, adLockOptimistic
With so
    .AddNew
        !no_so = no_so
        !tgl_so_krmclpn = tgl_krmclpn
        !no_sj_so_krmclpn = nosj_krmclpn
        !untuk_so_krmclpn = untuk_krmclpn
        !jenis_so_krmclpn = jenis_krmclpn
        !rol_so_krmclpn = rol_krmclpn
        !kg_so_krmclpn = kg_krmclpn
        !tgl_so_hslclpn = tgl_hslclpn
        !no_sj_so_hslclpn = nosj_hslclpn
        !warna_so_hslclpn = warna_hslclpn
        !rol_so_hslclpn = rol_hslclpn
        !kg_so_hslclpn = kg_hslclpn
        !tgl_so_langganan = tgl_krmlgn
        !no_sj_so_langganan = nosj_krmlgn
        !warna_so_langganan = warna_krmlgn
        !rol_so_langganan = rol_krmlgn
        !kg_so_langganan = kg_krmlgn
    .Update
End With

no_so.Text = ""
tgl_krmclpn.Text = ""
nosj_krmclpn.Text = ""
untuk_krmclpn.Text = ""
jenis_krmclpn.Text = ""
rol_krmclpn.Text = ""
kg_krmclpn.Text = ""
tgl_hslclpn.Text = ""
nosj_hslclpn.Text = ""
warna_hslclpn.Text = ""
rol_hslclpn.Text = ""
kg_hslclpn.Text = ""
tgl_krmlgn.Text = ""
nosj_krmlgn.Text = ""
warna_krmlgn.Text = ""
rol_krmlgn.Text = ""
kg_krmlgn.Text = ""
simpan_so.Enabled = False
batal_so.Enabled = False
input_so.Enabled = True
MsgBox ("Data Berhasil di Simpan..."), vbInformation, "Success"
End Sub
