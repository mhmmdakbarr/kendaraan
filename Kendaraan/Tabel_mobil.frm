VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F_Mobil 
   Caption         =   "T_mobil"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H80000002&
      Height          =   5175
      Left            =   8280
      TabIndex        =   22
      Top             =   0
      Width           =   6135
      Begin VB.Data Data1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Kendaraan\Kendaraan_A2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabel_mobil"
         Top             =   2040
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000002&
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   14
      Top             =   5160
      Width           =   14415
      Begin VB.CommandButton Ckeluar 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11640
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Chapus 
         Caption         =   "Hapus"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9600
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Cupdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Cedit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Csimpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Ctambah 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tambah"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      Caption         =   "Table Mobil"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   13
      Top             =   6600
      Width           =   14415
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Tabel_mobil.frx":0000
         Height          =   1695
         Left            =   480
         OleObjectBlob   =   "Tabel_mobil.frx":0014
         TabIndex        =   21
         Top             =   360
         Width           =   13575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Data Mobil"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox Tjumlah 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox Ttaper 
         Height          =   405
         Left            =   1800
         TabIndex        =   11
         Top             =   3000
         Width           =   3615
      End
      Begin VB.ComboBox Cjenis 
         Height          =   315
         ItemData        =   "Tabel_mobil.frx":10BF
         Left            =   1800
         List            =   "Tabel_mobil.frx":10C1
         TabIndex        =   10
         Top             =   2280
         Width           =   3615
      End
      Begin VB.ComboBox Cwarna 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Tnama 
         Height          =   405
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Tno 
         Height          =   405
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000002&
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         Caption         =   "Tahun Perakitan"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
         Caption         =   "Jenis Mobil"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         Caption         =   "Warna Mobil"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   "Nama Mobil"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         Caption         =   "Nomor Polisi"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "F_Mobil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kendaraan As Database
Dim tabel_mobil As Recordset
Private Sub Cedit_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan Nomor Polisi mobil yang di cari", "Cari Data")
tabel_mobil.Seek "=", pesan

If tabel_mobil.NoMatch Then
X = MsgBox("Maaf nomor polisi mobil yang di cari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tno.Text = tabel_mobil!No_polisi
Tno.Enabled = False
Tnama.Text = tabel_mobil!nama_mobil
Cwarna.Text = tabel_mobil!Warna_mobil
Cjenis.Text = tabel_mobil!jenis_mobil
Ttaper.Text = tabel_mobil!Tahun_perakitan
Tjumlah.Text = tabel_mobil!jumlah
End Sub

Private Sub Chapus_Click()
Dim pesan As String * 20
pesan = InputBox("Masukkan nomor polisi mobil yang dicari", "Cari Data")
tabel_mobil.Seek "=", pesan

If tabel_mobil.NoMatch Then
X = MsgBox("Maaf nomor polisi mobil yang dicari tidak ada", vbInformation, "informasi")
Exit Sub
End If

Tno.Text = tabel_mobil!No_polisi
Tno.Enabled = False
Tnama.Text = tabel_mobil!nama_mobil
Cwarna.Text = tabel_mobil!Warna_mobil
Cjenis.Text = tabel_mobil!jenis_mobil
Ttaper.Text = tabel_mobil!Tahun_perakitan
Tjumlah.Text = tabel_mobil!jumlah

X = MsgBox("Yakin Data Akan dihapus ?", vbYesNo, "konfirmasi")
    If X = vbYes Then
    tabel_mobil.Seek "=", Tno.Text
    tabel_mobil.Delete
    kosong
    Data1.Refresh
    DBGrid1.Refresh
    End If
End Sub

Private Sub Ckeluar_Click()
End
End Sub

Private Sub Csimpan_Click()
tabel_mobil.Seek "=", Tno.Text

If tabel_mobil.NoMatch Then

tabel_mobil.AddNew
tabel_mobil!No_polisi = Tno.Text
tabel_mobil!nama_mobil = Tnama.Text
tabel_mobil!Warna_mobil = Cwarna.Text
tabel_mobil!jenis_mobil = Cjenis.Text
tabel_mobil!Tahun_perakitan = Ttaper.Text
tabel_mobil!jumlah = Tjumlah.Text
tabel_mobil.Update

X = MsgBox("Data berhasil tersimpan", vbInformation, "pesan")
Data1.Refresh
DBGrid1.Refresh
Else
X = MsgBox("Maaf No polisi mobil ada yang sama", vbInformation, "pesan")
End If
End Sub

Private Sub Ctambah_Click()
Tno.SetFocus
kosong
End Sub
Private Sub kosong()
Tno.Text = ""
Tnama.Text = ""
Cwarna.Text = ""
Cjenis.Text = ""
Ttaper.Text = ""
Tjumlah.Text = ""
End Sub

Private Sub Cupdate_Click()
tabel_mobil.Edit
tabel_mobil!nama_mobil = Tnama.Text
tabel_mobil!Warna_mobil = Cwarna.Text
tabel_mobil!jenis_mobil = Cjenis.Text
tabel_mobil!Tahun_perakitan = Ttaper.Text
tabel_mobil!jumlah = Tjumlah.Text
tabel_mobil.Update

X = MsgBox("Data Berhasil diubah", vbInformation, "informasi")
Data1.Refresh
DBGrid1.Refresh
End Sub

Private Sub Form_Load()
Set penjualan = OpenDatabase("D:\Kendaraan\Kendaraan_A2.mdb")
Set tabel_mobil = penjualan.OpenRecordset("tabel_mobil")
tabel_mobil.Index = "Kunci_mobil"
End Sub



