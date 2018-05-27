VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPelanggan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pelanggan"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo txtKodeTarif 
      Bindings        =   "frmPelanggan.frx":0000
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Kodetarif"
      Text            =   "Pilih tarif"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10560
      Top             =   4560
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\UKK\PLPB\PL_Pasca_Bayar.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\UKK\PLPB\PL_Pasca_Bayar.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tarif"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   375
      Left            =   480
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   2655
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAlamat 
      Height          =   285
      Left            =   2010
      TabIndex        =   7
      Top             =   1485
      Width           =   3375
   End
   Begin VB.TextBox txtNama 
      Height          =   285
      Left            =   2010
      TabIndex        =   5
      Top             =   1105
      Width           =   3375
   End
   Begin VB.TextBox txtNoMeter 
      Height          =   285
      Left            =   2010
      TabIndex        =   3
      Top             =   725
      Width           =   1455
   End
   Begin VB.TextBox txtIDPelanggan 
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Top             =   345
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cari Pelanggan"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Kodetarif:"
      Height          =   255
      Index           =   4
      Left            =   165
      TabIndex        =   8
      Top             =   1905
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Alamat:"
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   6
      Top             =   1530
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nama:"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   4
      Top             =   1155
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "NoMeter:"
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IDPelanggan:"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   390
      Width           =   1815
   End
End
Attribute VB_Name = "frmPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
kosong
End Sub

Private Sub btnClose_Click()
Me.Hide
End Sub

Private Sub btnDelete_Click()
If txtIDPelanggan = "" Then
    MsgBox "Kode Pelanggan kosong"
    Else
        Call cari
        If Not rstarif.EOF Then
            If MsgBox("Yakin akan dihapus", vbYesNo, "Perhatian") = vbYes Then
            Delete = "delete * from pelanggan where IDPelanggan = '" & txtIDPelanggan & "'"
            CONN.Execute Delete
            Call kosong
            Form_Activate
            End If
        End If
End If
End Sub

Private Sub btnSave_Click()
If txtIDPelanggan = "" Or txtNoMeter = "" Or txtNama = "" Or txtAlamat = "" Or txtKodeTarif = "" Then
MsgBox "Data belum lengkap"
Else
Call cari
If rstarif.EOF Then
    simpan = "insert into Pelanggan values('" & txtIDPelanggan & "','" & txtNoMeter & "','" & txtNama & "','" & txtAlamat & "','" & txtKodeTarif & "')"
    CONN.Execute simpan
    Call kosong
    Form_Activate
    Else
    edit = "update Pelanggan set nometer = '" & txtNoMeter & "', nama = '" & txtNama & "', alamat ='" & txtAlamat & "', kodetarif='" & txtKodeTarif & "' where IDPelanggan ='" & txtIDPelanggan & "' "
    CONN.Execute edit
    Call kosong
    Form_Activate
End If
End If
End Sub

Private Sub Form_Activate()
Call koneksi
Ado.ConnectionString = LokasiDB
Ado.RecordSource = "Pelanggan"
Ado.Refresh
Set DG.DataSource = Ado
DG.Refresh
End Sub

Sub cari()
Call koneksi
rstarif.Open "select * from Pelanggan where IDPelanggan = '" & txtIDPelanggan & "'", CONN
rstarif.Requery
End Sub

Sub kosong()
txtIDPelanggan = ""
txtNoMeter = ""
txtNama = ""
txtAlamat = ""
txtKodeTarif = ""
End Sub

Sub databaru()
txtIDPelanggan.SetFocus
txtNoMeter = ""
txtNama = ""
txtAlamat = ""
txtKodeTarif = ""
End Sub
Sub ketemu()
txtIDPelanggan = rspelanggan!idpelanggan
txtNoMeter = rspelanggan!nometer
txtNama = rspelanggan!nama
txtAlamat = rspelanggan!alamat
txtKodeTarif = rspelanggan!kodetarif

End Sub

Private Sub Form_Load()
Call kosong
End Sub



Private Sub DG_Click()
txtIDPelanggan = DG.Columns(0).Value
txtNoMeter = DG.Columns(1).Value
txtNama = DG.Columns(2).Value
txtAlamat = DG.Columns(3).Value
txtKodeTarif = DG.Columns(4).Value
End Sub

Private Sub Text1_Change()
Call koneksi
rspelanggan.Open "select * from pelanggan where nama like '%" & Text1 & "%'", CONN
rspelanggan.Requery
If Not rspelanggan.EOF Then
    Ado.ConnectionString = LokasiDB
    Ado.RecordSource = "select * from pelanggan where nama like  '%" & Text1 & "%'"
    Ado.Refresh
    Set DG.DataSource = Ado
    DG.Refresh
    Else
        MsgBox "Data tidak ditemukan"
End If
End Sub



Private Sub txtIDPelanggan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtKode = "" Then
        MsgBox "Kode Pelanggan tidak boleh kosong"
        txtIDPelanggan.SetFocus
        Else
        Call cari
        If Not rspelanggan.EOF Then
            Call ketemu
        Else
            Call databaru
        End If
        
    End If
End If
End Sub
