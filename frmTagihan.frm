VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTagihan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tagihan"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Ado 
      Height          =   615
      Left            =   13080
      Top             =   2760
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
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTagihan.frx":0000
      Height          =   2055
      Left            =   7440
      TabIndex        =   19
      Top             =   3840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3625
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
      Caption         =   "Data Pengguna"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12840
      Top             =   4320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=PL_Pasca_Bayar.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=PL_Pasca_Bayar.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Penggunaan"
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
   Begin VB.ComboBox txtBulan 
      DataField       =   "Bulan"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmTagihan.frx":0015
      Left            =   1800
      List            =   "frmTagihan.frx":003D
      TabIndex        =   15
      Text            =   "Pilih bulan"
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox txtTahun 
      DataField       =   "Tahun"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmTagihan.frx":00A4
      Left            =   1800
      List            =   "frmTagihan.frx":00C0
      TabIndex        =   14
      Text            =   "Pilih tahun"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3360
      Width           =   3855
   End
   Begin VB.TextBox txtStatus 
      Height          =   285
      Left            =   1635
      TabIndex        =   6
      Top             =   2055
      Width           =   3375
   End
   Begin VB.TextBox txtJumlahmeter 
      Height          =   285
      Left            =   1635
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   2895
      Left            =   480
      TabIndex        =   12
      Top             =   3840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
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
      Caption         =   "Data Tagihan"
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
            LCID            =   2057
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
            LCID            =   2057
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
   Begin MSDataListLib.DataCombo txtID 
      Bindings        =   "frmTagihan.frx":00F4
      DataField       =   "IDPelanggan"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1800
      TabIndex        =   16
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "IDPelanggan"
      Text            =   "Pilih pelanggan"
   End
   Begin VB.Label tahun 
      DataField       =   "Tahun"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11880
      TabIndex        =   21
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label bulan 
      DataField       =   "Bulan"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11640
      TabIndex        =   20
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label akhir 
      DataField       =   "Meterakhir"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11520
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label awal 
      DataField       =   "Meterawal"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   11520
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Cari Status"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   255
      Index           =   4
      Left            =   -210
      TabIndex        =   5
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Jumlahmeter:"
      Height          =   255
      Index           =   3
      Left            =   -210
      TabIndex        =   3
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tahun:"
      Height          =   255
      Index           =   2
      Left            =   -210
      TabIndex        =   2
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Bulan:"
      Height          =   255
      Index           =   1
      Left            =   -210
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IDPelanggan:"
      Height          =   255
      Index           =   0
      Left            =   -210
      TabIndex        =   0
      Top             =   585
      Width           =   1815
   End
End
Attribute VB_Name = "frmTagihan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Change()
txtJumlahmeter = akhir - awal
txtBulan.Text = bulan.Caption
txtTahun.Text = tahun.Caption
End Sub

Private Sub DataGrid1_Click()
txtJumlahmeter = akhir - awal
End Sub

Private Sub btnCancel_Click()
kosong
End Sub

Private Sub btnClose_Click()
Me.Hide
End Sub

Private Sub btnDelete_Click()
If txtID = "Pilih pelanggan" Then
    MsgBox "Kode Pelanggan kosong"
    Else
        Call cari
        If Not rstarif.EOF Then
            If MsgBox("Yakin akan dihapus", vbYesNo, "Perhatian") = vbYes Then
            Delete = "delete * from Tagihan where IDPelanggan = '" & txtID & "'"
            CONN.Execute Delete
            Call kosong
            Form_Activate
            End If
        End If
End If
End Sub

Private Sub btnSave_Click()
If txtID = "Pilih pelanggan" Or txtBulan = "Pilh bulan" Or txtTahun = "pilih tahun" Or txtJumlahmeter = "" Or txtStatus = "" Then
MsgBox "Data belum lengkap"
Else
Call cari
If rstarif.EOF Then
    simpan = "insert into Tagihan values('" & txtID & "','" & txtBulan.Text & "','" & txtTahun.Text & "','" & txtJumlahmeter & "','" & txtStatus & "')"
    CONN.Execute simpan
    Call kosong
    Form_Activate
    Else
    edit = "update tagihan set Bulan = '" & txtBulan.Text & "', tahun = '" & txtTahun.Text & "', jumlahmeter ='" & txtJumlahmeter & "', status='" & txtStatus & "' where IDPelanggan ='" & txtID & "' "
    CONN.Execute edit
    Call kosong
    Form_Activate
End If
End If
End Sub

Private Sub Form_Activate()
Call koneksi
Ado.ConnectionString = LokasiDB
Ado.RecordSource = "Tagihan"
Ado.Refresh
Set DG.DataSource = Ado
DG.Refresh
End Sub

Sub cari()
Call koneksi
rstarif.Open "select * from Tagihan where IDPelanggan = '" & txtID & "'", CONN
rstarif.Requery
End Sub

Sub kosong()
txtIDPelanggan = ""
'txtBulan = "pilih bulan"
'txtTahun = "pilih tahun"
'txtJumlahmeter = ""
txtStatus = ""
End Sub

Sub databaru()
txtIDPelanggan.SetFocus
txtIDPelanggan = ""
'txtBulan = "pilih bulan"
'txtTahun = "pilih tahun"
txtJumlahmeter = ""
txtStatus = ""
End Sub

Sub ketemu()
txtID = rspelanggan!idpelanggan
txtBulan = rspelanggan!bulan
txtTahun = rspelanggan!tahun
txtJumlahmeter = rspelanggan!meterawal
txtStatus = rspelanggan!meterakhir

End Sub

Private Sub Form_Load()
Call kosong
End Sub



Private Sub DG_Click()
txtID = DG.Columns(0).Value
txtBulan = DG.Columns(1).Value
txtTahun = DG.Columns(2).Value
txtJumlahmeter = DG.Columns(3).Value
txtStatus = DG.Columns(4).Value
End Sub

Private Sub Text1_Change()
Call koneksi
rspelanggan.Open "select * from tagihan where status like '%" & Text1 & "%'", CONN
rspelanggan.Requery
If Not rspelanggan.EOF Then
    Ado.ConnectionString = LokasiDB
    Ado.RecordSource = "select * from tagihan where status like  '%" & Text1 & "%'"
    Ado.Refresh
    Set DG.DataSource = Ado
    DG.Refresh
    Else
        MsgBox "Data tidak ditemukan"
End If
End Sub





Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtKode = "" Then
        MsgBox "Kode Pelanggan tidak boleh kosong"
        txtID.SetFocus
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


