VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPembayaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembayaran"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker tgl 
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Format          =   103350273
      CurrentDate     =   43169
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPembayaran.frx":0000
      Height          =   2655
      Left            =   7800
      TabIndex        =   16
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
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
      Caption         =   "Data"
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
      Height          =   495
      Left            =   12960
      Top             =   1560
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   $"frmPembayaran.frx":0015
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
   Begin MSAdodcLib.Adodc Ado 
      Height          =   495
      Left            =   12600
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtBiayaadmin 
      Height          =   285
      Left            =   2205
      TabIndex        =   7
      Top             =   1785
      Width           =   3375
   End
   Begin VB.TextBox txtBulanbayar 
      Height          =   285
      Left            =   2205
      TabIndex        =   5
      Top             =   1410
      Width           =   3375
   End
   Begin VB.TextBox txtIDBayar 
      Height          =   285
      Left            =   2205
      TabIndex        =   1
      Top             =   270
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   2895
      Left            =   720
      TabIndex        =   13
      Top             =   3360
      Width           =   6615
      _ExtentX        =   11668
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
      Caption         =   "Data Pembayaran"
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
      Bindings        =   "frmPembayaran.frx":00ED
      DataField       =   "IDPelanggan"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2280
      TabIndex        =   15
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "IDPelanggan"
      Text            =   "Pilih pelanggan"
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      DataField       =   "IDPelanggan"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   13320
      TabIndex        =   23
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      DataField       =   "Bulan"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   13200
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Label3"
      DataField       =   "Tarif"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   13200
      TabIndex        =   21
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "Jumlahmeter"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   13200
      TabIndex        =   20
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bayar Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Cari Pelanggan"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Biayaadmin:"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   1830
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Bulanbayar:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tanggal:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IDPelanggan:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   690
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IDBayar:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   315
      Width           =   1815
   End
End
Attribute VB_Name = "frmPembayaran"
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
If txtIDBayar = "" Then
    MsgBox "Kode ID Bayar kosong"
    Else
        Call cari
        If Not rstarif.EOF Then
            If MsgBox("Yakin akan dihapus", vbYesNo, "Perhatian") = vbYes Then
            Delete = "delete * from pembayaran where idbayar = '" & txtIDBayar & "'"
            CONN.Execute Delete
            Call kosong
            Form_Activate
            End If
        End If
End If
End Sub

Private Sub btnSave_Click()
If txtIDBayar = "" Then
MsgBox "Data belum lengkap"
Else
Call cari
If rstarif.EOF Then
    simpan = "insert into pembayaran values('" & txtIDBayar & "','" & txtID.Text & "','" & tgl.Value & "','" & txtBulanbayar & "','" & txtBiayaadmin & "','" & Label2.Caption & "')"
    CONN.Execute simpan
    Call kosong
    Form_Activate
    Else
    edit = "update Penggunaan set IDPelanggan = '" & txtID.Text & "', tanggal = '" & tgl.Value & "', bulanbayar ='" & txtBulanbayar & "', biayaadmin='" & txtBiayaadmin & "', jumlah = '" & Label2.Caption & "' where idbayar ='" & txtIDBayar & "' "
    CONN.Execute edit
    Call kosong
    Form_Activate
End If
End If
End Sub

Private Sub DataGrid1_Click()
txtBulanbayar.Text = Label6.Caption
txtID = Label7.Caption
Label2.Caption = Label3.Caption * Label5.Caption
End Sub

Private Sub Form_Activate()
Call koneksi
Ado.ConnectionString = LokasiDB
Ado.RecordSource = "pembayaran"
Ado.Refresh
Set DG.DataSource = Ado
DG.Refresh
End Sub

Sub cari()
Call koneksi
rstarif.Open "select * from pembayaran where idbayar = '" & txtIDBayar & "'", CONN
rstarif.Requery
End Sub

Sub kosong()
txtIDBayar = ""
txtBulanbayar = ""
txtBiayaadmin = ""
txtID = "Pilih pelanggan"
End Sub

Sub databaru()
txtIDBayar.SetFocus
txtIDBayar = ""
txtBulanbayar = ""
txtBiayaadmin = ""
txtID = "Pilih pelanggan"
End Sub

Sub ketemu()
txtIDBayar = rspelanggan!idbayar
txtID = rspelanggan!idpelanggan
tgl = rspelanggan!tanggal
txtBulanbayar = rspelanggan!bulanbayar
txtBiayaadmin = rspelanggan!biayaadmin

End Sub

Private Sub Form_Load()
Call kosong
End Sub



Private Sub DG_Click()
txtIDBayar = DG.Columns(0).Value
txtID = DG.Columns(1).Value
tgl = DG.Columns(2).Value
txtBulanbayar = DG.Columns(3).Value
txtBiayaadmin = DG.Columns(4).Value
End Sub

Private Sub Text1_Change()
Call koneksi
rspelanggan.Open "select * from pembayaran where idbayar like '%" & Text1 & "%'", CONN
rspelanggan.Requery
If Not rspelanggan.EOF Then
    Ado.ConnectionString = LokasiDB
    Ado.RecordSource = "select * from pembayaran where idbayar like  '%" & Text1 & "%'"
    Ado.Refresh
    Set DG.DataSource = Ado
    DG.Refresh
    Else
        MsgBox "Data tidak ditemukan"
End If
End Sub





Private Sub txtIDBayar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtKode = "" Then
        MsgBox "Kode Pelanggan tidak boleh kosong"
        txtIDBayar.SetFocus
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
