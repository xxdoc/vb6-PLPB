VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTarif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Tarif Listrik"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   2895
      Left            =   240
      TabIndex        =   6
      Top             =   3120
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
      Caption         =   "Data Tarif"
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
   Begin MSAdodcLib.Adodc Ado 
      Height          =   375
      Left            =   1560
      Top             =   6960
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.TextBox txtTarif 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtDaya 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtKode 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Cari Daya"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Tarif"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Tarif / KWH"
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Daya"
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1500
   End
End
Attribute VB_Name = "frmTarif"
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
If txtKode = "" Then
    MsgBox "Kode Barang kosong"
    Else
        Call cari
        If Not rstarif.EOF Then
            If MsgBox("Yakin akan dihapus", vbYesNo, "Perhatian") = vbYes Then
            Delete = "delete * from tarif where KodeTarif = '" & txtKode & "'"
            CONN.Execute Delete
            Call kosong
            Form_Activate
            End If
        End If
End If
End Sub

Private Sub btnSave_Click()
If txtKode = "" Or txtDaya = "" Or txtTarif = "" Then
MsgBox "Data belum lengkap"
Else
Call cari
If rstarif.EOF Then
    simpan = "insert into tarif values('" & txtKode & "','" & txtDaya & "','" & txtTarif & "')"
    CONN.Execute simpan
    Call kosong
    Form_Activate
    Else
    edit = "update tarif set daya = '" & txtDaya & "', tarif = '" & txtTarif & "' where kodetarif ='" & txtKode & "' "
    CONN.Execute edit
    Call kosong
    Form_Activate
End If
End If
End Sub

Private Sub Form_Activate()
Call koneksi
Ado.ConnectionString = LokasiDB
Ado.RecordSource = "tarif"
Ado.Refresh
Set DG.DataSource = Ado
DG.Refresh
End Sub

Sub cari()
Call koneksi
rstarif.Open "select * from tarif where KodeTarif = '" & txtKode & "'", CONN
rstarif.Requery
End Sub

Sub kosong()
txtKode = ""
txtDaya = ""
txtTarif = ""
End Sub

Sub databaru()
txtKode.SetFocus
txtDaya = ""
txtTarif = ""
End Sub
Sub ketemu()
txtKode = rstarif!kodetarif
txtDaya = rstarif!Daya
txtTarif = rstarif!Tarif
txtKode.SetFocus
End Sub

Private Sub Form_Load()
Call kosong
End Sub

Private Sub txtDaya_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTarif.SetFocus
End If
End Sub

Private Sub DG_Click()
txtKode = DG.Columns(0).Value
txtDaya = DG.Columns(1).Value
txtTarif = DG.Columns(2).Value
End Sub

Private Sub Text1_Change()
Call koneksi
rstarif.Open "select * from tarif where Daya like '%" & Text1 & "%'", CONN
rstarif.Requery
If Not rstarif.EOF Then
    Ado.ConnectionString = LokasiDB
    Ado.RecordSource = "select * from tarif where Daya like  '%" & Text1 & "%'"
    Ado.Refresh
    Set DG.DataSource = Ado
    DG.Refresh
    Else
        MsgBox "Data tidak ditemukan"
End If
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtKode = "" Then
        MsgBox "Kode Tarif tidak boleh kosong"
        txtKode.SetFocus
        Else
        Call cari
        If Not rstarif.EOF Then
            Call ketemu
        Else
            Call databaru
        End If
        
    End If
End If
End Sub
