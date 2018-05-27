Attribute VB_Name = "Module1"

Public CONN As ADODB.Connection
Public RSUser As ADODB.Recordset
Public rstarif As ADODB.Recordset
Public rspelanggan As ADODB.Recordset
Public rspenggunaan As ADODB.Recordset
Public rstagihan As ADODB.Recordset
Public RSPembayaran As ADODB.Recordset
Public LokasiDB As String

Public Sub koneksi()
Set CONN = New ADODB.Connection
Set RSUser = New ADODB.Recordset
Set rstarif = New ADODB.Recordset
Set rspelanggan = New ADODB.Recordset
Set RSRenggunaan = New ADODB.Recordset
Set rstagihan = New ADODB.Recordset
Set RSPembayaran = New ADODB.Recordset
LokasiDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PL_Pasca_Bayar.mdb;"
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PL_Pasca_Bayar.mdb;"
End Sub


