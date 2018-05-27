VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dashboard"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDash.frx":0000
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   684
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1920
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":501E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":511A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":52C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":54029
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":55246
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":56D71
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":58100
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":58DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDash.frx":5A077
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "16:05"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/03/2018"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Akun user"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Master tarif"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Master pelanggan"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Penggunaan"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tagihan"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pembayaran"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Laporan"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Laporan tarif"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Laporan pelanggan"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Laporan penggunaan"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Laporan tagihan"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Laporan pembayaran"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informasi"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   10215
         TabIndex        =   2
         Top             =   840
         Width           =   10215
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PLPB SYSTEM - DEVELOP BY RENPRTN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   6120
      TabIndex        =   5
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PERUSAHAAN LISTRIK NEGARA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PERUSAHAAN LISTRIK NEGARA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   4080
      Picture         =   "frmDash.frx":5B4C3
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Menu mFIle 
      Caption         =   "File"
      Begin VB.Menu mAccount 
         Caption         =   "Account"
      End
      Begin VB.Menu mLog 
         Caption         =   "Logout"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mMaster 
      Caption         =   "Master"
      Begin VB.Menu mTarif 
         Caption         =   "Tarif"
      End
      Begin VB.Menu mPelanggan 
         Caption         =   "Pelanggan"
      End
      Begin VB.Menu mPenggunaan 
         Caption         =   "Penggunaan"
      End
   End
   Begin VB.Menu mTrans 
      Caption         =   "Transaksi"
      Begin VB.Menu mTagih 
         Caption         =   "Tagihan"
      End
      Begin VB.Menu mBayar 
         Caption         =   "Pembayaran"
      End
   End
   Begin VB.Menu mLab 
      Caption         =   "Laporan"
      Begin VB.Menu mltarif 
         Caption         =   "Laporan Tarif"
      End
      Begin VB.Menu mlp 
         Caption         =   "Laporan Pelanggan"
      End
      Begin VB.Menu lp 
         Caption         =   "Laporan Penggunaan"
      End
      Begin VB.Menu mlt 
         Caption         =   "Laporan Tagihan"
      End
      Begin VB.Menu mlpp 
         Caption         =   "Laporan Pembayaran"
      End
   End
End
Attribute VB_Name = "frmDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lp_Click()
DRPE.Show 1
End Sub

Private Sub mAccount_Click()
frmAccount.Show 1
End Sub

Private Sub mBayar_Click()
frmPembayaran.Show 1
End Sub

Private Sub mExit_Click()
If MsgBox("Yakin tutup aplikasi ?", vbQuestion + vbYesNo, "PLPB System") = vbYes Then
    End
End If
End Sub

Private Sub mLog_Click()
Me.Visible = False
frmLogin.Show 1

End Sub

Private Sub mlp_Click()
DRP.Show 1
End Sub

Private Sub mlpp_Click()
DRPA.Show 1
End Sub

Private Sub mlt_Click()
DRTA.Show 1
End Sub

Private Sub mltarif_Click()
DRT.Show 1
End Sub

Private Sub mPelanggan_Click()
frmPelanggan.Show 1
End Sub

Private Sub mPenggunaan_Click()
frmPenggunaan.Show 1
End Sub

Private Sub mTagih_Click()
frmTagihan.Show 1
End Sub

Private Sub mTarif_Click()
frmTarif.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
    frmAccount.Show 1
ElseIf Button.Index = 2 Then
    frmTarif.Show 1
ElseIf Button.Index = 3 Then
    frmPelanggan.Show 1
ElseIf Button.Index = 4 Then
    frmPenggunaan.Show 1
ElseIf Button.Index = 5 Then
    frmTagihan.Show 1
ElseIf Button.Index = 6 Then
    frmPembayaran.Show 1
ElseIf Button.Index = 7 Then
ElseIf Button.Index = 8 Then
    frmTip.Show 1
ElseIf Button.Index = 9 Then
    frmAbout.Show 1
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Index = 1 Then
    DRT.Show 1
ElseIf ButtonMenu.Index = 2 Then
    DRP.Show 1
ElseIf ButtonMenu.Index = 3 Then
    DRPE.Show 1
ElseIf ButtonMenu.Index = 4 Then
    DRTA.Show 1
ElseIf ButtonMenu.Index = 5 Then
    DRPA.Show 1
    
End If
End Sub
