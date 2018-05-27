VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "splash"
   ClientHeight    =   3495
   ClientLeft      =   -60
   ClientTop       =   -120
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5400
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   1980
      Left            =   4080
      Picture         =   "splash.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembayaran Listrik Pasca bayar"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembayaran Listrik Pasca bayar"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PLPB"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   120
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label3.Caption = "Loading " & ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Me.Visible = False
frmLogin.Show
End If
End Sub
