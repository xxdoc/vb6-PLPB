VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Akun user"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAccount.frx":0000
      Height          =   2655
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   5415
      _ExtentX        =   9551
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8760
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=PL_Pasca_Bayar.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=PL_Pasca_Bayar.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Users"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4560
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   615
      Left            =   3120
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Role"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmAccount.frx":0015
      Left            =   1200
      List            =   "frmAccount.frx":001F
      TabIndex        =   7
      Text            =   "Pilih Role"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1185
      TabIndex        =   5
      Top             =   1050
      Width           =   2775
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "Username"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1185
      TabIndex        =   3
      Top             =   660
      Width           =   2775
   End
   Begin VB.TextBox txtID 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1185
      TabIndex        =   1
      Top             =   285
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Role:"
      Height          =   255
      Index           =   3
      Left            =   -660
      TabIndex        =   6
      Top             =   1470
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   255
      Index           =   2
      Left            =   -660
      TabIndex        =   4
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   255
      Index           =   1
      Left            =   -660
      TabIndex        =   2
      Top             =   705
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   255
      Index           =   0
      Left            =   -660
      TabIndex        =   0
      Top             =   330
      Width           =   1815
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.UpdateBatch adAffectCurrent
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.CancelBatch adAffectCurrent
End Sub
