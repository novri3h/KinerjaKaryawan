VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmAbsensi 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entry Absen Pegawai"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3120
      TabIndex        =   14
      Text            =   "X (30)"
      Top             =   1260
      Width           =   495
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2220
      TabIndex        =   13
      Text            =   "X(2)"
      Top             =   1260
      Width           =   315
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   900
      TabIndex        =   12
      Text            =   "X (30)"
      Top             =   1260
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Text            =   "X (30)"
      Top             =   660
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Text            =   "X (2)"
      Top             =   660
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Text            =   "X (20)"
      Top             =   660
      Width           =   1755
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   420
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
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
      Connect         =   $"FrmAbsensi.frx":0000
      OLEDBString     =   $"FrmAbsensi.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select pegawai.kdpeg, nmpeg,masuk,keterangan from pegawai,absensi where pegawai.kdpeg=absensi.kdpeg"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmAbsensi.frx":0132
      Height          =   2775
      Left            =   60
      TabIndex        =   6
      Top             =   1020
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "kdpeg"
         Caption         =   "kdpeg"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nmpeg"
         Caption         =   "Nama Pegawai"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "masuk"
         Caption         =   "Masuk"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "keterangan"
         Caption         =   "Keterangan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3014.929
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4740
      TabIndex        =   5
      Top             =   3840
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Top             =   3840
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      Top             =   3840
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1860
      TabIndex        =   2
      Top             =   3840
      Width           =   915
   End
   Begin MSComCtl2.DTPicker tgl 
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   0
      CustomFormat    =   "99-99-9999"
      Format          =   22609923
      CurrentDate     =   37913
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   420
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pegawai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   420
      Width           =   1755
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
