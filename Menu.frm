VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm Menu 
   BackColor       =   &H8000000C&
   Caption         =   "Menu"
   ClientHeight    =   3195
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "Menu.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Lap 
      Left            =   1575
      Top             =   1215
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Lap1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnudata 
      Caption         =   "&Data"
      Begin VB.Menu mnupegawai 
         Caption         =   "&Pegawai"
      End
   End
   Begin VB.Menu mnutransaksi 
      Caption         =   "&F1 PKP"
      Begin VB.Menu mnupenilai 
         Caption         =   "&Penilai"
      End
      Begin VB.Menu mnuteguran 
         Caption         =   "&Teguran"
      End
      Begin VB.Menu mnunilai 
         Caption         =   "&Nilai"
      End
   End
   Begin VB.Menu mnulaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnulapdata 
         Caption         =   "&Laporan Data Pribadi Pegawai"
      End
      Begin VB.Menu mnudaftar 
         Caption         =   "Daftar Pegawai"
      End
      Begin VB.Menu mnurekpegawai 
         Caption         =   "Rekomendasi Perpegawai"
      End
      Begin VB.Menu mnurekpenilai 
         Caption         =   "Rekomendasi Perpenilai"
      End
      Begin VB.Menu mnurekap 
         Caption         =   "Rekapitulasi Penilaian Pegawai"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnukeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnudaftar_Click()
    Lap.ReportFileName = App.Path & "\daftarpegawai.rpt"
    Lap.DataFiles(0) = App.Path & "\pdam.mdb"
    Lap.WindowState = crptMaximized
    Lap.Action = 0
End Sub

Private Sub mnukeluar_Click()
    End
End Sub

Private Sub mnulapdata_Click()
    FrmDataPribadi.Show
End Sub

Private Sub mnunilai_Click()
    FrmNilai.Show
End Sub

Private Sub mnupegawai_Click()
    FrmPegawai.Show
End Sub

Private Sub mnupenilai_Click()
    FrmPenilai.Show
End Sub

Private Sub mnurekap_Click()
    Lap1.ReportFileName = App.Path & "\rekap.rpt"
    Lap1.DataFiles(0) = App.Path & "\pdam.mdb"
    Lap1.WindowState = crptMaximized
    Lap1.Action = 0
End Sub

Private Sub mnurekpegawai_Click()
    FrmRekPerpegawai.Show
End Sub

Private Sub mnurekpenilai_Click()
    FrmRekPerpenilai.Show
End Sub

Private Sub mnuteguran_Click()
    FrmTeguran.Show
End Sub
