VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPegawai 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entry Pegawai"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text11 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   3
      Text            =   "X (30)"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6900
      TabIndex        =   21
      Text            =   "9(2)"
      Top             =   4380
      Width           =   1035
   End
   Begin VB.TextBox Text9 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6900
      TabIndex        =   20
      Text            =   "9(2)"
      Top             =   4020
      Width           =   1035
   End
   Begin VB.TextBox Text8 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6900
      TabIndex        =   19
      Text            =   "9(2)"
      Top             =   3660
      Width           =   1035
   End
   Begin VB.TextBox Text7 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6900
      TabIndex        =   18
      Text            =   "9(2)"
      Top             =   3300
      Width           =   1035
   End
   Begin VB.TextBox Text6 
      ForeColor       =   &H00FF0000&
      Height          =   1185
      Left            =   5760
      TabIndex        =   17
      Text            =   "X (200)"
      Top             =   1980
      Width           =   4635
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H0080C0FF&
      Caption         =   "SARJANA"
      Height          =   195
      Left            =   7920
      TabIndex        =   16
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H0080C0FF&
      Caption         =   "SM/D-3"
      Height          =   195
      Left            =   7920
      TabIndex        =   15
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080C0FF&
      Caption         =   "SLTA"
      Height          =   195
      Left            =   7920
      TabIndex        =   14
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SLTP"
      Height          =   195
      Left            =   7920
      TabIndex        =   13
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080C0FF&
      Caption         =   "SD"
      Height          =   195
      Left            =   7920
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      ForeColor       =   &H00FF0000&
      Height          =   705
      Left            =   2220
      TabIndex        =   2
      Text            =   "X (30)"
      Top             =   720
      Width           =   3435
   End
   Begin VB.TextBox Text4 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   11
      Text            =   "9(1)"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "TIDAK"
      Height          =   375
      Left            =   2220
      TabIndex        =   10
      Top             =   3900
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "YA"
      Height          =   495
      Left            =   2220
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   6
      Text            =   "X (30)"
      Top             =   2460
      Width           =   3435
   End
   Begin MSComCtl2.DTPicker tgl 
      Height          =   315
      Left            =   2220
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   92274691
      CurrentDate     =   37929
   End
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2220
      TabIndex        =   5
      Text            =   "X(30)"
      Top             =   2100
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2220
      TabIndex        =   4
      Text            =   "X(20)"
      Top             =   1740
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   9720
      TabIndex        =   27
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   8640
      TabIndex        =   26
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   7620
      TabIndex        =   25
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6660
      TabIndex        =   24
      Top             =   4800
      Width           =   915
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   1
      Text            =   "X (30)"
      Top             =   420
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "X (20)"
      Top             =   120
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker tgl1 
      Height          =   315
      Left            =   2220
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   92078083
      CurrentDate     =   37929
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telp:"
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
      Left            =   360
      TabIndex        =   42
      Top             =   1500
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanpa Ket :"
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
      Left            =   5760
      TabIndex        =   41
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Terlambat :"
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
      Left            =   5760
      TabIndex        =   40
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Izin :"
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
      Left            =   5760
      TabIndex        =   39
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sakit :"
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
      Left            =   5760
      TabIndex        =   38
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pengalaman dan Penempatan :"
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
      Left            =   5220
      TabIndex        =   37
      Top             =   1740
      Width           =   3135
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pendidikan Diakui : "
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
      Left            =   5820
      TabIndex        =   36
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Anak : "
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
      Left            =   240
      TabIndex        =   35
      Top             =   4380
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kawin : "
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
      Left            =   240
      TabIndex        =   34
      Top             =   3540
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Pengangkatan : "
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
      Left            =   60
      TabIndex        =   33
      Top             =   3180
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Lahir : "
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
      Left            =   300
      TabIndex        =   32
      Top             =   2820
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tempat Lahir : "
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
      Left            =   300
      TabIndex        =   31
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jabatan : "
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
      Left            =   300
      TabIndex        =   30
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Golongan : "
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
      Left            =   300
      TabIndex        =   29
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat : "
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
      Left            =   240
      TabIndex        =   28
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pegawai :"
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
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIP :"
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
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DbPegawai As Database
Dim RecPegawai As Recordset
Dim Edit As Boolean
Const Putih = &H8000000E
Const Hitam = &H8000000F

Private Sub Command1_Click()
    Simpan
End Sub

Private Sub Command2_Click()
    Text1.Enabled = False
    Edit = True
    Cmd True, False
    Kotor True
End Sub

Private Sub Command3_Click()
    If MsgBox("Anda yakin akan menghapus data Pegawai " & Text2.Text & " Ini ??", vbQuestion + vbYesNo, "Konfirmasi Hapus") = vbYes Then
        Cari
        RecPegawai.Delete
        Text1.Text = ""
    End If
End Sub

Private Sub Command4_Click()

    Unload Me
End Sub

Private Sub Form_activate()
    Cari
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{tab}"
            KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Set DbPegawai = OpenDatabase(App.Path & "\pdam.mdb")
    Set RecPegawai = DbPegawai.OpenRecordset("Pegawai")
    RecPegawai.Index = "idxPegawai"
    Text1.Text = ""
    Combo1.Clear
    Combo1.AddItem "IV A"
    Combo1.AddItem "IV B"
    Combo1.AddItem "IV C"
    Combo1.AddItem "IV D"
    Combo1.AddItem "III A"
    Combo1.AddItem "III B"
    Combo1.AddItem "III C"
    Combo1.AddItem "III D"
    Combo1.AddItem "II A"
    Combo1.AddItem "II B"
    Combo1.AddItem "II C"
    Combo1.AddItem "II D"
    Combo1.AddItem "I A"
    Combo1.AddItem "I B"
    Combo1.AddItem "I C"
    Combo1.AddItem "I D"
    Combo2.AddItem "Kabag"
    Combo2.AddItem "Kasubag"
    Combo2.AddItem "Kasi"
    tgl.Value = Date
    tgl.Value = Date
    
End Sub
Private Sub Kotor(Ena As Boolean)
Dim Warna
Dim Kut
    If Ena Then
        Warna = Putih
        Kut = True
    Else
        Warna = Hitam
        kud = False
    End If
    
    Text2.Enabled = Ena
    Text3.Enabled = Ena
    Text4.Enabled = Ena
    Text5.Enabled = Ena
    Text11.Enabled = Ena
    
    Text6.Enabled = Ena
    Text7.Enabled = Ena
    Text8.Enabled = Ena
    Text9.Enabled = Ena
    Text10.Enabled = Ena
    tgl.Enabled = Ena
    tgl1.Enabled = Ena
    Combo1.Enabled = Ena
    Combo2.Enabled = Ena
    Option1.Enabled = Ena
    Option2.Enabled = Ena
    Check1.Enabled = Kut
    Check2.Enabled = Kut
    Check3.Enabled = Kut
    Check4.Enabled = Kut
    Check5.Enabled = Kut
    Combo1.BackColor = Warna
    Combo2.BackColor = Warna
    Text2.BackColor = Warna
    Text3.BackColor = Warna
    Text4.BackColor = Warna
    Text5.BackColor = Warna
    Text11.BackColor = Warna
    Text6.BackColor = Warna
    Text7.BackColor = Warna
    Text8.BackColor = Warna
    Text9.BackColor = Warna
    Text10.BackColor = Warna
    

End Sub

Private Sub Bersih()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text11.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
    tgl.Value = Date
    tgl1.Value = Date
    Option2.Value = True
    Check1.Value = vbUnchecked
    Check2.Value = vbUnchecked
    Check3.Value = vbUnchecked
    Check4.Value = vbUnchecked
    Check5.Value = vbUnchecked
End Sub

Private Sub Cmd(Simpan As Boolean, Edit As Boolean)
    Command1.Enabled = Simpan
    Command2.Enabled = Edit
    Command3.Enabled = Edit
    
End Sub
Private Sub Simpan()
Dim Periksa As Boolean
Dim Pendidikan As String
Periksa = IIf(Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text11.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "", False, True)
If Periksa Then
    If Edit Then
        RecPegawai.Seek "=", Text1.Text
        RecPegawai.Edit
    Else
        RecPegawai.AddNew
    End If
   
    RecPegawai!nip_peg = Trim(Text1.Text)
    RecPegawai!namapeg = Trim(Text2.Text)
    RecPegawai!golpeg = Trim(Combo1.Text)
    RecPegawai!jabpeg = Trim(Combo2.Text)
    RecPegawai!tempatlahir = Trim(Text3.Text)
    RecPegawai!tanggallahir = tgl.Value
    RecPegawai!tglangkat = tgl1.Value
    RecPegawai!kawin = IIf(Option1.Value = True, "K", "B")
    RecPegawai!anak = Val(Text4.Text)
    RecPegawai!alamatrumah = Trim(Text5.Text)
    RecPegawai!telp = Trim(Text11.Text)
    Pendidikan = IIf(Check1.Value = vbChecked, "0", "1")
    Pendidikan = IIf(Check2.Value = vbChecked, Pendidikan & "0", Pendidikan & "1")
    Pendidikan = IIf(Check3.Value = vbChecked, Pendidikan & "0", Pendidikan & "1")
    Pendidikan = IIf(Check4.Value = vbChecked, Pendidikan & "0", Pendidikan & "1")
    Pendidikan = IIf(Check5.Value = vbChecked, Pendidikan & "0", Pendidikan & "1")
    RecPegawai!penddiakui = Pendidikan
    RecPegawai!pengpenpeg = Trim(Text6.Text)
    RecPegawai!sakit = Val(Text7.Text)
    RecPegawai!izin = Val(Text8.Text)
    RecPegawai!terlambat = Val(Text9.Text)
    RecPegawai!tanpaket = Val(Text10.Text)
    
    RecPegawai.Update
    Kotor False
    Cmd False, True
    Edit = False
    Text1.Enabled = True
    Text1.SetFocus
Else
    MsgBox "Data belum lengkap, tolong di isi yang masih kosong", vbCritical, "Erorr.."
End If
End Sub

Private Sub Cari()
RecPegawai.Seek "=", Text1.Text
If Not RecPegawai.NoMatch Then
    Tampil
    Cmd False, True
    Kotor False
Else
    Bersih
    Cmd True, False
    Kotor True
End If
End Sub
Private Sub Tampil()
    Text1.Text = RecPegawai!nip_peg
    Text2.Text = RecPegawai!namapeg
    Combo1.Text = RecPegawai!golpeg
    Combo2.Text = RecPegawai!jabpeg
    Text3.Text = RecPegawai!tempatlahir
    tgl.Value = RecPegawai!tanggallahir
    tgl1.Value = RecPegawai!tglangkat
    Option1.Value = IIf(RecPegawai!kawin = "B", True, False)
    Option2.Value = IIf(RecPegawai!kawin = "B", False, True)
    Text4.Text = RecPegawai!anak
    Text5.Text = RecPegawai!alamatrumah
    Text11.Text = RecPegawai!telp
    Text6.Text = RecPegawai!pengpenpeg
    Check1.Value = IIf(Mid(RecPegawai!penddiakui, 1, 1) = "0", vbChecked, vbUnchecked)
    Check2.Value = IIf(Mid(RecPegawai!penddiakui, 2, 1) = "0", vbChecked, vbUnchecked)
    Check3.Value = IIf(Mid(RecPegawai!penddiakui, 3, 1) = "0", vbChecked, vbUnchecked)
    Check4.Value = IIf(Mid(RecPegawai!penddiakui, 4, 1) = "0", vbChecked, vbUnchecked)
    Check5.Value = IIf(Mid(RecPegawai!penddiakui, 5, 1) = "0", vbChecked, vbUnchecked)
    Text7.Text = RecPegawai!sakit
    Text8.Text = RecPegawai!izin
    Text9.Text = RecPegawai!terlambat
    Text10.Text = RecPegawai!tanpaket
    
End Sub

Private Sub Text1_Change()
    Cari

End Sub


