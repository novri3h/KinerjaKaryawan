VERSION 5.00
Begin VB.Form FrmTeguran 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entry Teguran"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      ForeColor       =   &H00FF0000&
      Height          =   1605
      Left            =   2220
      TabIndex        =   4
      Text            =   "X(200)"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   3
      Text            =   "9(2)"
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2220
      TabIndex        =   2
      Text            =   "X(3)"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4740
      TabIndex        =   8
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3780
      TabIndex        =   7
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2820
      TabIndex        =   6
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1860
      TabIndex        =   5
      Top             =   3300
      Width           =   915
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   1
      Text            =   "X (30)"
      Top             =   480
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   0
      Text            =   "X (20)"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan : "
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
      TabIndex        =   13
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah : "
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
      TabIndex        =   12
      Top             =   1260
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Teguran : "
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
      TabIndex        =   11
      Top             =   900
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmTeguran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DbTeguran As Database
Dim RecTeguran As Recordset
Dim DbPegawai As Database
Dim RecPegawai As Recordset
Dim Edit As Boolean
Const Putih = &H8000000E
Const Hitam = &H8000000F

Private Sub Combo1_Change()
    Combo1_Click
End Sub

Private Sub Combo1_Click()
    Cari
End Sub

Private Sub Command1_Click()
    Simpan
End Sub

Private Sub Command2_Click()
    Text1.Enabled = False
    
    Edit = True
    Cmd True, False
    Kotor True
    Combo1.Enabled = False
End Sub

Private Sub Command3_Click()
    If MsgBox("Anda yakin akan menghapus data Teguran " & Text2.Text & " Ini ??", vbQuestion + vbYesNo, "Konfirmasi Hapus") = vbYes Then
        Cari
        RecTeguran.Delete
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
    Set DbTeguran = OpenDatabase(App.Path & "\pdam.mdb")
    Set RecTeguran = DbTeguran.OpenRecordset("Teguran")
    RecTeguran.Index = "idxTeguran"
    Set DbPegawai = OpenDatabase(App.Path & "\pdam.mdb")
    Set RecPegawai = DbPegawai.OpenRecordset("pegawai")
    RecPegawai.Index = "idxpegawai"
    Text1.Text = ""
    Combo1.Clear
    Combo1.AddItem "Lisan"
    Combo1.AddItem "Tertulis"
    Combo1.AddItem "Peringatan Direksi"
       
    
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
    
    
    Combo1.Enabled = Ena
    Text4.Enabled = Ena
    Text3.Enabled = Ena
    

End Sub

Private Sub Bersih()
    'Text2.Text = ""
    'Combo1.Text = ""
    Text4.Text = ""
    Text3.Text = ""
End Sub

Private Sub Cmd(Simpan As Boolean, Edit As Boolean)
    Command1.Enabled = Simpan
    Command2.Enabled = Edit
    Command3.Enabled = Edit
    
End Sub
Private Sub Simpan()
Dim Periksa As Boolean
Dim Pendidikan As String
Periksa = IIf(Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Text3.Text = "" Or Text4.Text = "", False, True)
If Periksa Then
    If Edit Then
        RecTeguran.Seek "=", Text1.Text, Mid(Combo1.Text, 1, 1)
        RecTeguran.Edit
    Else
        RecTeguran.AddNew
    End If
   
    RecTeguran!nip_peg = Trim(Text1.Text)
    RecTeguran!jenisteg = Mid(Combo1.Text, 1, 1)
    RecTeguran!jumlah = Val(Text3.Text)
    RecTeguran!sebab = Trim(Text4.Text)
    
    
    RecTeguran.Update
    Kotor False
    Cmd False, True
    Edit = False
    Text1.Enabled = True
    Combo1.Enabled = True
    Text1.SetFocus
Else
    MsgBox "Data belum lengkap, tolong di isi yang masih kosong", vbCritical, "Erorr.."
End If
End Sub

Private Sub Cari()
Carinip
RecTeguran.Seek "=", Text1.Text, Mid(Combo1.Text, 1, 1)
If Not RecTeguran.NoMatch Then
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
    
    Text3.Text = RecTeguran!jumlah
    Text4.Text = RecTeguran!sebab
    
    
End Sub
Private Sub Carinip()
    RecPegawai.Seek "=", Text1.Text
    If Not RecPegawai.NoMatch Then
        Text2.Text = RecPegawai!namapeg
    Else
        Text2.Text = ""
    End If
End Sub
Private Sub Text1_Change()
    Cari

End Sub




