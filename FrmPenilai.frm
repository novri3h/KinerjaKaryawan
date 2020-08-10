VERSION 5.00
Begin VB.Form FrmPenilai 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entry Penilai"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
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
   ScaleHeight     =   2070
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   1620
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1620
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   1620
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1620
      Width           =   915
   End
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1320
      TabIndex        =   7
      Text            =   "X (30)"
      Top             =   1200
      Width           =   2475
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1320
      TabIndex        =   6
      Text            =   "X (20)"
      Top             =   840
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "X (30)"
      Top             =   480
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "X (20)"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jabatan :"
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
      Left            =   -660
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Golongan :"
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
      Left            =   -660
      TabIndex        =   2
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Penilai :"
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
      Left            =   -660
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIP Penilai :"
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
      Left            =   -660
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmPenilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dbpenilai As Database
Dim Recpenilai As Recordset
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
    If MsgBox("Anda yakin akan menghapus data Penilai " & Text2.Text & " Ini ??", vbQuestion + vbYesNo, "Konfirmasi Hapus") = vbYes Then
        Cari
        Recpenilai.Delete
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
    Set Dbpenilai = OpenDatabase(App.Path & "\pdam.mdb")
    Set Recpenilai = Dbpenilai.OpenRecordset("Penilai")
    Recpenilai.Index = "idxPenilai"
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
    Combo1.Enabled = Ena
    Combo2.Enabled = Ena
    

End Sub

Private Sub Bersih()
    Text2.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
End Sub

Private Sub Cmd(Simpan As Boolean, Edit As Boolean)
    Command1.Enabled = Simpan
    Command2.Enabled = Edit
    Command3.Enabled = Edit
    
End Sub
Private Sub Simpan()
Dim Periksa As Boolean
Dim Pendidikan As String
Periksa = IIf(Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "", False, True)
If Periksa Then
    If Edit Then
        Recpenilai.Seek "=", Text1.Text
        Recpenilai.Edit
    Else
        Recpenilai.AddNew
    End If
   
    Recpenilai!nippen = Trim(Text1.Text)
    Recpenilai!namapen = Trim(Text2.Text)
    Recpenilai!golpen = Trim(Combo1.Text)
    Recpenilai!jabpen = Trim(Combo2.Text)
    
    
    Recpenilai.Update
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
Recpenilai.Seek "=", Text1.Text
If Not Recpenilai.NoMatch Then
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
    Text1.Text = Recpenilai!nippen
    Text2.Text = Recpenilai!namapen
    Combo1.Text = Recpenilai!golpen
    Combo2.Text = Recpenilai!jabpen
    
    
End Sub

Private Sub Text1_Change()
    Cari

End Sub



