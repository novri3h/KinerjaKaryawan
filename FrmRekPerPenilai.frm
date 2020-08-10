VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRekPerpenilai 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekomendasi Per-Penilai"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4620
   Begin Crystal.CrystalReport Lap 
      Left            =   1620
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3540
      TabIndex        =   6
      Top             =   840
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   2550
      TabIndex        =   5
      Top             =   840
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   435
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   75
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pegawai"
      Height          =   330
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NIP Pegawai"
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "FrmRekPerpenilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Dbpenilai As Database
    Dim Recpenilai As Recordset
    
Private Sub Combo1_Change()
    Combo1_Click
End Sub

Private Sub Combo1_Click()
    Recpenilai.Seek "=", Combo1.Text
    If Not Recpenilai.NoMatch Then
        Text1.Text = Recpenilai!namapen
        Command1.Enabled = True
        Command2.Enabled = True
    Else
        Text1.Text = ""
        Command1.Enabled = False
        Command2.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Lap.ReportFileName = App.Path & "\daftarrekomendasiperpenilai.rpt"
    Lap.DataFiles(0) = App.Path & "\pdam.mdb"
    Lap.WindowState = crptMaximized
    Lap.Destination = crptToPrinter
    Lap.ReplaceSelectionFormula "{penilai.nippen}='" & Trim(Combo1.Text) & "'"
    Lap.Action = 2
End Sub

Private Sub Command2_Click()
    Lap.ReportFileName = App.Path & "\daftarrekomendasiperpenilai.rpt"
    Lap.DataFiles(0) = App.Path & "\pdam.mdb"
    Lap.WindowState = crptMaximized
    Lap.Destination = crptToWindow
    Lap.ReplaceSelectionFormula "{penilai.nippen}='" & Trim(Combo1.Text) & "'"
    Lap.Action = 2
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_activate()
    Set Dbpenilai = OpenDatabase(App.Path & "\pdam.mdb")
    Set Recpenilai = Dbpenilai.OpenRecordset("penilai")
    Recpenilai.Index = "idxpenilai"
    Text1.Text = ""
    Command1.Enabled = False
    Command2.Enabled = False
    Combo1.Clear
    Do While Not Recpenilai.EOF
        Combo1.AddItem Recpenilai!nippen
        Recpenilai.MoveNext
    Loop

End Sub
