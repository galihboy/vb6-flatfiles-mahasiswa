VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APLIKASI PENGELOLAAN DATA MAHASISWA. IF. UNIKOM. 2013"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "Ubah"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtCari 
      Height          =   285
      Left            =   3720
      TabIndex        =   14
      Top             =   3315
      Width           =   1815
   End
   Begin VB.ComboBox cboCari 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3315
      Width           =   975
   End
   Begin VB.ListBox lstDaftar 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
      Begin VB.OptionButton optJenisKelamin 
         Caption         =   "Perempuan"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optJenisKelamin 
         Caption         =   "Laki-laki"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txtNama 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "-"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtNIM 
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "-"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Developed by Galih Hermawan (IF UNIKOM). Februari 2013."
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   4815
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   1920
      X2              =   1920
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Label Label7 
      Caption         =   "Cari"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   1800
      X2              =   1800
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Label Label6 
      Caption         =   "Daftar mahasiswa."
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label5 
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Nama"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "NIM"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PROGRAM STUDI TEKNIK INFORMATIKA. UNIKOM."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA MAHASISWA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const strNamaFile = "dataMahasiswa.txt"

Dim strAlamatFile As String
Dim bTambah As Boolean, bEdit As Boolean
Dim strDiubah As String

Private Sub cmdBatal_Click()
    bTambah = False
    bEdit = False
    txtNIM.Locked = False
    KunciTextBox True
    TombolNormal True
    lstDaftar.Enabled = True
End Sub

Private Sub cmdCari_Click()
    Dim strCari As String, arrDataHasilCari() As String
    Dim i As Integer, boolKetemu As Boolean

    strCari = LCase(cboCari.Text)
    
    If Trim(txtCari) = vbNullString Then
        MasukkanIsiFileKeListBox strAlamatFile
    Else
        CariData strAlamatFile, strCari, Trim(txtCari), boolKetemu, arrDataHasilCari
        
        If boolKetemu Then
            lstDaftar.Clear
            For i = 0 To UBound(arrDataHasilCari) - 1
                lstDaftar.AddItem arrDataHasilCari(i)
            Next i
        Else
            MsgBox "Data tidak ditemukan."
        End If
    End If
End Sub

Private Sub cmdHapus_Click()
    Dim strData As String, strKelamin As String, strKelaminPanjang As String
    Dim tanya As Integer
    
    If optJenisKelamin(0).Value = True Then
        strKelamin = "l"
        strKelaminPanjang = "Laki-laki"
    Else
        strKelamin = "p"
        strKelaminPanjang = "Perempuan"
    End If
    
    tanya = MsgBox("Anda yakin akan menghapus data berikut?" & vbCrLf & vbCrLf & _
                   "NIM: " & txtNIM & vbCrLf & _
                   "Nama: " & txtNama & vbCrLf & _
                   "Jenis Kelamin: " & strKelaminPanjang, _
                   vbYesNo + vbExclamation, "Peringatan!")
    
    strData = Trim(txtNIM) & "#" & Trim(txtNama) & "#" & strKelamin
    
    If tanya = vbYes Then
        'MsgBox "HAPUS"
        HapusData strAlamatFile, strData
    End If
    
    'Refresh ListBox
    MasukkanIsiFileKeListBox strAlamatFile
    
End Sub

Private Sub cmdSimpan_Click()
    '---
    Dim strData As String, strKelamin As String
    
    If optJenisKelamin(0).Value = True Then
        strKelamin = "l"
    Else
        strKelamin = "p"
    End If
    
    strData = Trim(txtNIM) & "#" & Trim(txtNama) & "#" & strKelamin
    
    If bTambah Then
        bTambah = False
        TambahData strAlamatFile, strData
    ElseIf bEdit Then
        bEdit = False
        txtNIM.Locked = False
        UbahData strAlamatFile, strData, strDiubah
    End If
    
    KunciTextBox True
    TombolNormal True
    lstDaftar.Enabled = True
    'Refresh ListBox
    MasukkanIsiFileKeListBox strAlamatFile
End Sub

Private Sub cmdTambah_Click()
    bTambah = True
    KunciTextBox False
    TombolNormal False
    'Kosongkan textbox
    txtNIM.Text = vbNullChar
    txtNama.Text = vbNullChar
    optJenisKelamin(0).Value = True
End Sub

Private Sub cmdTutup_Click()
    End
End Sub

Private Sub cmdUbah_Click()
    Dim strKelamin As String
    
    bEdit = True
    KunciTextBox False
    txtNIM.Locked = True
    TombolNormal False
    'Kosongkan textbox
    'txtNIM.Text = vbNullChar
    'txtNama.Text = vbNullChar
    'optJenisKelamin(0).Value = True
    
    If optJenisKelamin(0).Value = True Then
        strKelamin = "l"
    Else
        strKelamin = "p"
    End If
    
    strDiubah = Trim(txtNIM) & "#" & Trim(txtNama) & "#" & strKelamin
End Sub

Private Sub Form_Load()
    Dim iMsg As Integer
    strAlamatFile = App.Path & "\" & strNamaFile
    'Cek keberadaan file
    If Not CekFileAda(strAlamatFile) Then
        iMsg = MsgBox("File " & strAlamatFile & " tidak ditemukan!" & vbCrLf & vbCrLf & _
               "Apakah Anda ingin memuat data contoh?", _
               vbYesNo + vbExclamation, "Peringatan!")
        If iMsg = vbYes Then
            LoadSampleDataMahasiswa strAlamatFile
            MasukkanIsiFileKeListBox strAlamatFile
        End If
    Else
        MasukkanIsiFileKeListBox strAlamatFile
    End If
    'MsgBox lstDaftar.Text
    'Kunci TextBox
    KunciTextBox True
    'Tombol Normal
    TombolNormal True
    cboCari.ListIndex = 0
End Sub

Private Sub MasukkanIsiFileKeListBox(strAlamatFile As String)
    lstDaftar.Clear
    BacaIsiFile strAlamatFile, lstDaftar
End Sub

Private Sub lstDaftar_Click()
    Dim strTeks As String, arrDataTeks() As String
    strTeks = lstDaftar.Text
    If bTambah Or bEdit Then Exit Sub
    If (CekFormatTeks(strTeks, arrDataTeks)) Then
        txtNIM = arrDataTeks(0)
        txtNama = arrDataTeks(1)
        CekJenisKelamin arrDataTeks(2)
    End If
End Sub

Private Sub CekJenisKelamin(sJenisKelamin As String)
    If sJenisKelamin = "l" Then
        optJenisKelamin(0).Value = True
    Else
        optJenisKelamin(1).Value = True
    End If
End Sub

Sub KunciTextBox(boolNilai As Boolean)
    txtNIM.Enabled = Not boolNilai
    txtNama.Enabled = Not boolNilai
    Frame1.Enabled = Not boolNilai
End Sub

Sub TombolNormal(boolNilai As Boolean)
    cmdTambah.Enabled = boolNilai
    cmdUbah.Enabled = boolNilai
    cmdSimpan.Enabled = Not boolNilai
    cmdBatal.Enabled = Not boolNilai
    cmdHapus.Enabled = boolNilai
End Sub
