VERSION 5.00
Begin VB.Form frmFungsiProsedur 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   2160
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ekstrak"
      Height          =   3495
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "061730700553"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Jumlah Karakter/Digit"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Nomor Urut Mahasiswa"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Kelompok Kelas"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Tahun Masuk"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Jurusan"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Jurusan"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan NIM."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmFungsiProsedur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Left(Text1.Text, 2)
Text3.Text = getNamaJurusan(Text2.Text)
Text4.Text = Mid(Text1.Text, 3, 2)
Text5.Text = Mid(Text1.Text, 5, 4)
Text6.Text = Right(Text1.Text, 4)
Text7.Text = Len(Text1.Text)
End Sub
Function getNamaJurusan(ByVal KodeJurusan As String) As String
Dim NamaJurusan As String
If KodeJurusan = "01" Then
    NamaJurusan = "Teknik Mesin"
ElseIf KodeJurusan = "02" Then
    NamaJurusan = "Teknik Sipil"
ElseIf KodeJurusan = "03" Then
    NamaJurusan = "Teknik Kimia"
ElseIf KodeJurusan = "04" Then
    NamaJurusan = "Teknik Elektro"
ElseIf KodeJurusan = "05" Then
    NamaJurusan = "Akuntansi"
ElseIf KodeJurusan = "06" Then
    NamaJurusan = "Teknik Komputer"
ElseIf KodeJurusan = "07" Then
    NamaJurusan = "Manajemen Informatika"
ElseIf KodeJurusan = "08" Then
    NamaJurusan = "Administrasi Bisni"
ElseIf KodeJurusan = "09" Then
    NamaJurusan = "Bahasa Inggris"
End If
getNamaJurusan = NamaJurusan
End Function
