VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   Caption         =   "Home"
   ClientHeight    =   8490
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Rekapitulasi Sekolah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   14775
      Begin VB.Label Label6 
         Caption         =   "7 Ekstrakulikuler"
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image Image6 
         Height          =   585
         Left            =   5640
         Picture         =   "Form2.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "113 Pelajaran"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   735
         Left            =   4080
         Picture         =   "Form2.frx":35D6
         Stretch         =   -1  'True
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "11 Kelas"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   705
         Left            =   2640
         Picture         =   "Form2.frx":4CD8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "40 Guru"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   1560
         Picture         =   "Form2.frx":66E5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "153 Siswa"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   660
         Left            =   600
         Picture         =   "Form2.frx":80CF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "PROFILE SMK INFORMATIKA AL-IRSYAD AL-ISLAMIYAH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "Form2.frx":A69A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   14760
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu flLgout 
         Caption         =   "Logout"
      End
      Begin VB.Menu flExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu nmSekolah 
      Caption         =   "Sekolah"
      Begin VB.Menu crjurusan 
         Caption         =   "Jurusan"
      End
      Begin VB.Menu fsSekolah 
         Caption         =   "Fasilitas Sekolah"
      End
      Begin VB.Menu vsMisi 
         Caption         =   "Visi Misi"
      End
      Begin VB.Menu ifSekolah 
         Caption         =   "Info Sekolah"
      End
   End
   Begin VB.Menu lKasi 
      Caption         =   "Lokasi "
      Begin VB.Menu lksSekolah 
         Caption         =   "Lokasi Sekolah"
      End
   End
   Begin VB.Menu ctContact 
      Caption         =   "Hubungi"
   End
   Begin VB.Menu ttTentang 
      Caption         =   "Tentang"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clContact_Click()

End Sub

Private Sub flExit_Click()
End
End Sub

Private Sub flLgout_Click()
Form1.Show
Form2.Hide
End Sub

