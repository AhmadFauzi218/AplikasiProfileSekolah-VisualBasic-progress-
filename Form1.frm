VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FormLoginProfile"
   ClientHeight    =   5190
   ClientLeft      =   4335
   ClientTop       =   2685
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   12180
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      MaskColor       =   &H00808080&
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   720
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Silahkan Login Terlebih Dahulu"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim login As Integer
Private Sub Command1_Click()
User = Text1.Text
Password = Text2.Text
If User = "pengunjung" And Password = "pengunjung" Then
MsgBox "Selamat Datang"
Form2.Show
Form1.Hide
Else
login = login + 1
MsgBox "anda salah memasukan password dan username" & login & "kali"
If login = 2 Then
MsgBox "kesempatan anda satu kali lagi", vbExclamation
End If
If login = 3 Then
MsgBox "Anda sudah salah memasukan password 3 kali, Aplikasi ini akan ditutup ", vbExclamation
End
End If
End If
End Sub

Private Sub Command2_Click()
pesan = MsgBox("Anda Yakin Mau Keluar ??", vbQuestion + vbYesNo, "Question")
If pesan = vbYes Then
End
Else
Form1.SetFocus
End If
End Sub
