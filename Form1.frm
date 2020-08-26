VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Jika Anda memasukkan data: 123456789 ke dalam
  'MaskEdBox1, maka ketika kursor meninggalkan
  'MaskEdBox1, style data menjadi: Rp 123.456.789,00
  'dan ketika Anda menyimpan data tersebut ke suatu
  'database atau ingin mengambil data ini untuk
  'keperluan perhitungan, maka data yang diolah
  'merupakan data aslinya, yaitu: 123456789
  MsgBox MaskEdBox1.Text  '--> Menampilkan: 123456789
End Sub

Private Sub Form_Load()
  MaskEdBox1.Format = "Rp #,##0.00;(Rp #,##0.00)"
End Sub

