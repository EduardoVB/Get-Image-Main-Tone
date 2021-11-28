VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6672
   ClientLeft      =   2832
   ClientTop       =   2160
   ClientWidth     =   8232
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6672
   ScaleWidth      =   8232
   Begin VB.CommandButton Command1 
      Caption         =   "Change image"
      Height          =   480
      Left            =   360
      TabIndex        =   1
      Top             =   72
      Width           =   2856
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   324
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4800
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   720
      Width           =   7200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "It sets the background color according to the main color of the image. Also this label ForeColor"
      Height          =   552
      Left            =   3348
      TabIndex        =   2
      Top             =   72
      Width           =   4296
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim d As New cDlg
    
    d.Filter = "Image files|*.jpg;*.bmp;*.ico;*.wmf"
    d.ShowOpen
    If Not d.Canceled Then
        Set Picture1.Picture = LoadPicture(d.FileName)
        Update
    End If
End Sub

Private Sub Form_Load()
    Update
End Sub

Private Sub Update()
    Me.BackColor = PicMainTone(Picture1.Picture)
    Label1.ForeColor = IIf(IsDarkColor(Me.BackColor), vbWhite, vbBlack)
End Sub
