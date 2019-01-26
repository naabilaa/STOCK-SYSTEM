VERSION 5.00
Begin VB.Form Form_Barcode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   14070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   14070
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   10
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   12840
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   9
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   11520
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   8
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10320
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   7
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9120
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   6
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7920
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   5
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6720
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   4
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5520
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   3
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   2
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   1
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   0
      Left            =   4440
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form_Barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Visible = False
Me.PrintForm
Command1.Visible = True
End Sub

Private Sub Form_Load()
    'Clipboard.Clear
    
Me.Height = 14500
End Sub

Private Sub Form_Resize()
Me.Left = (MDIForm1.Width - Me.Width) / 2
Me.Top = (MDIForm1.Height - Me.Height) / 2
End Sub
