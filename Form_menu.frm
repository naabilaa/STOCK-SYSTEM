VERSION 5.00
Begin VB.Form Form_menu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4650
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BARCODE MAKER"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form_Stock.Show
End Sub

Private Sub Command2_Click()
Form_Barcode_Main.Show

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Resize()
Me.Left = (MDIForm1.Width - Me.Width) / 2
Me.Top = (MDIForm1.Height - Me.Height) / 2
End Sub
