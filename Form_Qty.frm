VERSION 5.00
Begin VB.Form Form_Qty 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Enter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txt_Quantity 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the quantity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form_Qty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ANGKA As Double

Private Sub Command1_Click()
MsgBox ANGKA
End Sub

Private Sub Form_Resize()
Me.Left = (MDIForm1.Width - Me.Width) / 2
Me.Top = (MDIForm1.Height - Me.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form_Stock.Enabled = True
End Sub

Private Sub txt_Quantity_KeyPress(KeyAscii As Integer)
On Error Resume Next

ANGKA = txt_Quantity.Text

If Right(txt_Quantity.Text, 5) = "ENTER" Then

If txt_Quantity.Text <> "" Then




Form_Stock.Dat_Stock.Recordset("qty") = ANGKA
Form_Stock.Dat_Stock.Recordset.Update

Form_Stock.txt_Search.Text = ""
Form_Stock.Enabled = True
Unload Me
End If
Exit Sub
End If

If Check1.Value <> 1 Then


txt_Quantity.SelStart = Len(txt_Quantity.Text) + 1
txt_Quantity.SetFocus




Exit Sub
End If


If KeyAscii = 13 Then

If txt_Quantity.Text <> "" Then
Form_Stock.Dat_Stock.Recordset("qty") = txt_Quantity.Text
Form_Stock.Dat_Stock.Recordset.Update

Form_Stock.txt_Search.Text = ""
Form_Stock.Enabled = True
Unload Me
End If
End If
End Sub


