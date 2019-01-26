VERSION 5.00
Begin VB.Form Form_Barcode_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barcode Maker"
   ClientHeight    =   6720
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11895
   Begin VB.CommandButton Command3 
      Caption         =   "PRINT PREVIEW"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   16
      Top             =   5280
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEND IMAGE"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "BC TYPE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   6360
      TabIndex        =   11
      Top             =   840
      Width           =   1935
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1935
         ScaleWidth      =   1575
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         Begin VB.OptionButton Option1 
            Caption         =   "39"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            ToolTipText     =   "Code 3 of 9"
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "i25"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   3
            ToolTipText     =   "Interleaved 2 of 5"
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "128"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   4
            ToolTipText     =   "Code 128"
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Codabar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Codabar/NW-7"
            Top             =   1320
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SETTINGS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   8520
      TabIndex        =   12
      Top             =   840
      Width           =   3015
      Begin VB.CheckBox Check1 
         Caption         =   "2x Size"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Double size"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Label"
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Human readable"
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Check Character"
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
         Index           =   2
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Optional"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Optimize for Printing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Increases spacing"
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COPY TO CLIPBOARD"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Left            =   120
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BARCODE DATA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   14
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form_Barcode_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''
'' == BARCODE MAKER ==                     ''
'' By Paul Bahlawan                        ''
'' Aug 2003 (Code 3 of 9) v0.1              ''
''                                         ''
'' Update Oct 2004: v0.2                    ''
''  -Add "Check Character" option          ''
''  -Add "Optimize for Printing" option    ''
''  -Add "Copy to clipboard" command       ''
''  -Various minor changes                 ''
''                                         ''
'' Update Nov 2004: v0.3                    ''
''  -Add Interleaved 2 of 5                ''
''  -Add Code 128                          ''
''  -simplify 3 of 9 sub                   ''
''  -Add Codabar/NW-7                      ''
''                                         ''
'' Update Dec 2005: v0.4                    ''
''  -Add checksum to Interleaved 2 of 5    ''
''  -put start & stop chrs back in Codabar ''
''                                         ''
'' TO DO: add 2x functionality             ''
''                                         ''
'' Based on specs from:                    ''
'' www.adams1.com/pub/russadam/            ''
'' www.barcodeman.com/info/barspec.php     ''
'''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim BCtype As Long


Private Sub makeBC()
    Select Case BCtype
        Case 0
            make39
        Case 1
            makei25
        Case 2
            make128
        Case 3
            makeCodabar
    End Select
End Sub


Private Sub make39()
Dim x As Long, y As Long, pos As Long
Dim Bardata As String
Dim Cur As String
Dim CurVal As Long
Dim chksum As Long
Dim chkchr As String
Dim temp As String
Dim BC(43) As String
    '3 of the 9 elements are wide: 0=narrow, 1=wide
    BC(0) = "000110100" '0
    BC(1) = "100100001" '1
    BC(2) = "001100001" '2
    BC(3) = "101100000" '3
    BC(4) = "000110001" '4
    BC(5) = "100110000" '5
    BC(6) = "001110000" '6
    BC(7) = "000100101" '7
    BC(8) = "100100100" '8
    BC(9) = "001100100" '9
    BC(10) = "100001001" 'A
    BC(11) = "001001001" 'B
    BC(12) = "101001000" 'C
    BC(13) = "000011001" 'D
    BC(14) = "100011000" 'E
    BC(15) = "001011000" 'F
    BC(16) = "000001101" 'G
    BC(17) = "100001100" 'H
    BC(18) = "001001100" 'I
    BC(19) = "000011100" 'J
    BC(20) = "100000011" 'K
    BC(21) = "001000011" 'L
    BC(22) = "101000010" 'M
    BC(23) = "000010011" 'N
    BC(24) = "100010010" 'O
    BC(25) = "001010010" 'P
    BC(26) = "000000111" 'Q
    BC(27) = "100000110" 'R
    BC(28) = "001000110" 'S
    BC(29) = "000010110" 'T
    BC(30) = "110000001" 'U
    BC(31) = "011000001" 'V
    BC(32) = "111000000" 'W
    BC(33) = "010010001" 'X
    BC(34) = "110010000" 'Y
    BC(35) = "011010000" 'Z
    BC(36) = "010000101" '-
    BC(37) = "110000100" '.
    BC(38) = "011000100" '<SPC>
    BC(39) = "010101000" '$
    BC(40) = "010100010" '/
    BC(41) = "010001010" '+
    BC(42) = "000101010" '%
    BC(43) = "010010100" '*  (used for start/stop character only)
    
    Picture1.Cls
    If Text1.Text = "" Then Exit Sub
    pos = 20
    Bardata = UCase(Text1.Text)
    
    'Check for invalid characters, build temp string & calculate check sum
    For x = 1 To Len(Bardata)
        Cur = Mid$(Bardata, x, 1)
        Select Case Cur
            Case "0" To "9"
                CurVal = Val(Cur)
            Case "A" To "Z"
                CurVal = Asc(Cur) - 55
            Case "-"
                CurVal = 36
            Case "."
                CurVal = 37
            Case " "
                CurVal = 38
            Case "$"
                CurVal = 39
            Case "/"
                CurVal = 40
            Case "+"
                CurVal = 41
            Case "%"
                CurVal = 42
            Case Else 'oops!
                Picture1.Print Cur & " is Invalid"
                Exit Sub
        End Select
        temp = temp & BC(CurVal) & "0" '"0"= add intercharactor gap (1 narrow space)
        chksum = chksum + CurVal
    Next
    
    'Add Check Character? (rarely used, but i put it here anyway...)
    If Check1(2).Value Then
        chksum = chksum Mod 43
        temp = temp & BC(chksum) & "0"
        chkchr = Mid$("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*", chksum + 1, 1)
    End If
    
    'Add Start & Stop characters (must have 'em for valid barcodes)
    temp = BC(43) & "0" & temp & BC(43)
    
    'Generate Barcode
    For x = 1 To Len(temp)
        If x Mod 2 = 0 Then
            'SPACE
            pos = pos + 1 + (2 * Val(Mid$(temp, x, 1))) + Check1(0).Value
        Else
            'BAR
            For y = 1 To 1 + (2 * Val(Mid$(temp, x, 1)))
                Picture1.Line (pos, 1)-(pos, 58 - Check1(1) * 8)
                pos = pos + 1
            Next
        End If
    Next

    'Add Label?
    If Check1(1).Value Then
        Picture1.CurrentX = 35 + Len(Bardata) * (5 + Check1(0).Value * 2) 'kinda center
        Picture1.CurrentY = 50
        Picture1.Print Bardata & chkchr;
    End If
End Sub


Private Sub makei25()
Dim x As Long, y As Long, pos As Long
Dim Bardata As String
Dim Cur As String
Dim temp As String
Dim chksum As Long
Dim BC(11) As String
    '2 of the 5 elements are wide: 0=narrow, 1=wide
    BC(0) = "00110" '0
    BC(1) = "10001" '1
    BC(2) = "01001" '2
    BC(3) = "11000" '3
    BC(4) = "00101" '4
    BC(5) = "10100" '5
    BC(6) = "01100" '6
    BC(7) = "00011" '7
    BC(8) = "10010" '8
    BC(9) = "01010" '9
    BC(10) = "0000" 'Start chr
    BC(11) = "100" 'Stop chr
    
    Picture1.Cls
    If Text1.Text = "" Then Exit Sub
    pos = 20
    Bardata = Text1.Text
    
    'make even num of digits by adding a leading 0
    If Len(Bardata) Mod 2 And Not Check1(2).Value Then Bardata = "0" & Bardata
    If Not (Len(Bardata) Mod 2) And Check1(2).Value Then Bardata = "0" & Bardata
    
    'Check for invalid characters and calculate check sum
    For x = 1 To Len(Bardata)
        Cur = Mid$(Bardata, x, 1)
        If Cur < "0" Or Cur > "9" Then
            Picture1.Print Cur & " is Invalid"
            Exit Sub
        End If
        'make checksum
        If x Mod 2 Then
            chksum = chksum + CLng(Cur) * 3
        Else
            chksum = chksum + CLng(Cur)
        End If
    Next
    
    'add check chr to bardata (if selected)
    If Check1(2).Value Then
        chksum = (10 - chksum Mod 10) Mod 10
        Bardata = Bardata & Chr$(48 + chksum)
    End If
    
    'interleave the code into a temp string - what'd you think the name meant?
    For x = 1 To Len(Bardata) Step 2
        For y = 1 To 5
            temp = temp & Mid$(BC(Val(Mid$(Bardata, x, 1))), y, 1)
            temp = temp & Mid$(BC(Val(Mid$(Bardata, x + 1, 1))), y, 1)
        Next
    Next
    
    'add Start & Stop characters
    temp = BC(10) & temp & BC(11)
    
    'Generate Barcode
    For x = 1 To Len(temp)
        If x Mod 2 = 0 Then
                'SPACE
                pos = pos + 1 + (2 * Val(Mid$(temp, x, 1))) + Check1(0).Value
        Else
                'BAR
                For y = 1 To 1 + (2 * Val(Mid$(temp, x, 1)))
                    Picture1.Line (pos, 1)-(pos, 58 - Check1(1) * 8)
                    pos = pos + 1
                Next
        End If
    Next

    'Add Label?
    If Check1(1).Value Then
        Picture1.CurrentX = 20 + Len(Bardata) * (2 + Check1(0).Value * 1.3) 'kinda center
        Picture1.CurrentY = 50
        Picture1.Print Bardata;
    End If
End Sub


Private Sub make128()
Dim x As Long, y As Long, pos As Long
Dim Bardata As String
Dim Cur As String
Dim CurVal As Long
Dim chksum As Long
Dim temp As String
Dim BC(106) As String
    'code 128 is basically the ASCII chr set.
    '4 element sizes : 1=narrowest, 4=widest
    BC(0) = "212222" '<SPC>
    BC(1) = "222122" '!
    BC(2) = "222221" '"
    BC(3) = "121223" '#
    BC(4) = "121322" '$
    BC(5) = "131222" '%
    BC(6) = "122213" '&
    BC(7) = "122312" ''
    BC(8) = "132212" '(
    BC(9) = "221213" ')
    BC(10) = "221312" '*
    BC(11) = "231212" '+
    BC(12) = "112232" ',
    BC(13) = "122132" '-
    BC(14) = "122231" '.
    BC(15) = "113222" '/
    BC(16) = "123122" '0
    BC(17) = "123221" '1
    BC(18) = "223211" '2
    BC(19) = "221132" '3
    BC(20) = "221231" '4
    BC(21) = "213212" '5
    BC(22) = "223112" '6
    BC(23) = "312131" '7
    BC(24) = "311222" '8
    BC(25) = "321122" '9
    BC(26) = "321221" ':
    BC(27) = "312212" ';
    BC(28) = "322112" '<
    BC(29) = "322211" '=
    BC(30) = "212123" '>
    BC(31) = "212321" '?
    BC(32) = "232121" '@
    BC(33) = "111323" 'A
    BC(34) = "131123" 'B
    BC(35) = "131321" 'C
    BC(36) = "112313" 'D
    BC(37) = "132113" 'E
    BC(38) = "132311" 'F
    BC(39) = "211313" 'G
    BC(40) = "231113" 'H
    BC(41) = "231311" 'I
    BC(42) = "112133" 'J
    BC(43) = "112331" 'K
    BC(44) = "132131" 'L
    BC(45) = "113123" 'M
    BC(46) = "113321" 'N
    BC(47) = "133121" 'O
    BC(48) = "313121" 'P
    BC(49) = "211331" 'Q
    BC(50) = "231131" 'R
    BC(51) = "213113" 'S
    BC(52) = "213311" 'T
    BC(53) = "213131" 'U
    BC(54) = "311123" 'V
    BC(55) = "311321" 'W
    BC(56) = "331121" 'X
    BC(57) = "312113" 'Y
    BC(58) = "312311" 'Z
    BC(59) = "332111" '[
    BC(60) = "314111" '\
    BC(61) = "221411" ']
    BC(62) = "431111" '^
    BC(63) = "111224" '_
    BC(64) = "111422" '`
    BC(65) = "121124" 'a
    BC(66) = "121421" 'b
    BC(67) = "141122" 'c
    BC(68) = "141221" 'd
    BC(69) = "112214" 'e
    BC(70) = "112412" 'f
    BC(71) = "122114" 'g
    BC(72) = "122411" 'h
    BC(73) = "142112" 'i
    BC(74) = "142211" 'j
    BC(75) = "241211" 'k
    BC(76) = "221114" 'l
    BC(77) = "413111" 'm
    BC(78) = "241112" 'n
    BC(79) = "134111" 'o
    BC(80) = "111242" 'p
    BC(81) = "121142" 'q
    BC(82) = "121241" 'r
    BC(83) = "114212" 's
    BC(84) = "124112" 't
    BC(85) = "124211" 'u
    BC(86) = "411212" 'v
    BC(87) = "421112" 'w
    BC(88) = "421211" 'x
    BC(89) = "212141" 'y
    BC(90) = "214121" 'z
    BC(91) = "412121" '{
    BC(92) = "111143" '|
    BC(93) = "111341" '}
    BC(94) = "131141" '~
    BC(95) = "114113" '<DEL>        *not used in this sub
    BC(96) = "114311" 'FNC 3        *not used in this sub
    BC(97) = "411113" 'FNC 2        *not used in this sub
    BC(98) = "411311" 'SHIFT        *not used in this sub
    BC(99) = "113141" 'CODE C       *not used in this sub
    BC(100) = "114131" 'FNC 4       *not used in this sub
    BC(101) = "311141" 'CODE A      *not used in this sub
    BC(102) = "411131" 'FNC 1       *not used in this sub
    BC(103) = "211412" 'START A     *not used in this sub
    BC(104) = "211214" 'START B
    BC(105) = "211232" 'START C     *not used in this sub
    BC(106) = "2331112" 'STOP

    Picture1.Cls
    If Text1.Text = "" Then Exit Sub
    pos = 20
    Bardata = Text1.Text

    'Check for invalid characters, calculate check sum & build temp string
    For x = 1 To Len(Bardata)
        Cur = Mid$(Bardata, x, 1)
        If Cur < " " Or Cur > "~" Then
            Picture1.Print "Invalid Character(s)"
            Exit Sub
        End If
        CurVal = Asc(Cur) - 32
        temp = temp + BC(CurVal)
        chksum = chksum + CurVal * x
    Next
    
    'Add start, stop & check characters
    chksum = (chksum + 104) Mod 103
    temp = BC(104) & temp & BC(chksum) & BC(106)

    'Generate Barcode
    For x = 1 To Len(temp)
        If x Mod 2 = 0 Then
                'SPACE
                pos = pos + (Val(Mid$(temp, x, 1))) + Check1(0).Value
        Else
                'BAR
                For y = 1 To (Val(Mid$(temp, x, 1)))
                    Picture1.Line (pos, 1)-(pos, 58 - Check1(1) * 8)
                    pos = pos + 1
                Next
        End If
    Next

    'Add Label?
    If Check1(1).Value Then
        Picture1.CurrentX = 30 + Len(Bardata) * (3 + Check1(0).Value * 2) 'kinda center
        Picture1.CurrentY = 50
        Picture1.Print Bardata;
    End If
End Sub


Private Sub makeCodabar()
Dim x As Long, y As Long, pos As Long
Dim Bardata As String
Dim Cur As String
Dim CurVal As Long
Dim temp As String
Dim BC(19) As String
    'Codabar, also known as NW-7
    BC(0) = "0000011" '0
    BC(1) = "0000110" '1
    BC(2) = "0001001" '2
    BC(3) = "1100000" '3
    BC(4) = "0010010" '4
    BC(5) = "1000010" '5
    BC(6) = "0100001" '6
    BC(7) = "0100100" '7
    BC(8) = "0110000" '8
    BC(9) = "1001000" '9
    BC(10) = "0001100" '-
    BC(11) = "0011000" '$
    BC(12) = "1000101" ':
    BC(13) = "1010001" '/
    BC(14) = "1010100" '.
    BC(15) = "0010101" '+
    BC(16) = "0011010" 'start/stop A
    BC(17) = "0101001" 'start/stop B
    BC(18) = "0001011" 'start/stop C
    BC(19) = "0001110" 'start/stop D
    
    Picture1.Cls
    If Text1.Text = "" Then Exit Sub
    pos = 20
    Bardata = Text1.Text

    For x = 1 To Len(Bardata)
        Cur = Mid$(Bardata, x, 1)
        Select Case Cur
            Case "0" To "9"
                CurVal = Val(Cur)
            Case "a" To "d"
                CurVal = Asc(Cur) - 81
            Case "-"
                CurVal = 10
            Case "$"
                CurVal = 11
            Case ":"
                CurVal = 12
            Case "/"
                CurVal = 13
            Case "."
                CurVal = 14
            Case "+"
                CurVal = 15
            Case Else 'oops!
                Picture1.Print Cur & " is Invalid"
                Exit Sub
        End Select
        temp = temp & BC(CurVal) & "0" '"0"= add intercharactor gap (1 narrow space)
    Next

    'Add Start & Stop characters (using "A" for both here)
    temp = BC(16) & "0" & temp & BC(16)
    
    'Generate Barcode
    For x = 1 To Len(temp)
        If x Mod 2 = 0 Then
            'SPACE
            pos = pos + 1 + (2 * Val(Mid$(temp, x, 1))) + Check1(0).Value
        Else
            'BAR
            For y = 1 To 1 + (2 * Val(Mid$(temp, x, 1)))
                Picture1.Line (pos, 1)-(pos, 58 - Check1(1) * 8)
                pos = pos + 1
            Next
        End If
    Next

    'Add Label?
    If Check1(1).Value Then
        Picture1.CurrentX = 30 + Len(Bardata) * (3 + Check1(0).Value * 2) 'kinda center
        Picture1.CurrentY = 50
        Picture1.Print Bardata;
    End If
End Sub


Private Sub Command2_Click()

Form_Barcode.Show
End Sub

Private Sub Command3_Click()
Dim i As Integer
For i = 0 To 9
Text1.Text = i
Form_Barcode.Show
Form_Barcode.Picture1(i).Picture = Form_Barcode_Main.Picture1.Image

Next i

Text1.Text = "ENTER"
Form_Barcode.Picture1(10).Picture = Form_Barcode_Main.Picture1.Image


End Sub

Private Sub Form_Resize()
    Picture1.Width = Me.Width - 360
    makeBC
    Me.Left = (MDIForm1.Width - Me.Width) / 2
Me.Top = (MDIForm1.Height - Me.Height) / 2
End Sub


Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Check1(2).ToolTipText = "Optional"
            Check1(2).Value = 0
            Check1(2).Enabled = True
        Case 1
            Check1(2).ToolTipText = "Optional"
            Check1(2).Value = 0
            Check1(2).Enabled = True
        Case 2
            Check1(2).ToolTipText = "Not optional"
            Check1(2).Value = 1
            Check1(2).Enabled = False
        Case 3
            Check1(2).ToolTipText = "Not used"
            Check1(2).Value = 0
            Check1(2).Enabled = False
    End Select
    BCtype = Index
    makeBC
End Sub


Private Sub Text1_Change()
    makeBC
End Sub


Private Sub Check1_Click(Index As Integer)
    makeBC
End Sub


Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetData Picture1.Image
End Sub
