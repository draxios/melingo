VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Melingo!"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3555
      TabIndex        =   37
      Top             =   495
      Width           =   870
   End
   Begin VB.Timer tTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5235
      Top             =   315
   End
   Begin VB.Frame frmBoard 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   45
      TabIndex        =   30
      Top             =   975
      Width           =   3405
      Begin VB.TextBox txt2 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   855
         MaxLength       =   1
         TabIndex        =   3
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txt3 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   4
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txt4 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2100
         MaxLength       =   1
         TabIndex        =   5
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txt5 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2730
         MaxLength       =   1
         TabIndex        =   6
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txt1 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   2
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txt6 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   34
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txt10 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2730
         MaxLength       =   1
         TabIndex        =   10
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txt9 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2100
         MaxLength       =   1
         TabIndex        =   9
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txt8 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   8
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txt7 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   855
         MaxLength       =   1
         TabIndex        =   7
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txt11 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   33
         Top             =   1500
         Width           =   465
      End
      Begin VB.TextBox txt15 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2730
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1500
         Width           =   465
      End
      Begin VB.TextBox txt14 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2100
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1500
         Width           =   465
      End
      Begin VB.TextBox txt13 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1500
         Width           =   465
      End
      Begin VB.TextBox txt12 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   855
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1500
         Width           =   465
      End
      Begin VB.TextBox txt16 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   32
         Top             =   2160
         Width           =   465
      End
      Begin VB.TextBox txt20 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2730
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2160
         Width           =   465
      End
      Begin VB.TextBox txt19 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2100
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2160
         Width           =   465
      End
      Begin VB.TextBox txt18 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2160
         Width           =   465
      End
      Begin VB.TextBox txt17 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   855
         MaxLength       =   1
         TabIndex        =   15
         Top             =   2160
         Width           =   465
      End
      Begin VB.TextBox txt21 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   31
         Top             =   2820
         Width           =   465
      End
      Begin VB.TextBox txt25 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   2730
         MaxLength       =   1
         TabIndex        =   22
         Top             =   2820
         Width           =   465
      End
      Begin VB.TextBox txt24 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   2100
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2820
         Width           =   465
      End
      Begin VB.TextBox txt23 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2820
         Width           =   465
      End
      Begin VB.TextBox txt22 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   855
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2820
         Width           =   465
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         Caption         =   "Melon"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   35
         Top             =   3570
         Width           =   3105
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00A9DEBE&
      Caption         =   "Guess!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3555
      TabIndex        =   29
      Top             =   1110
      Width           =   870
   End
   Begin VB.TextBox txtBuffer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3570
      TabIndex        =   28
      Top             =   4665
      Width           =   780
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   5235
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3555
      TabIndex        =   1
      Top             =   45
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   45
      TabIndex        =   0
      Top             =   -225
      Width           =   3405
      Begin VB.TextBox txtD1 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   210
         MaxLength       =   1
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "m"
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtD5 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2730
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "n"
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtD4 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "o"
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtD3 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "l"
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtD2 
         BackColor       =   &H00D6D8AF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "e"
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Shape bar2 
      BackColor       =   &H00A9DEBE&
      BackStyle       =   1  'Opaque
      Height          =   15
      Left            =   3795
      Top             =   2190
      Width           =   360
   End
   Begin VB.Shape bar1 
      Height          =   2220
      Left            =   3765
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label bigx 
      Caption         =   "X"
      Height          =   555
      Left            =   3795
      TabIndex        =   36
      Top             =   1605
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Buffer As String
Dim Suffer As Long

If WordCount = 0 Then

    d1.Filter = "Wordlists (*.txt) | *.txt"
    d1.ShowOpen

    If d1.FileName = "" Then Exit Sub

        txtD1 = "": txtD2 = "": txtD3 = "": txtD4 = "": txtD5 = ""


        Open d1.FileName For Input As #1

        Do Until (EOF(1)) = True
    
            ReDim Preserve Wordlist(WordCount)
            Line Input #1, Buffer
            Wordlist(WordCount) = Buffer
            WordCount = WordCount + 1
    
        Loop
    
        Close #1
    
End If
    
    Reset
    Randomize Timer
    Suffer = Rnd() * WordCount
    Word = Wordlist(Suffer)
    txtD1 = Left(Wordlist(Suffer), 1)
    txt1 = Left(Wordlist(Suffer), 1)
    txt6 = Left(Wordlist(Suffer), 1)
    txt11 = Left(Wordlist(Suffer), 1)
    txt16 = Left(Wordlist(Suffer), 1)
    txt21 = Left(Wordlist(Suffer), 1)
    txtBuffer = Word
    Round = 1
    
    tTimer.Enabled = True

UniBuff = -1
'txtD2 = Mid(Wordlist(Suffer), 2, 1)
'txtD3 = Mid(Wordlist(Suffer), 3, 1)
'txtD4 = Mid(Wordlist(Suffer), 4, 1)
'txtD5 = Right(Wordlist(Suffer), 1)





End Sub

Private Sub Reset()
txt1 = "": txt2 = "": txt3 = "": txt4 = "": txt5 = ""
txt6 = "": txt7 = "": txt8 = "": txt9 = "": txt10 = ""
txt11 = "": txt12 = "": txt13 = "": txt14 = "": txt15 = ""
txt16 = "": txt17 = "": txt18 = "": txt19 = "": txt20 = ""
txt21 = "": txt22 = "": txt23 = "": txt24 = "": txt25 = ""
txt1.BackColor = &HD6D8AF: txt2.BackColor = &HD6D8AF: txt3.BackColor = &HD6D8AF
txt4.BackColor = &HD6D8AF: txt5.BackColor = &HD6D8AF: txt6.BackColor = &HD6D8AF
txt7.BackColor = &HD6D8AF: txt8.BackColor = &HD6D8AF: txt9.BackColor = &HD6D8AF
txt10.BackColor = &HD6D8AF: txt11.BackColor = &HD6D8AF: txt12.BackColor = &HD6D8AF
txt13.BackColor = &HD6D8AF: txt14.BackColor = &HD6D8AF: txt15.BackColor = &HD6D8AF
txt16.BackColor = &HD6D8AF: txt17.BackColor = &HD6D8AF: txt18.BackColor = &HD6D8AF
txt19.BackColor = &HD6D8AF: txt20.BackColor = &HD6D8AF: txt21.BackColor = &HD6D8AF
txt22.BackColor = &HD6D8AF: txt23.BackColor = &HD6D8AF: txt24.BackColor = &HD6D8AF
txt25.BackColor = &HD6D8AF
txt1.ForeColor = &H80000008: txt2.ForeColor = &H80000008: txt3.ForeColor = &H80000008: txt4.ForeColor = &H80000008: txt5.ForeColor = &H80000008
txt6.ForeColor = &H80000008: txt7.ForeColor = &H80000008: txt8.ForeColor = &H80000008: txt9.ForeColor = &H80000008: txt10.ForeColor = &H80000008
txt15.ForeColor = &H80000008: txt11.ForeColor = &H80000008: txt12.ForeColor = &H80000008: txt13.ForeColor = &H80000008: txt14.ForeColor = &H80000008
txt16.ForeColor = &H80000008: txt17.ForeColor = &H80000008: txt18.ForeColor = &H80000008: txt19.ForeColor = &H80000008: txt20.ForeColor = &H80000008
Round = 1
txtD1.Text = ""
txtD2.Text = ""
txtD3.Text = ""
txtD4.Text = ""
txtD5.Text = ""
lblWord.Caption = ""
End Sub
'&H00A9DEBE& green
'&H00D6D8AF& blue
Private Sub Command2_Click()


Select Case Round

Case 1:
'Second Letter---------------------------------------------------------
If txt2.Text = Mid(Word, 2, 1) Then
   txtD2.Text = Mid(Word, 2, 1)
   txt7 = Mid(Word, 2, 1)
   txt12 = Mid(Word, 2, 1)
   txt17 = Mid(Word, 2, 1)
   txt22 = Mid(Word, 2, 1)
End If
If txt2.Text = Mid(Word, 3, 1) Then txt2.BackColor = &HA9DEBE
If txt2.Text = Mid(Word, 4, 1) Then txt2.BackColor = &HA9DEBE
If txt2.Text = Mid(Word, 5, 1) Then txt2.BackColor = &HA9DEBE
'Third Letter---------------------------------------------------------
If txt3.Text = Mid(Word, 3, 1) Then
   txtD3.Text = Mid(Word, 3, 1)
   txt8 = Mid(Word, 3, 1)
   txt13 = Mid(Word, 3, 1)
   txt18 = Mid(Word, 3, 1)
   txt23 = Mid(Word, 3, 1)
End If
If txt3.Text = Mid(Word, 2, 1) Then txt3.BackColor = &HA9DEBE
If txt3.Text = Mid(Word, 4, 1) Then txt3.BackColor = &HA9DEBE
If txt3.Text = Mid(Word, 5, 1) Then txt3.BackColor = &HA9DEBE
'Fourth Letter---------------------------------------------------------
If txt4.Text = Mid(Word, 4, 1) Then
   txtD4.Text = Mid(Word, 4, 1)
   txt9 = Mid(Word, 4, 1)
   txt14 = Mid(Word, 4, 1)
   txt19 = Mid(Word, 4, 1)
   txt24 = Mid(Word, 4, 1)
End If
If txt4.Text = Mid(Word, 3, 1) Then txt4.BackColor = &HA9DEBE
If txt4.Text = Mid(Word, 2, 1) Then txt4.BackColor = &HA9DEBE
If txt4.Text = Mid(Word, 5, 1) Then txt4.BackColor = &HA9DEBE
'Final Letter---------------------------------------------------------
If txt5.Text = Mid(Word, 5, 1) Then
   txtD5.Text = Mid(Word, 5, 1)
   txt10 = Mid(Word, 5, 1)
   txt15 = Mid(Word, 5, 1)
   txt20 = Mid(Word, 5, 1)
   txt25 = Mid(Word, 5, 1)
End If
If txt5.Text = Mid(Word, 2, 1) Then txt5.BackColor = &HA9DEBE
If txt5.Text = Mid(Word, 3, 1) Then txt5.BackColor = &HA9DEBE
If txt5.Text = Mid(Word, 4, 1) Then txt5.BackColor = &HA9DEBE
'----------------------------------------------------------------------

If txt1 & txt2 & txt3 & txt4 & txt5 = Word Then
    lblWord.Caption = "You won!"
    UniBuff = -2
    Exit Sub
End If

Round = Round + 1
txt1.ForeColor = &HAEC855: txt2.ForeColor = &HAEC855: txt3.ForeColor = &HAEC855: txt4.ForeColor = &HAEC855: txt5.ForeColor = &HAEC855
'-------------------------------------------


Case 2:
'Second Letter---------------------------------------------------------
If txt7.Text = Mid(Word, 2, 1) Then
   txtD2.Text = Mid(Word, 2, 1)
   txt12 = Mid(Word, 2, 1)
   txt17 = Mid(Word, 2, 1)
   txt22 = Mid(Word, 2, 1)
End If
If txt7.Text = Mid(Word, 3, 1) Then txt7.BackColor = &HA9DEBE
If txt7.Text = Mid(Word, 4, 1) Then txt7.BackColor = &HA9DEBE
If txt7.Text = Mid(Word, 5, 1) Then txt7.BackColor = &HA9DEBE
'Third Letter---------------------------------------------------------
If txt8.Text = Mid(Word, 3, 1) Then
   txtD3.Text = Mid(Word, 3, 1)
   txt13 = Mid(Word, 3, 1)
   txt18 = Mid(Word, 3, 1)
   txt23 = Mid(Word, 3, 1)
End If
If txt8.Text = Mid(Word, 2, 1) Then txt8.BackColor = &HA9DEBE
If txt8.Text = Mid(Word, 4, 1) Then txt8.BackColor = &HA9DEBE
If txt8.Text = Mid(Word, 5, 1) Then txt8.BackColor = &HA9DEBE
'Fourth Letter---------------------------------------------------------
If txt9.Text = Mid(Word, 4, 1) Then
   txtD4.Text = Mid(Word, 4, 1)
   txt14 = Mid(Word, 4, 1)
   txt19 = Mid(Word, 4, 1)
   txt24 = Mid(Word, 4, 1)
End If
If txt9.Text = Mid(Word, 3, 1) Then txt9.BackColor = &HA9DEBE
If txt9.Text = Mid(Word, 2, 1) Then txt9.BackColor = &HA9DEBE
If txt9.Text = Mid(Word, 5, 1) Then txt9.BackColor = &HA9DEBE
'Final Letter---------------------------------------------------------
If txt10.Text = Mid(Word, 5, 1) Then
   txtD5.Text = Mid(Word, 5, 1)
   txt15 = Mid(Word, 5, 1)
   txt20 = Mid(Word, 5, 1)
   txt25 = Mid(Word, 5, 1)
End If
If txt10.Text = Mid(Word, 2, 1) Then txt10.BackColor = &HA9DEBE
If txt10.Text = Mid(Word, 3, 1) Then txt10.BackColor = &HA9DEBE
If txt10.Text = Mid(Word, 4, 1) Then txt10.BackColor = &HA9DEBE
'----------------------------------------------------------------------

If txt6 & txt7 & txt8 & txt9 & txt10 = Word Then
    lblWord.Caption = "You won!"
    UniBuff = -2
    Exit Sub
End If

Round = Round + 1
txt6.ForeColor = &HAEC855: txt7.ForeColor = &HAEC855: txt8.ForeColor = &HAEC855: txt9.ForeColor = &HAEC855: txt10.ForeColor = &HAEC855
'-----------------------------------------------


Case 3:
'Second Letter---------------------------------------------------------
If txt12.Text = Mid(Word, 2, 1) Then
    txtD2.Text = Mid(Word, 2, 1)
    txt17 = Mid(Word, 2, 1)
    txt22 = Mid(Word, 2, 1)
End If
If txt12.Text = Mid(Word, 3, 1) Then txt12.BackColor = &HA9DEBE
If txt12.Text = Mid(Word, 4, 1) Then txt12.BackColor = &HA9DEBE
If txt12.Text = Mid(Word, 5, 1) Then txt12.BackColor = &HA9DEBE
'Third Letter---------------------------------------------------------
If txt13.Text = Mid(Word, 3, 1) Then
   txtD3.Text = Mid(Word, 3, 1)
   txt18 = Mid(Word, 3, 1)
   txt23 = Mid(Word, 3, 1)
End If
If txt13.Text = Mid(Word, 2, 1) Then txt13.BackColor = &HA9DEBE
If txt13.Text = Mid(Word, 4, 1) Then txt13.BackColor = &HA9DEBE
If txt13.Text = Mid(Word, 5, 1) Then txt13.BackColor = &HA9DEBE
'Fourth Letter---------------------------------------------------------
If txt14.Text = Mid(Word, 4, 1) Then
   txtD4.Text = Mid(Word, 4, 1)
   txt19 = Mid(Word, 4, 1)
   txt24 = Mid(Word, 4, 1)
End If
If txt14.Text = Mid(Word, 3, 1) Then txt14.BackColor = &HA9DEBE
If txt14.Text = Mid(Word, 2, 1) Then txt14.BackColor = &HA9DEBE
If txt14.Text = Mid(Word, 5, 1) Then txt14.BackColor = &HA9DEBE
'Final Letter---------------------------------------------------------
If txt15.Text = Mid(Word, 5, 1) Then
   txtD5.Text = Mid(Word, 5, 1)
   txt20 = Mid(Word, 5, 1)
   txt25 = Mid(Word, 5, 1)
End If
If txt15.Text = Mid(Word, 2, 1) Then txt15.BackColor = &HA9DEBE
If txt15.Text = Mid(Word, 3, 1) Then txt15.BackColor = &HA9DEBE
If txt15.Text = Mid(Word, 4, 1) Then txt15.BackColor = &HA9DEBE
'----------------------------------------------------------------------

If txt11 & txt12 & txt13 & txt14 & txt15 = Word Then
    lblWord.Caption = "You won!"
    UniBuff = -2
    Exit Sub
End If

Round = Round + 1
txt11.ForeColor = &HAEC855: txt12.ForeColor = &HAEC855: txt13.ForeColor = &HAEC855: txt14.ForeColor = &HAEC855: txt15.ForeColor = &HAEC855
'------------------------------------------


Case 4:
'Second Letter---------------------------------------------------------
If txt17.Text = Mid(Word, 2, 1) Then
    txtD2.Text = Mid(Word, 2, 1)
    txt22 = Mid(Word, 2, 1)
End If
If txt17.Text = Mid(Word, 3, 1) Then txt17.BackColor = &HA9DEBE
If txt17.Text = Mid(Word, 4, 1) Then txt17.BackColor = &HA9DEBE
If txt17.Text = Mid(Word, 5, 1) Then txt17.BackColor = &HA9DEBE
'Third Letter---------------------------------------------------------
If txt18.Text = Mid(Word, 3, 1) Then
    txtD3.Text = Mid(Word, 3, 1)
    txt23 = Mid(Word, 3, 1)
End If
If txt18.Text = Mid(Word, 2, 1) Then txt18.BackColor = &HA9DEBE
If txt18.Text = Mid(Word, 4, 1) Then txt18.BackColor = &HA9DEBE
If txt18.Text = Mid(Word, 5, 1) Then txt18.BackColor = &HA9DEBE
'Fourth Letter---------------------------------------------------------
If txt19.Text = Mid(Word, 4, 1) Then
    txtD4.Text = Mid(Word, 4, 1)
    txt24 = Mid(Word, 4, 1)
End If
If txt19.Text = Mid(Word, 3, 1) Then txt19.BackColor = &HA9DEBE
If txt19.Text = Mid(Word, 2, 1) Then txt19.BackColor = &HA9DEBE
If txt19.Text = Mid(Word, 5, 1) Then txt19.BackColor = &HA9DEBE
'Final Letter---------------------------------------------------------
If txt20.Text = Mid(Word, 5, 1) Then
    txtD5.Text = Mid(Word, 5, 1)
    txt25 = Mid(Word, 5, 1)
End If
If txt20.Text = Mid(Word, 2, 1) Then txt20.BackColor = &HA9DEBE
If txt20.Text = Mid(Word, 3, 1) Then txt20.BackColor = &HA9DEBE
If txt20.Text = Mid(Word, 4, 1) Then txt20.BackColor = &HA9DEBE
'----------------------------------------------------------------------

If txt16 & txt17 & txt18 & txt19 & txt20 = Word Then
    lblWord.Caption = "You won!"
    UniBuff = -2
    Exit Sub
End If
Round = Round + 1
txt16.ForeColor = &HAEC855: txt17.ForeColor = &HAEC855: txt18.ForeColor = &HAEC855: txt19.ForeColor = &HAEC855: txt20.ForeColor = &HAEC855


'-----------------------------------
Case 5:
If txt21 & txt22 & txt23 & txt24 & txt25 = Word Then
    lblWord.Caption = "You won!"
    UniBuff = -2
    Exit Sub
Else
    lblWord.Caption = "You lose!"
    lblWord.Caption = Word
    UniBuff = -2
    Exit Sub
End If

End Select

UniBuff = -1


End Sub



Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
bigx.Visible = True
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
bigx.Visible = False
End Sub

Private Sub Command3_Click()
tTimer.Enabled = False
bigx.Visible = False
UniBuff = 0
bar2.Height = 1
End Sub

Private Sub Form_Load()
Me.Caption = "Melingo v" & App.Major & "." & App.Minor & App.Revision
'lblWord.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub Form_Terminate()
Unload Me
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub


Private Sub lblWord_DblClick()
If txtBuffer.Visible = True Then
    txtBuffer.Visible = False
Else
    txtBuffer.Visible = True
End If
End Sub

Private Sub tTimer_Timer()
If UniBuff = -2 Then
    tTimer.Enabled = False
    bigx.Visible = False
    UniBuff = 0
    bar2.Height = 1
    If Board1(0) = 0 Then
        Winner True
    Else
        Winner False
    End If
    Exit Sub
End If

If UniBuff = -1 Then
    tTimer.Enabled = False
    bigx.Visible = False
    UniBuff = 0
    bar2.Height = 1
    tTimer.Enabled = True
End If

UniBuff = UniBuff + 1
'ten times 200
bar2.Height = UniBuff * 100
If UniBuff > 20 Then
    bigx.Visible = True
    Command2_Click
    UniBuff = -1
End If
    

End Sub
