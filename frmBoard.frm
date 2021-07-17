VERSION 5.00
Begin VB.Form frmBoard 
   AutoRedraw      =   -1  'True
   Caption         =   "Crazy Melingo Board !"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   Icon            =   "frmBoard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmBoard 
      BackColor       =   &H00A9DEBE&
      Height          =   4260
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   3765
      Begin VB.Timer tLoad 
         Enabled         =   0   'False
         Interval        =   25
         Left            =   225
         Top             =   3720
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   1
         Left            =   825
         MaxLength       =   2
         TabIndex        =   26
         Top             =   180
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   21
         Left            =   840
         MaxLength       =   2
         TabIndex        =   24
         Top             =   2820
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   22
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   23
         Top             =   2805
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   23
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   22
         Top             =   2805
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   24
         Left            =   2895
         MaxLength       =   2
         TabIndex        =   21
         Top             =   2805
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   20
         Left            =   165
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2820
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   16
         Left            =   840
         MaxLength       =   2
         TabIndex        =   19
         Top             =   2190
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   17
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   18
         Top             =   2190
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   18
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2145
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   19
         Left            =   2895
         MaxLength       =   2
         TabIndex        =   16
         Top             =   2145
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   15
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   15
         Top             =   2190
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   11
         Left            =   840
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1560
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   12
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1575
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   13
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1485
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   14
         Left            =   2895
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1485
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   10
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1560
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   6
         Left            =   810
         MaxLength       =   2
         TabIndex        =   9
         Top             =   855
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   7
         Left            =   1545
         MaxLength       =   2
         TabIndex        =   8
         Top             =   855
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   8
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   7
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   9
         Left            =   2955
         MaxLength       =   2
         TabIndex        =   6
         Top             =   810
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   5
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   5
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   0
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Top             =   180
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   4
         Left            =   2955
         MaxLength       =   2
         TabIndex        =   3
         Top             =   150
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   3
         Left            =   2235
         MaxLength       =   2
         TabIndex        =   2
         Top             =   195
         Width           =   700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0077BB68&
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
         Index           =   2
         Left            =   1545
         MaxLength       =   2
         TabIndex        =   1
         Top             =   210
         Width           =   700
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PULL"
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
         TabIndex        =   25
         Top             =   3570
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblWord_Click()
Dim x As Byte

If FirstGuess = True Then
    Randomize Timer
    buffz = Val(txt(Rnd() * 24))
    For x = 0 To 24
    If buffz = Val(txt(x)) Then
       
       txt(x) = "O"
       txt(x).BackColor = &H449342
       txt(x).Tag = "-1"
       FirstGuess = False
    End If
    
    
    Next x
    lblWord.Caption = buffz
    'Check for win
Else
    Randomize Timer
    buffz = Val(txt(Rnd() * 24))
    For x = 0 To 24
    If buffz = Val(txt(x)) Then
       lblWord.Caption = buffz
       txt(x) = "O"
       txt(x).BackColor = &H449342
       txt(x).Tag = "-1"
    End If
    
    Next x
    
    'Check for win
End If

End Sub

Private Sub tLoad_Timer()
If buffz < 25 Then
    txt(buffz) = Board1(buffz)
    buffz = buffz + 1
Else
    buffz = 0
    tLoad.Enabled = False
End If
End Sub
