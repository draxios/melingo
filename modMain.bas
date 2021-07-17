Attribute VB_Name = "modMain"
Public Wordlist() As String
Public WordCount As Long
Public Word As String
Public Round As Integer
Public UniBuff As Long
Public Board1(27) As Integer
Public buffz As Integer
Public bufz As Integer
Public FirstGuess As Boolean

Public Sub Winner(newplayer As Boolean)
If newplayer = True Then
    Dim buffz As Integer
    Dim bufz As Integer
    For buffz = 0 To 24
        Randomize Timer
        bufz = (Rnd() * 60) + 1
        Board1(buffz) = bufz
    Next buffz
    Board1(25) = -1
    Board1(26) = -1
    Board1(27) = -1
    For buffz = 0 To 9
Restarter:
        Randomize Timer
        bufz = (Rnd() * 24)
        If frmBoard.txt(buffz).Tag <> "-1" Then
            frmBoard.txt(buffz).BackColor = &H449342
            frmBoard.txt(buffz).Tag = "-1"
        Else
            GoTo Restarter
        End If
        
        
    Next buffz
    frmBoard.tLoad.Enabled = True
    FirstGuess = True
End If


frmBoard.Show

End Sub
