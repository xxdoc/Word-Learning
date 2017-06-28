VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "xvoice.dll"
Begin VB.Form frmSpeakAndSpell 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Word Learning - Anthoni Wiese"
   ClientHeight    =   4935
   ClientLeft      =   5940
   ClientTop       =   2595
   ClientWidth     =   4470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   2280
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS syn 
      Height          =   255
      Left            =   0
      OleObjectBlob   =   "frmSpeakAndSpell.frx":0000
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrSpeaker 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.ListBox lstWords 
      Height          =   4545
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   200.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4590
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu mnuDebuggers 
      Caption         =   "debugger"
      Visible         =   0   'False
      Begin VB.Menu mnuSet 
         Caption         =   "Set Time_Between_Words"
         Index           =   0
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Set Time_Between_Letters"
         Index           =   1
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Set Time_Until_Input"
         Index           =   2
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Set Time_Between_Word_And_Letters"
         Index           =   3
      End
      Begin VB.Menu mnuSeeWordList 
         Caption         =   "See Word List"
      End
   End
End
Attribute VB_Name = "frmSpeakAndSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wordStage As Long
Private wordIndex As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyNumpad0 Then
        mnuDebuggers.Visible = Not mnuDebuggers.Visible
    End If
End Sub

Private Sub Form_Load()
    ' Load Database
    Dim fNum As Long
    Dim fCnt As String, fLns() As String
    fNum = FreeFile
    Open DATABASE_LOCATION For Input As FreeFile
    fCnt = Input(LOF(fNum), fNum)
    Close #fNum
    fLns = Split(fCnt, vbCrLf)
    For i = 0 To UBound(fLns)
        lstWords.AddItem fLns(i)
    Next
    
    wordIndex = 0
    wordStage = 0
    
    ' Initiate Word Learning Process
    tmrInit.Enabled = True
End Sub

Private Sub mnuSeeWordList_Click()
    mnuSeeWordList.Checked = Not mnuSeeWordList.Checked
    lstWords.Visible = mnuSeeWordList.Checked
    If lstWords.Visible = True Then
        lblChar.Width = 2175
    Else
        lblChar.Width = 4215
    End If
End Sub

Private Sub mnuSet_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0:
            TIME_BETWEEN_WORDS = InputBox("Enter Time Between Words (In Milliseconds)", "TIME_BETWEEN_WORDS", TIME_BETWEEN_WORDS)
        Case 1:
            TIME_BETWEEN_LETTERS = InputBox("Enter Time Between Words (In Milliseconds)", "TIME_BETWEEN_LETTERS", TIME_BETWEEN_LETTERS)
        Case 2:
            TIME_UNTIL_INPUT = InputBox("Enter Time Until Input (In Milliseconds)", "TIME_UNTIL_INPUT", TIME_UNTIL_INPUT)
        Case 3:
            TIME_BETWEEN_WORD_AND_LETTERS = InputBox("Enter Time Between Word and Letters (In Milliseconds)", "TIME_BETWEEN_WORD_AND_LETTERS", TIME_BETWEEN_WORD_AND_LETTERS)
    End Select
End Sub

Private Sub tmrInit_Timer()
    If tmrInit.Tag = "1" Then
        tmrSpeaker.Enabled = True
        tmrInit.Enabled = False
        Exit Sub
    End If
    syn.Speak "You will be givehn A werd, and thehn have it spelled owt to you. After,  you will be asked to ehnter the werd. Pleese lis ehn starting, now."
    tmrInit.Tag = "1"
    tmrInit.Interval = 10000
End Sub

Private Sub tmrSpeaker_Timer()
    Dim s As String
    
    Select Case wordStage
        Case 0:
            ' Pronounciate whole word
            If wordIndex >= lstWords.ListCount Then
                syn.Speak "You have won the game! Returning to the main menu."
                tmrSpeaker.Enabled = False
                Me.Hide
                frmMain.Show
                Exit Sub
            End If
            
            syn.Speak WordToModPro(lstWords.List(wordIndex))
            tmrSpeaker.Enabled = False
            tmrSpeaker.Interval = TIME_BETWEEN_WORD_AND_LETTERS
            tmrSpeaker.Enabled = True
            wordStage = wordStage + 1
        Case Is <= Len(lstWords.List(wordIndex))
            ' Pronounciate Letter By Letter
            syn.Speak LetterToWord(Mid$(lstWords.List(wordIndex), wordStage, 1))
            lblChar.Caption = UCase$(Mid$(lstWords.List(wordIndex), wordStage, 1))
            wordStage = wordStage + 1
            tmrSpeaker.Enabled = False
            tmrSpeaker.Interval = TIME_BETWEEN_LETTERS
            tmrSpeaker.Enabled = True
        Case Is = Len(lstWords.List(wordIndex)) + 1
            syn.Speak WordToModPro(lstWords.List(wordIndex))
            wordStage = wordStage + 1
            tmrSpeaker.Interval = TIME_UNTIL_INPUT
            lblChar.Caption = ""
        Case Is = Len(lstWords.List(wordIndex)) + 2
            syn.Speak "Pleese  ehnter thuh  werd that wuhs sehd, now."
            s = InputBox("Enter the word:", "Word")
            s = LCase$(s)
            If s = LCase$(lstWords.List(wordIndex)) Then
                syn.Speak "You got it right! And, thuh next werd will begin shortly."
            Else
                syn.Speak "I'm sorry, you did not ehnter it correctly. Get ready."
            End If
            
            tmrSpeaker.Enabled = False
            tmrSpeaker.Interval = TIME_BETWEEN_WORDS
            tmrSpeaker.Enabled = True
            wordStage = 0
            wordIndex = wordIndex + 1
    End Select
End Sub

Private Function WordToModPro(word As String) As String
    WordToModPro = word
    word = LCase$(word)
    Dim fNum As Long, fCnt As String, fLns() As String
    Dim prts() As String
    
    fNum = FreeFile
    Open DATABASE_MODPRO_LOCATION For Input As fNum
    fCnt = Input(LOF(fNum), fNum)
    Close #fNum
    fLns = Split(fCnt, vbCrLf)
    For i = 0 To UBound(fLns)
        prts = Split(fLns(i), "|")
        If LCase$(prts(0)) = word Then
            WordToModPro = prts(1)
            Exit Function
        End If
    Next
End Function

Private Function LetterToWord(letter As String) As String
    letter = UCase$(letter)
    Select Case letter
        Case "A":
            LetterToWord = "A"
        Case "B":
            LetterToWord = "Bee"
        Case "C":
            LetterToWord = "See"
        Case "D":
            LetterToWord = "Dee"
        Case "E":
            LetterToWord = "ee"
        Case "F":
            LetterToWord = "Eff"
        Case "G":
            LetterToWord = "Jee"
        Case "H":
            LetterToWord = "H"
        Case "I":
            LetterToWord = "Eye"
        Case "J":
            LetterToWord = "Jay"
        Case "K":
            LetterToWord = "Kay"
        Case "L":
            LetterToWord = "El"
        Case "M":
            LetterToWord = "Ehm"
        Case "N":
            LetterToWord = "Ehn"
        Case "O":
            LetterToWord = "Oh"
        Case "P":
            LetterToWord = "Pee"
        Case "Q":
            LetterToWord = "Cue"
        Case "R":
            LetterToWord = "Ar"
        Case "S":
            LetterToWord = "Es"
        Case "T":
            LetterToWord = "Tee"
        Case "U":
            LetterToWord = "You"
        Case "V":
            LetterToWord = "Vee"
        Case "W":
            LetterToWord = "Double U"
        Case "X":
            LetterToWord = "Ex"
        Case "Y":
            LetterToWord = "Why"
        Case "Z":
            LetterToWord = "Zee"
    End Select
End Function
