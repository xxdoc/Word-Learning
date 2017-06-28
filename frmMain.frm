VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "xvoice.dll"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Word Learning - Anthoni Wiese"
   ClientHeight    =   1665
   ClientLeft      =   8265
   ClientTop       =   2745
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   1920
   Begin VB.Timer tmrEnable 
      Interval        =   3000
      Left            =   240
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Database"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   255
      Left            =   0
      OleObjectBlob   =   "frmMain.frx":0000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmSpeakAndSpell.Show
    Me.Hide
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    DirectSS1.Speak "Hello, and welcome to the word learning software."
    Call modSettings.InitVars
End Sub

Private Sub tmrEnable_Timer()
    On Error Resume Next
    For Each obj In Me.Controls
        obj.Enabled = True
    Next
End Sub
