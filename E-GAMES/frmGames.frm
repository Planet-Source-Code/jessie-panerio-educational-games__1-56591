VERSION 5.00
Begin VB.Form frmGames 
   BackColor       =   &H00400000&
   Caption         =   $"frmGames.frx":0000
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGames.frx":008E
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmGames.frx":0F58
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label rs_filldatabase 
      Caption         =   $"frmGames.frx":2223FA
      Height          =   375
      Left            =   11880
      TabIndex        =   0
      Top             =   8040
      Width           =   6015
   End
   Begin VB.Menu mnuGames 
      Caption         =   "&Games"
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVC 
         Caption         =   "&Vowels and  Consonants"
         Index           =   1
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuessingGame 
         Caption         =   "&Guessing Game"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScoreVC 
         Caption         =   "&High Score Vowels and Consonant"
      End
      Begin VB.Menu mnuwalalang 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScoreG 
         Caption         =   "&High Score Guessing Game"
      End
      Begin VB.Menu line8 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuheidi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHowto 
         Caption         =   "Need Help?"
      End
      Begin VB.Menu mnuheidy 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    Set cnn = New ADODB.Connection
    cnn.CursorLocation = adUseClient
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\vowelsconsonants.mdb"
    cnn.Open
    frmGames.Caption = frmConfiguration.rs_filldatabase.Caption
  
End Sub


Private Sub Form_Resize()
    
    On Error Resume Next
    ShockwaveFlash1.Left = Me.ScaleLeft
    ShockwaveFlash1.Top = Me.ScaleTop
    ShockwaveFlash1.Width = Me.ScaleWidth
    
End Sub

Private Sub mnuConfig_Click()
     
    frmPassword.Show
    frmPassword.txtPassword.SetFocus
    Me.Enabled = False
   
End Sub

Private Sub mnuGuessingGame_Click()
    
   frmguessinggame.Show
   frmGames.Enabled = False

End Sub

Private Sub mnuHighScoreG_Click()
   
   frmHighScore.Show
   
End Sub

Private Sub mnuHighScoreVC_Click()
   
   frmHighScoreVC.Show
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Unload Me
   
End Sub

Private Sub mnuHowto_Click()
   
   Call WinHelp(0, App.HelpFile, HelpC, 0)
   
End Sub

Private Sub mnuVC_Click(Index As Integer)
   
   frmVowelsConsonants.Show
   frmGames.Enabled = False
   
End Sub
