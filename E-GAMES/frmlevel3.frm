VERSION 5.00
Begin VB.Form frmlevel3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 3  Vowels-Consonants"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Level 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Press Enter to continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "frmlevel3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    frmVowelsConsonants.Enabled = True
   
    frmVowelsConsonants.txtLevel.Text = "3"
    frmVowelsConsonants.txtCorrectAnswer = "0"
    frmVowelsConsonants.txtQuestionNumber = "0"
    frmVowelsConsonants.cmdCheckAnswer.Enabled = False
    frmVowelsConsonants.cmdStart.Enabled = True
    frmVowelsConsonants.txtVowel.Locked = True
    frmVowelsConsonants.txtConsonants.Locked = True
    frmVowelsConsonants.txtVowel.Text = ""
    frmVowelsConsonants.txtConsonants.Text = ""
    With frmVowelsConsonants
    For i = 1 To 10
         .Text2(i).Text = ""
         .Text2(i).Locked = True
     Next i
    End With
    frmVowelsConsonants.Show
    Unload Me
  End If
  
End Sub

Private Sub Form_Load()

  frmVowelsConsonants.txtLevel2VC = 100 - (Val(frmVowelsConsonants.txtWrongAnswer.Text) * 10)
  frmVowelsConsonants.Enabled = False
  frmVowelsConsonants.txtCorrectAnswer.Text = ""
  frmVowelsConsonants.txtWrongAnswer.Text = ""
End Sub


