VERSION 5.00
Begin VB.Form frmlevel3g 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 3  Guessing Game"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
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
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "frmlevel3g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
  With frmguessinggame
  
  If KeyAscii = 13 Then
    frmguessinggame.Enabled = True
   
    .txtLevel.Text = "3"
    .txtCorrectAnswer = "0"
    .txtQuestionNumber = "0"
    .txtWrongAnswer.Text = "0"
    .cmdCheckAnswer.Enabled = False
    .cmdStart.Enabled = True
    .lblQuestion.Caption = ""
    .Show
     For i = 0 To 8
         .Text1(i).Text = ""
         .Text1(i).Locked = True
     Next i
     
    Unload Me
  End If
  End With
End Sub

Private Sub Form_Load()
  frmguessinggame.txtLevel2VC = 100 - (Val(frmguessinggame.txtWrongAnswer.Text) * 10)
  frmguessinggame.Enabled = False
  frmguessinggame.txtCorrectAnswer.Text = ""
  frmguessinggame.txtWrongAnswer.Text = ""
End Sub


