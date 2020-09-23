VERSION 5.00
Begin VB.Form frmlevel1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Results"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4320
      Top             =   0
   End
   Begin VB.TextBox txtName2 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtTotalScore 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox tdate 
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please Enter your Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblTotalScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Your Total Score is:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
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
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Congratulation!"
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
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmlevel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Set adoresults = New ADODB.Recordset
  adoresults.Open "Select * from resultsvc", cnn, adOpenStatic, adLockPessimistic

  frmVowelsConsonants.txtLevel3VC = 100 - (Val(frmVowelsConsonants.txtWrongAnswer.Text) * 10)
  frmVowelsConsonants.txtCSVC.Text = Val(frmVowelsConsonants.txtLevel1VC) + Val(frmVowelsConsonants.txtLevel2VC) + Val(frmVowelsConsonants.txtLevel3VC)
  lblTotalScore.Caption = frmVowelsConsonants.txtCSVC.Text
  frmVowelsConsonants.Enabled = False
  frmVowelsConsonants.txtCorrectAnswer.Text = ""
  frmVowelsConsonants.txtWrongAnswer.Text = ""
  frmVowelsConsonants.txtCSVC.Text = lblTotalScore.Caption
End Sub

Private Sub Timer1_Timer()
  lblName.ForeColor = QBColor(Rnd * 10)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
 With adoresults
 If KeyAscii = 13 Then
   If txtName.Text = "" Then
     txtName.SetFocus
     MsgBox "Please Enter your Name", vbCritical, "Guessing Game"
   ElseIf txtName.Text <> "" Then
       txtName2.Text = txtName.Text
       txtTotalScore.Text = lblTotalScore.Caption
       Do While Not adoresults.EOF
            If (adoresults!Name = txtName.Text) Then
              MsgBox adoresults!Name & " is already listed in the database. Choose Another.", 16
              txtName.Text = ""
              txtName.SetFocus
              Exit Sub
            Else
              adoresults.MoveNext
            End If
       Loop
      
       If lblTotalScore.Caption > 200 Then
         adoresults.AddNew
         LoadResults
         writeresults
         adoresults.Update
       End If
   
    frmVowelsConsonants.Enabled = True
    frmVowelsConsonants.Show
    
    frmVowelsConsonants.cmdCheckAnswer.Enabled = False
    frmVowelsConsonants.cmdStart.Enabled = True
    Unload frmVowelsConsonants
    Unload Me
    frmGames.Show
    frmGames.Enabled = True
    
  End If
End If
End With
End Sub

Private Sub LoadResults()
    On Error Resume Next
    If adoresults.BOF = True Or adoresults.EOF = True Then
        Exit Sub
    End If
    
    txtName.Text = adoresults!Name & ""
    txtTotalScore.Text = adoresults!scores & ""
    tdate.Text = adoresults!Date & ""
End Sub

Private Sub writeresults()
    On Error Resume Next
    tdate.Text = Date
    adoresults!Name = txtName2.Text
    adoresults!scores = lblTotalScore.Caption
    adoresults!Date = tdate.Text
End Sub
    

