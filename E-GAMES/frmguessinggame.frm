VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form frmguessinggame 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guessing Game"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlashG 
      Height          =   1935
      Left            =   4440
      TabIndex        =   41
      Top             =   4920
      Width           =   3615
      _cx             =   6376
      _cy             =   3413
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   6120
      Width           =   4215
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3840
         Top             =   360
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400000&
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   2040
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00400000&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   10935
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txtQuestionNumber 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtWrongAnswer 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCorrectAnswer 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00400000&
         Caption         =   "Question #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00400000&
         Caption         =   "Wrong Answer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   8280
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00400000&
         Caption         =   "Correct Answer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00400000&
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   10935
      Begin VB.Label rs_filldatabase 
         Caption         =   $"frmguessinggame.frx":0000
         Height          =   375
         Left            =   -4440
         TabIndex        =   42
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label lblQuestion 
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1815
         Left            =   1440
         TabIndex        =   11
         Top             =   720
         Width           =   9255
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "Question?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   8160
      TabIndex        =   0
      Top             =   4680
      Width           =   2895
      Begin VB.TextBox txtCSVC 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtLevel3VC 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtLevel2VC 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtLevel1VC 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "Computed Score:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "Level 3:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "Level 2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "Level 1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00400000&
      Height          =   1335
      Left            =   120
      TabIndex        =   31
      Top             =   4800
      Width           =   4215
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00C0C000&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Picture         =   "frmguessinggame.frx":008A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCheckAnswer 
         BackColor       =   &H00C0C000&
         Caption         =   "Check Answer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         Picture         =   "frmguessinggame.frx":0394
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C000&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2880
         Picture         =   "frmguessinggame.frx":069E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAnswer 
         Height          =   375
         Left            =   1560
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmguessinggame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i
Dim length
Dim response
Private Sub cmdCheckAnswer_Click()

  response = Text1(0).Text + Text1(1).Text + Text1(2).Text + Text1(3).Text + Text1(4).Text + Text1(5).Text + Text1(6).Text + Text1(7).Text + Text1(8).Text
   
  If txtAnswer.Text = response Then
    
    MsgBox "CORRECT!", vbInformation, "Yeahhh!"
    
       txtCorrectAnswer.Text = Val(txtCorrectAnswer.Text) + 1
      If txtCorrectAnswer.Text = "10" And txtLevel.Text = "1" Then
         frmlevel2g.Show
      ElseIf (txtCorrectAnswer.Text = "10") And (txtLevel.Text = "2") Then
         frmlevel3g.Show
      ElseIf (txtCorrectAnswer.Text = "10") And (txtLevel.Text = "3") Then
         frmlevel1g.Show
      Else
         cmdStart_Click
      End If
  Else
    
     MsgBox "WRONG!", vbInformation, "Opps!"
   
    txtWrongAnswer.Text = Val(txtWrongAnswer.Text) + 1
  End If
  
  
  
End Sub

Private Sub cmdClose_Click()
  Unload Me
  frmGames.Enabled = True
  frmGames.Show
End Sub

Private Sub cmdStart_Click()

cmdStart.Enabled = False
cmdCheckAnswer.Enabled = True
 If adoguess.BOF = True And adoguess.EOF = True Then
 MsgBox "Empty Database"
 cmdCheckAnswer.Enabled = False
 Exit Sub
 Else
    adoguess.MoveLast
    adoguess.MoveFirst
    'Random engine initialized
     Randomize
    'Generate random numbers
     adoguess.Move Int((adoguess.RecordCount * Rnd))
        
    
    lblQuestion.Caption = adoguess!definition
    
    
    For i = 0 To 8
        
        'Place each letter to the textbox
        
        Text1(i).Text = Mid(adoguess!word, i + 1, 1)
          
          If Text1(i).Text <> "" Then
            Text1(i).Locked = True
            Text1(i).BackColor = &HFFC0C0
          Else
            Text1(i).BackColor = &H400000
        
          End If
    Next i
   
   
    
    txtAnswer.Text = adoguess!word
    
    length = Len(txtAnswer.Text)
    
    If (length Mod 2) = 0 Then
    
      For i = 1 To length Step 2
       Text1(i).Text = ""
       Text1(i).Locked = False
      Next i
    
    Else
     
      For i = 0 To length Step 2
       Text1(i).Text = ""
       Text1(i).Locked = False
      Next i
     
    End If
    
    
     
   txtQuestionNumber.Text = Val(txtQuestionNumber.Text) + 1
   End If
End Sub

Private Sub Form_Load()
 Set adoguess = New ADODB.Recordset
 adoguess.Open "Select * from guessinggame", cnn, adOpenStatic, adLockPessimistic
 
 cmdCheckAnswer.Enabled = False
 txtQuestionNumber.Text = "0"
 txtCorrectAnswer.Text = "0"
 txtWrongAnswer.Text = "0"
 txtLevel.Text = "1"
 
 ShockwaveFlashG.Movie = Path & "guessinggame.swf"
 ShockwaveFlashG.Play
   
 
  
End Sub

Private Sub Text1_Change(Index As Integer)
  'validation -letters only
If Val(Text1(Index).Text) <> 0 Then
    Text1(Index).Text = ""
    Exit Sub
End If
Text1(Index).Text = UCase(Text1(Index).Text)
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Time
lblDate.Caption = Date
End Sub
