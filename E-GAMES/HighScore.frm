VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHighScore 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scores"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HighScore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintRG 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FF8080&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FF8080&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H00FF8080&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdTop 
      BackColor       =   &H00FF8080&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12632064
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Guessing Game Scores"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   1133
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLast_Click()
If adoresults.RecordCount <= 1 Then Exit Sub
adoresults.MoveLast
End Sub

Private Sub cmdNext_Click()
If adoresults.AbsolutePosition >= adoresults.RecordCount Or adoresults.RecordCount <= 1 Then Exit Sub
adoresults.MoveNext
End Sub

Private Sub cmdPrev_Click()
    If adoresults.AbsolutePosition <= 1 Then Exit Sub
    adoresults.MovePrevious
End Sub

Private Sub cmdTop_Click()
If adoresults.RecordCount <= 1 Then Exit Sub
adoresults.MoveFirst
End Sub


Private Sub cmdPrintRG_Click()

With adoresults
Set DataReport1.DataSource = adoresults
DataReport1.Show
End With

End Sub



Private Sub Form_Load()
 frmGames.Enabled = False
 Set adoresults = New ADODB.Recordset
 adoresults.Open "Select * from results", cnn, adOpenStatic, adLockPessimistic
 Set DataGrid1.DataSource = adoresults
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmGames.Enabled = True
Unload Me
End Sub
