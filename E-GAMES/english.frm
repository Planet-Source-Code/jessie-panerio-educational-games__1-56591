VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form frmVowelsConsonants 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vowels-Consonants"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlashV 
      Height          =   1335
      Left            =   5520
      TabIndex        =   47
      Top             =   5520
      Width           =   2895
      _cx             =   5106
      _cy             =   2355
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
      Height          =   1455
      Left            =   8520
      TabIndex        =   42
      Top             =   5400
      Width           =   2535
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2160
         Top             =   960
      End
      Begin VB.Label Label12 
         BackColor       =   &H00400000&
         Caption         =   "Date:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00400000&
         Caption         =   "Time:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   615
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
         Left            =   960
         TabIndex        =   44
         Top             =   840
         Width           =   1095
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
         Left            =   960
         TabIndex        =   43
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00400000&
      Height          =   1455
      Left            =   120
      TabIndex        =   36
      Top             =   5400
      Width           =   5295
      Begin VB.CommandButton cmdCheckAnswer 
         BackColor       =   &H00C0C000&
         Caption         =   "Check &Answer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         Picture         =   "english.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C000&
         Caption         =   "Clea&r"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2880
         Picture         =   "english.frx":0BC2
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C000&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         Picture         =   "english.frx":1784
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00C0C000&
         Caption         =   "&Start Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Picture         =   "english.frx":2346
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   41
         Text            =   "Text3"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400000&
      Height          =   2415
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   8295
      Begin VB.TextBox txtVowel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtConsonants 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   31
         Text            =   "Text5"
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   33
         Text            =   "Text4"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         Caption         =   "How Many Consonants are there?"
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
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   6735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400000&
         Caption         =   "How Many Vowels are there?"
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
         TabIndex        =   34
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   10
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   9
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   8
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   7
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   6
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   5
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   4
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2415
      Left            =   8520
      TabIndex        =   10
      Top             =   2760
      Width           =   2535
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1800
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
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   360
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
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   735
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
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "Computed Score:"
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txtCorrectAnswer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   495
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtWrongAnswer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   495
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtQuestionNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   615
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
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   480
         Width           =   1695
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
         Left            =   8400
         TabIndex        =   26
         Top             =   480
         Width           =   1575
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
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   480
         Width           =   1215
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
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   10935
   End
   Begin VB.Label rs_filldatabase 
      Caption         =   $"english.frx":2650
      Height          =   375
      Left            =   -4560
      TabIndex        =   48
      Top             =   2520
      Width           =   4575
   End
End
Attribute VB_Name = "frmVowelsConsonants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i
Dim wordlength
Dim answervowel, answerconsonant
Dim rec

Private Sub cmdCheckAnswer_Click()

If Val(txtVowel.Text) = Val(Text4.Text) Then
  If Val(txtConsonants.Text) = Val(Text5.Text) Then
    
    
    txtCorrectAnswer.Text = Val(txtCorrectAnswer.Text) + 1
    MsgBox "Correct", vbInformation, "Yeahhh"
      
      If (txtCorrectAnswer.Text = "10") And (txtLevel.Text = "1") Then
         frmlevel2.Show
      ElseIf (txtCorrectAnswer.Text = "10") And (txtLevel.Text = "2") Then
         frmlevel3.Show
      ElseIf (txtCorrectAnswer.Text = "10") And (txtLevel.Text = "3") Then
         frmlevel1.Show
    
       Else
      
    cmdCheckAnswer.Enabled = True
    cmdStart.Enabled = True
    txtVowel.Text = ""
    txtConsonants.Text = ""
 
    cmdStart_Click
  
      End If
  
  
  Else
    txtWrongAnswer.Text = Val(txtWrongAnswer.Text) + 1
    
    MsgBox "Wrong", vbInformation, "Opps!"

    txtVowel.Text = ""
    txtConsonants.Text = ""
    txtVowel.SetFocus
    cmdCheckAnswer.Enabled = True
    cmdStart.Enabled = False
  End If
Else
    txtWrongAnswer.Text = Val(txtWrongAnswer.Text) + 1
    MsgBox "Wrong", vbInformation, "Opps!"
    txtVowel.Text = ""
    txtConsonants.Text = ""
    txtVowel.SetFocus
    cmdCheckAnswer.Enabled = True
    cmdStart.Enabled = False
End If

End Sub

Private Sub cmdClear_Click()
txtVowel.Text = ""
txtConsonants.Text = ""
txtVowel.SetFocus
End Sub

Private Sub cmdClose_Click()
  
  Unload Me
  frmGames.Enabled = True
  frmGames.Show

End Sub

Private Sub cmdStart_Click()
On Error Resume Next
txtVowel.SetFocus
txtVowel.Locked = False
txtConsonants.Locked = False

If txtLevel.Text = "1" Then

txtQuestionNumber.Text = Val(txtQuestionNumber.Text) + 1


    For i = 1 To 8
        Text2(i).Locked = True
    Next i
     
     
    adoword.MoveLast
    adoword.MoveFirst
    'Random engine initialized
     Randomize
    'Generate random numbers
     adoword.Move Int((adoword.RecordCount * Rnd)) '+ 1)
   
    'Complete Word
    Text3.Text = adoword!level1
    wordlength = Len(Text3.Text)
 
    answervowel = 0

    For i = 1 To 10
        'Place each letter to the textbox
        Text2(i).Text = Mid(adoword!level1, i + 0, 1)
          If Text2(i).Text <> "" Then
            Text2(i).BackColor = &HFFC0C0
          Else
            Text2(i).BackColor = &H400000
          End If
        'Determine and count vowels
        If Text2(i).Text = "A" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "E" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "I" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "O" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "U" Then
        answervowel = answervowel + 1
        End If
   Next i

   Text4.Text = answervowel


   Text5.Text = Val(wordlength) - Val(Text4.Text)




ElseIf txtLevel.Text = "2" Then

txtQuestionNumber.Text = Val(txtQuestionNumber.Text) + 1


 For i = 1 To 10
        Text2(i).Locked = True
    Next i

    adoword.MoveLast
    adoword.MoveFirst
    'Random engine Initialized
     Randomize
    'Generate random numbers
     adoword.Move Int((adoword.RecordCount * Rnd)) '+ 1)
   
    'Complete Word
    Text3.Text = adoword!level2
    wordlength = Len(Text3.Text)
 
    answervowel = 0

    For i = 1 To 10
        'Place each letter to the textbox
        Text2(i).Text = Mid(adoword!level2, i + 0, 1)
          If Text2(i).Text <> "" Then
            Text2(i).BackColor = &HFFC0C0
          Else
            Text2(i).BackColor = &H400000
          End If
        'Determine and count the vowels
        If Text2(i).Text = "A" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "E" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "I" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "O" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "U" Then
        answervowel = answervowel + 1
        End If
   Next i

   Text4.Text = answervowel


   Text5.Text = Val(wordlength) - Val(Text4.Text)


   

ElseIf txtLevel.Text = "3" Then


txtQuestionNumber.Text = Val(txtQuestionNumber.Text) + 1

    For i = 1 To 10
        Text2(i).Locked = True
    Next i

    adoword.MoveLast
    adoword.MoveFirst
    'Random engine initialized
     Randomize
    'Generate random numbers
     adoword.Move Int((adoword.RecordCount * Rnd)) '+ 1)
   
    'Complete Word
    Text3.Text = adoword!level3
    wordlength = Len(Text3.Text)
 
    answervowel = 0

    For i = 1 To 10
        'Place each letter to the textbox
        Text2(i).Text = Mid(adoword!level3, i + 0, 1)
          If Text2(i).Text <> "" Then
            Text2(i).BackColor = &HFFC0C0
          Else
            Text2(i).BackColor = &H400000
          End If
        'Determine and count the vowels
        If Text2(i).Text = "A" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "E" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "I" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "O" Then
        answervowel = answervowel + 1
        ElseIf Text2(i).Text = "U" Then
        answervowel = answervowel + 1
        End If
   Next i

   Text4.Text = answervowel


   Text5.Text = Val(wordlength) - Val(Text4.Text)




End If



    txtVowel.Locked = False
    txtConsonants.Locked = False
    cmdStart.Enabled = False
    cmdCheckAnswer.Enabled = True
    cmdClear.Enabled = True


End Sub

Private Sub Form_Load()
  Set adoword = New ADODB.Recordset
  adoword.Open "Select * from vowelsconsonants", cnn, adOpenStatic, adLockPessimistic
   
  ShockwaveFlashV.Movie = Path & "vowelsconsonantsgame.swf"
  ShockwaveFlashV.Play
    
  txtLevel.Text = "1"

End Sub

Private Sub Timer1_Timer()
lblDate.Caption = Date
lblTime.Caption = Time
End Sub

Private Sub txtConsonants_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub


Private Sub txtVowel_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
