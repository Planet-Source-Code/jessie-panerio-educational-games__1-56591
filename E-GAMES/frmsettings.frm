VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfiguration 
   BackColor       =   &H00000000&
   Caption         =   "Configuration"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      TabHeight       =   794
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Guessing Game"
      TabPicture(0)   =   "frmsettings.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGNew"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGEdit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGSave"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGDelete"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdGFind"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdGClose"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Vowels Consonant Game"
      TabPicture(1)   =   "frmsettings.frx":08DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5"
      Tab(1).Control(1)=   "Text4"
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(3)=   "cmdVClose"
      Tab(1).Control(4)=   "cmdVFind"
      Tab(1).Control(5)=   "cmdVDelete"
      Tab(1).Control(6)=   "cmdVSave"
      Tab(1).Control(7)=   "cmdVCancel"
      Tab(1).Control(8)=   "cmdVEdit"
      Tab(1).Control(9)=   "cmdVNew"
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(11)=   "Label9"
      Tab(1).Control(12)=   "Label8"
      Tab(1).Control(13)=   "Label7"
      Tab(1).Control(14)=   "Label6"
      Tab(1).Control(15)=   "Label5"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Password"
      TabPicture(2)   =   "frmsettings.frx":11B4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPClose"
      Tab(2).Control(1)=   "txtnewpassword"
      Tab(2).Control(2)=   "txtoldpassword"
      Tab(2).Control(3)=   "cmdOK"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "Label1"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdPClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   -71520
         TabIndex        =   33
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74760
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   4440
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74760
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74760
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "frmsettings.frx":1A8E
         Top             =   2760
         Width           =   7095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   360
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1200
         Width           =   6975
      End
      Begin VB.CommandButton cmdVClose 
         Caption         =   "&Close"
         Height          =   855
         Left            =   -68400
         TabIndex        =   19
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdVFind 
         Caption         =   "&Find"
         Height          =   855
         Left            =   -69480
         TabIndex        =   18
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdVDelete 
         Caption         =   "&Delete"
         Height          =   855
         Left            =   -70560
         TabIndex        =   17
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdVSave 
         Caption         =   "&Save"
         Height          =   855
         Left            =   -71640
         TabIndex        =   16
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdVCancel 
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   -72720
         TabIndex        =   15
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdVEdit 
         Caption         =   "&Edit"
         Height          =   855
         Left            =   -73800
         TabIndex        =   14
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdVNew 
         Caption         =   "&New"
         Height          =   855
         Left            =   -74880
         TabIndex        =   13
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtnewpassword 
         Height          =   375
         Left            =   -72240
         TabIndex        =   10
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtoldpassword 
         Height          =   375
         Left            =   -72240
         TabIndex        =   9
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   615
         Left            =   -73920
         TabIndex        =   8
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CommandButton cmdGClose 
         Caption         =   "&Close"
         Height          =   735
         Left            =   6720
         TabIndex        =   7
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdGFind 
         Caption         =   "&Find"
         Height          =   735
         Left            =   5640
         TabIndex        =   6
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdGDelete 
         Caption         =   "&Delete"
         Height          =   735
         Left            =   4560
         TabIndex        =   5
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdGSave 
         Caption         =   "&Save"
         Height          =   735
         Left            =   3360
         TabIndex        =   4
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton cmdGCancel 
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   2280
         TabIndex        =   3
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdGEdit 
         Caption         =   "&Edit"
         Height          =   735
         Left            =   1200
         TabIndex        =   2
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdGNew 
         Caption         =   "&New"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Maximum of 10 Letters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -73920
         TabIndex        =   32
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Maximum of 7 Letters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -73920
         TabIndex        =   31
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Maximum of 4 Letters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -73920
         TabIndex        =   30
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Enter the Word for Level 3 (Difficult):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   26
         Top             =   4080
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "Enter the Word for Level 2 (Average):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   25
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Enter the Word for Level 1 (Easy):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Enter the Definition of the Word:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Enter the Word to be Defined:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Enter New Password:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   12
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Old Password:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbSelection_Click()
If cmbSelection.Text = "Guessing Game" Then
lblWord.Visible = True
lblDefinition.Visible = True
txtWord.Visible = True
txtDefinition.Visible = True
lblLevel1.Visible = False
lblLevel2.Visible = False
lblLevel3.Visible = False
txtLevel1.Visible = False
txtLevel2.Visible = False
txtLevel3.Visible = False
lbl4.Visible = False
lbl7.Visible = False
lbl10.Visible = False
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False

Else
lblWord.Visible = False
lblDefinition.Visible = False
txtWord.Visible = False
txtDefinition.Visible = False
lblLevel1.Visible = True
lblLevel2.Visible = True
lblLevel3.Visible = True
txtLevel1.Visible = True
txtLevel2.Visible = True
txtLevel3.Visible = True
lbl4.Visible = True
lbl7.Visible = True
lbl10.Visible = True
lbl1.Visible = True
lbl2.Visible = True
lbl3.Visible = True


End If

'Vowels and Consonants Game
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


