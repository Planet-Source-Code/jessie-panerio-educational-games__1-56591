VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form frmConfiguration 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   794
      BackColor       =   -2147483638
      MouseIcon       =   "frmconfiguration.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Guessing Game"
      TabPicture(0)   =   "frmconfiguration.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rs"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PicFind1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Vowels Consonant Game"
      TabPicture(1)   =   "frmconfiguration.frx":0BEE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PicFind2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Author"
      TabPicture(2)   =   "frmconfiguration.frx":17C0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label23"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ShockwaveFlashA"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "ShockwaveFlashE"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ShockwaveFlashF"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Help"
      TabPicture(3)   =   "frmconfiguration.frx":2392
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "rs_filldatabase"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C000&
         Height          =   6375
         Left            =   -75000
         ScaleHeight     =   6315
         ScaleWidth      =   7755
         TabIndex        =   84
         Top             =   480
         Width           =   7815
         Begin VB.CommandButton cmdCloseHelp 
            BackColor       =   &H00FF8080&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C000&
            Caption         =   "Note: Always click the validation button or press Enter after typing the word to avoid duplication of data. "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   4920
            Width           =   7695
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   7560
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   7560
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C000&
            Caption         =   "Note: If questions are keep repeating, you should enter or add  more data to avoid repeatition of question."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   4560
            Width           =   7695
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Add, Edit, Delete Records in the Database"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1320
            TabIndex        =   95
            Top             =   2640
            Width           =   4815
         End
         Begin VB.Image Image3 
            Height          =   345
            Left            =   840
            Picture         =   "frmconfiguration.frx":26AC
            Stretch         =   -1  'True
            Top             =   3720
            Width           =   360
         End
         Begin VB.Image Image2 
            Height          =   375
            Left            =   840
            Picture         =   "frmconfiguration.frx":3576
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   840
            Picture         =   "frmconfiguration.frx":4138
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   345
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Print Records in the Database"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1320
            TabIndex        =   94
            Top             =   3840
            Width           =   3495
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Search for a Records in the Database"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1320
            TabIndex        =   93
            Top             =   3240
            Width           =   4335
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Add, Edit, Delete Records in the Database"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   -6360
            TabIndex        =   92
            Top             =   2880
            Width           =   5775
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "In this module you can :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   135
            Left            =   720
            TabIndex        =   88
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "In this module you can :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   720
            TabIndex        =   91
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "What does it do?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Left            =   480
            TabIndex        =   90
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "What does it do?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   480
            TabIndex        =   87
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Configuration Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1605
            TabIndex        =   86
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Configuration Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   1605
            TabIndex        =   89
            Top             =   360
            Width           =   4575
         End
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlashF 
         Height          =   615
         Left            =   -75120
         TabIndex        =   83
         Top             =   6120
         Width           =   1455
         _cx             =   2566
         _cy             =   1085
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
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlashE 
         Height          =   615
         Left            =   -74760
         TabIndex        =   82
         Top             =   6120
         Width           =   7575
         _cx             =   13361
         _cy             =   1085
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
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlashA 
         Height          =   5655
         Left            =   -75000
         TabIndex        =   81
         Top             =   480
         Width           =   7815
         _cx             =   13785
         _cy             =   9975
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
         BGColor         =   "000000"
         SWRemote        =   ""
         MovieData       =   ""
      End
      Begin VB.PictureBox PicFind2 
         BackColor       =   &H00C0C000&
         Height          =   6855
         Left            =   -75000
         ScaleHeight     =   6795
         ScaleWidth      =   7755
         TabIndex        =   1
         Top             =   -6840
         Visible         =   0   'False
         Width           =   7815
         Begin VB.CommandButton cmdPrintCVC 
            BackColor       =   &H00FF8080&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   5400
            Width           =   855
         End
         Begin VB.CommandButton cmdLevel1Find 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   3720
            Picture         =   "frmconfiguration.frx":4CFA
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Validation Button"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtVCFRecordNumber 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "Text7"
            Top             =   4680
            Width           =   4335
         End
         Begin VB.CommandButton cmdVCMoveNext 
            BackColor       =   &H00C0FFC0&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton cmdVCMoveFirst 
            BackColor       =   &H00C0FFC0&
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton cmdVCMovePrevious 
            BackColor       =   &H00C0FFC0&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton cmdVCMoveLast 
            BackColor       =   &H00C0FFC0&
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton cmdVCFCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2018
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   5400
            Width           =   855
         End
         Begin VB.CommandButton cmdVCFSave 
            BackColor       =   &H00FF8080&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2978
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5400
            Width           =   855
         End
         Begin VB.CommandButton cmdVCFDelete 
            BackColor       =   &H00FF8080&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3938
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   5400
            Width           =   855
         End
         Begin VB.CommandButton cmdVCFEdit 
            BackColor       =   &H00FF8080&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   5400
            Width           =   855
         End
         Begin VB.CommandButton cmdVCFClose 
            BackColor       =   &H00FF8080&
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
            Height          =   735
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   5400
            Width           =   855
         End
         Begin VB.TextBox txtVCFLevel3 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   12
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtVCFLevel2 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   11
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CommandButton cmdLevel2Find 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   3720
            Picture         =   "frmconfiguration.frx":749C
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Validation Button"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdLevel3Find 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   3720
            Picture         =   "frmconfiguration.frx":9C3E
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Validation Button"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtVCFLevel1 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   8
            Top             =   480
            Width           =   2295
         End
         Begin VB.OptionButton optLevel1 
            BackColor       =   &H00C0C000&
            Caption         =   " Word for Level 1 (Easy)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3015
         End
         Begin VB.OptionButton optLevel2 
            BackColor       =   &H00C0C000&
            Caption         =   " Word for Level 2 (Average)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   3495
         End
         Begin VB.OptionButton optLevel3 
            BackColor       =   &H00C0C000&
            Caption         =   " Word for Level 3 (Difficult)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   2040
            Width           =   3375
         End
         Begin VB.TextBox txtEasy 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtAverage 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   3
            Top             =   1440
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtDifficult 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   2
            Top             =   2280
            Visible         =   0   'False
            Width           =   2295
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            CausesValidation=   0   'False
            Height          =   1455
            Left            =   240
            TabIndex        =   24
            Top             =   2880
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2566
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777088
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
         Begin VB.Label lbl1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Enter a New value for Level 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4800
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Enter a New value for Level 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4800
            TabIndex        =   26
            Top             =   1200
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Enter a New value for Level 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4800
            TabIndex        =   25
            Top             =   2040
            Visible         =   0   'False
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C000&
         Height          =   6135
         Left            =   -75000
         ScaleHeight     =   6075
         ScaleWidth      =   7755
         TabIndex        =   60
         Top             =   480
         Width           =   7815
         Begin VB.CommandButton cmd3 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   5400
            Picture         =   "frmconfiguration.frx":C3E0
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Validation Button"
            Top             =   3600
            Width           =   495
         End
         Begin VB.CommandButton cmd2 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   5400
            Picture         =   "frmconfiguration.frx":EB82
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Validation Button"
            Top             =   2160
            Width           =   495
         End
         Begin VB.CommandButton cmd1 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   5400
            Picture         =   "frmconfiguration.frx":11324
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Validation Button"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtVCLevel3 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   68
            Top             =   3600
            Width           =   3855
         End
         Begin VB.TextBox txtVCLevel2 
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
            Left            =   1200
            MaxLength       =   7
            TabIndex        =   67
            Top             =   2160
            Width           =   3855
         End
         Begin VB.TextBox txtVCLevel1 
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
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   66
            Top             =   720
            Width           =   3855
         End
         Begin VB.CommandButton cmdVClose 
            BackColor       =   &H00FF8080&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5558
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdVFind 
            BackColor       =   &H00FF8080&
            Caption         =   "&Find"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4478
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdVSave 
            BackColor       =   &H00FF8080&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3398
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdVCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2318
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdVNew 
            BackColor       =   &H00FF8080&
            Caption         =   "&New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1238
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   4920
            Width           =   975
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
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
            Left            =   1200
            TabIndex        =   74
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
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
            Left            =   1200
            TabIndex        =   73
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
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
            Left            =   1200
            TabIndex        =   72
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C000&
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1200
            TabIndex        =   71
            Top             =   3360
            Width           =   3975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C000&
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1200
            TabIndex        =   70
            Top             =   1920
            Width           =   3975
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C000&
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1200
            TabIndex        =   69
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.PictureBox PicFind1 
         BackColor       =   &H00C0C000&
         Height          =   7095
         Left            =   0
         ScaleHeight     =   7035
         ScaleWidth      =   7755
         TabIndex        =   28
         Top             =   -7080
         Visible         =   0   'False
         Width           =   7815
         Begin VB.CommandButton cmdPrintCG 
            BackColor       =   &H00FF8080&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4916
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   5640
            Width           =   855
         End
         Begin VB.CommandButton cmdCClose 
            BackColor       =   &H00FF8080&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   5640
            Width           =   855
         End
         Begin VB.CommandButton cmdCEdit 
            BackColor       =   &H00FF8080&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1054
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   5640
            Width           =   855
         End
         Begin VB.CommandButton cmdCDelete 
            BackColor       =   &H00FF8080&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3934
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   5640
            Width           =   855
         End
         Begin VB.CommandButton cmdCSave 
            BackColor       =   &H00FF8080&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2974
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   5640
            Width           =   855
         End
         Begin VB.CommandButton cmdCCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2014
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   5640
            Width           =   855
         End
         Begin VB.TextBox txtFDefinitionC 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   1920
            Width           =   6975
         End
         Begin VB.TextBox txtFWordC 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1080
            Width           =   6975
         End
         Begin VB.TextBox txtFWord 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   36
            Top             =   120
            Width           =   4575
         End
         Begin VB.CommandButton cmdLast 
            BackColor       =   &H00C0FFC0&
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
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   4800
            Width           =   735
         End
         Begin VB.CommandButton cmdPrevious 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   34
            Top             =   4800
            Width           =   735
         End
         Begin VB.CommandButton cmdFirst 
            BackColor       =   &H00C0FFC0&
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
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   4800
            Width           =   735
         End
         Begin VB.CommandButton cmdNext 
            BackColor       =   &H00C0FFC0&
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
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox txtFindRNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   4800
            Width           =   4095
         End
         Begin VB.CommandButton cmdFindButton 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   6840
            Picture         =   "frmconfiguration.frx":13AC6
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Validation Button"
            Top             =   120
            Width           =   495
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1335
            Left            =   360
            TabIndex        =   29
            Top             =   3240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2355
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777088
            DefColWidth     =   267
            HeadLines       =   1
            RowHeight       =   17
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
               Name            =   "Arial"
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
               SizeMode        =   1
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl91 
            BackColor       =   &H00C0C000&
            Caption         =   "(You can now edit the definition below)"
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
            Left            =   1560
            TabIndex        =   59
            Top             =   1680
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label lbl90 
            BackColor       =   &H00C0C000&
            Caption         =   "(You can now edit the word below)"
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
            Left            =   1080
            TabIndex        =   58
            Top             =   840
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            Caption         =   "Definition "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C000&
            Caption         =   "Word "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C000&
            Caption         =   "Enter the Word "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C000&
         Height          =   6255
         Left            =   0
         ScaleHeight     =   6195
         ScaleWidth      =   7755
         TabIndex        =   47
         Top             =   480
         Width           =   7815
         Begin VB.CommandButton cmdFindWord 
            BackColor       =   &H00FF8080&
            Height          =   495
            Left            =   7080
            Picture         =   "frmconfiguration.frx":16268
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Validation Button"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtCDefinition 
            BackColor       =   &H00FFFF80&
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
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Top             =   1680
            Width           =   7455
         End
         Begin VB.TextBox txtCWord 
            BackColor       =   &H00FFFF80&
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
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   53
            Top             =   480
            Width           =   6855
         End
         Begin VB.CommandButton cmdGClose 
            BackColor       =   &H00FF8080&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdGFind 
            BackColor       =   &H00FF8080&
            Caption         =   "&Find"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdGSave 
            BackColor       =   &H00FF8080&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdGCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   4920
            Width           =   975
         End
         Begin VB.CommandButton cmdGNew 
            BackColor       =   &H00FF8080&
            Caption         =   "&New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   4920
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C000&
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
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
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
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "(maximum of 9 letters)"
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
            Left            =   3480
            TabIndex        =   55
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Label Label23 
         Caption         =   $"frmconfiguration.frx":18A0A
         Height          =   375
         Left            =   -73920
         TabIndex        =   98
         Top             =   3720
         Width           =   4815
      End
      Begin VB.Label rs 
         Caption         =   $"frmconfiguration.frx":18A98
         Height          =   375
         Left            =   -4680
         TabIndex        =   97
         Top             =   4800
         Width           =   4695
      End
      Begin VB.Label rs_filldatabase 
         Caption         =   $"frmconfiguration.frx":18B22
         Height          =   375
         Left            =   -79320
         TabIndex        =   96
         Top             =   6000
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'

Private Sub display()

  If adoguessgame.BOF = True Or adoguessgame.EOF = True Then
    Exit Sub
  Else
    txtFWordC.Text = adoguessgame.Fields("word")
    txtFDefinitionC.Text = adoguessgame.Fields("definition")
    txtFindRNo.Text = "Record " & adoguessgame.AbsolutePosition & " of " & adoguessgame.RecordCount
  End If
  
End Sub

Private Sub cmdCloseHelp_Click()
    
    Unload Me
    frmGames.Enabled = True
    frmGames.Show
    
End Sub

Private Sub cmdPrintCG_Click()

  With adoguessgame
     Set DataReport3.DataSource = adoguessgame
     DataReport3.Show
  End With
  
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdCCancel_Click()

     cmdFindButton.Enabled = True
     lbl90.Visible = False
     lbl91.Visible = False
     txtFWord.Locked = False
     txtFWordC.Locked = True
     txtFDefinitionC.Locked = True
     cmdCEdit.Enabled = True
     cmdCCancel.Enabled = False
     cmdCSave.Enabled = False
     cmdCDelete.Enabled = True
     cmdCClose.Enabled = True
     
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdCClose_Click()
     
     txtCWord.Text = ""
     txtCDefinition.Text = ""
   
     Do While PicFind1.Top > 1 - PicFind1.Height
       PicFind1.Top = PicFind1.Top - 1
       DoEvents
     Loop
       PicFind1.Visible = False
   
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdCDelete_Click()

   Dim res
   lbl90.Visible = False
   lbl91.Visible = False
   txtFWordC.Locked = True
   txtFWord.Locked = False
   txtFDefinitionC.Locked = True
     If adoguessgame.BOF = True Or adoguessgame.EOF = True Then
       Exit Sub
     ElseIf adoguessgame.BOF = True And adoguessgame.EOF = True Then
       MsgBox "Empty Database", vbInformation, "Guessing Game"
       Exit Sub
     Else
       res = MsgBox("Are you sure you want to Delete  " & txtFWordC.Text, vbYesNo, "Confirmation")
         If res = vbYes Then
           adoguessgame.Delete
              If adoguessgame.EOF = True Then
                 adoguessgame.MovePrevious
                 display
              Else
                 adoguessgame.MoveNext
              End If
         Else
           Exit Sub
         End If
     End If
      
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdCEdit_Click()
  
  On Error Resume Next
  If adoguessgame.BOF = True And adoguessgame.EOF = True Then
    cmdCEdit.Enabled = False
    txtFWordC.Locked = True
    txtFDefinitionC.Locked = True
    txtFWordC.Text = ""
    txtFDefinitionC.Text = ""
    Exit Sub
  Else
    cmdFindButton.Enabled = False
    lbl90.Visible = True
    lbl91.Visible = True
    txtFWord.Locked = True
    txtFWordC.Locked = False
    txtFDefinitionC.Locked = False
      With adoguessgame
        txtFWordC.Text = !word & vbNullString
        txtFDefinitionC.Text = !definition & vbNullString
      End With
       cmdCEdit.Enabled = False
       cmdCCancel.Enabled = True
       cmdCSave.Enabled = True
       cmdCDelete.Enabled = False
       cmdCClose.Enabled = False
  End If
     
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdCSave_Click()
  
  On Error Resume Next
  lbl90.Visible = False
  lbl91.Visible = False
  txtFWord.Locked = False
  txtFWordC.Locked = True
  txtFDefinitionC.Locked = True
  cmdFindButton.Enabled = True
    With adoguessgame
     !word = txtFWordC.Text
     !definition = txtFDefinitionC.Text
    End With
    
    adoguessgame.UpdateBatch adAffectCurrent
    Set DataGrid1.DataSource = adoguessgame
    
    cmdCEdit.Enabled = True
    cmdCCancel.Enabled = False
    cmdCSave.Enabled = False
    cmdCDelete.Enabled = True
    cmdCClose.Enabled = True

End Sub

Private Sub cmdFindButton_Click()
'On Error Resume Next
  
  adoguessgame.Requery
  adoguessgame.Find "word = " & "'" & frmConfiguration.txtFWord.Text & "'", , adSearchForward, 1
  
  If adoguessgame.EOF Then
      MsgBox "The Word you've Entered does not Exist!", vbInformation, "Guessing Game"
      adoguessgame.MoveFirst
      txtFWord.Text = ""
  Else
      display
  End If
End Sub

Private Sub cmdFindWord_Click()
 On Error Resume Next
  
  adovalidate.Requery
  adovalidate.Find "word = " & "'" & frmConfiguration.txtCWord.Text & "'", , adSearchForward, 1
  
  If adovalidate.EOF Then
      txtCDefinition.Locked = False
      txtCDefinition.SetFocus
      adovalidate.MoveFirst
        
  Else
      MsgBox "Sorry! The Word you've Entered already exist!", vbInformation, "Guessing Game"
      txtCWord.Text = ""
  End If
 
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdFirst_Click()
  If adoguessgame.BOF = True Or adoguessgame.EOF = True Then
    Exit Sub
  Else
    adoguessgame.MoveFirst
    txtFWord.Text = ""
    display
  End If
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdGClose_Click()
     
  Unload Me
  frmGames.Enabled = True
  frmGames.Show
                                                                                                                                                                                                                 frmGames.Caption = frmConfiguration.rs.Caption
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdGFind_Click()
  
    Set DataGrid1.DataSource = adoguessgame
    
    display
    cmdCEdit.Enabled = True
    cmdCCancel.Enabled = False
    cmdCSave.Enabled = False
    cmdCDelete.Enabled = True
    cmdCClose.Enabled = True
   
   Do While PicFind1.Top < 0
    PicFind1.Top = PicFind1.Top + 1
      DoEvents
      PicFind1.Visible = True
   Loop
   
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdGNew_Click()
   
   adoguessgame.AddNew
   LoadDataInControls
   
   cmdGNew.Enabled = False
   cmdGCancel.Enabled = True
   cmdGClose.Enabled = False
   cmdGFind.Enabled = False
   cmdGSave.Enabled = True
   txtCWord.SetFocus
   txtCWord.Locked = False
   txtCDefinition.Locked = True
   
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub LoadDataInControls()
    If adoguessgame.BOF = True Or adoguessgame.EOF = True Then
        Exit Sub
    End If
    
    txtCWord.Text = adoguessgame!word & ""
    txtCDefinition.Text = adoguessgame!definition & ""
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdGSave_Click()
           
         
     Dim res As VbMsgBoxResult
      
     res = MsgBox("Save Data?", vbYesNo, "Confirmation")
      If res = vbYes Then
      
         If txtCWord.Text = "" Or txtCDefinition.Text = "" Then
           MsgBox "Missing Data", vbInformation, "Warning"
           txtCWord.SetFocus
           cmdGSave.Enabled = True
           
         Else
                            
           WriteDataFromControls
           adoguessgame.Update
           adoguessgame.Requery
           cmdGNew.Enabled = True
           cmdGCancel.Enabled = False
           cmdGClose.Enabled = True
           cmdGFind.Enabled = True
           cmdGSave.Enabled = False
           txtCWord.Locked = True
           txtCDefinition.Locked = True
           
                   
         End If
         
      Else
         adoguessgame.CancelUpdate
         cmdGNew.Enabled = True
         cmdGCancel.Enabled = False
         cmdGClose.Enabled = True
         cmdGFind.Enabled = True
         cmdGSave.Enabled = False
         txtCWord.Text = ""
         txtCDefinition.Text = ""
         txtCWord.Locked = True
         txtCDefinition.Locked = True
      End If
        
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub WriteDataFromControls()
    
    On Error Resume Next
    adoguessgame!word = txtCWord.Text
    adoguessgame!definition = txtCDefinition.Text
    
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdGCancel_Click()
     
     adoguessgame.CancelUpdate
     adoguessgame.Requery
     cmdGNew.Enabled = True
     cmdGFind.Enabled = True
     cmdGClose.Enabled = True
     cmdGCancel.Enabled = False
     cmdGSave.Enabled = False
     txtCWord.Locked = True
     txtCDefinition.Locked = True
     txtCWord.Text = ""
     txtCDefinition.Text = ""

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdGDelete_Click()
 
 With adoguessgame
   If .BOF And .EOF = True Then
    MsgBox "Empty Database!", vbInformation, "Guessing Game"
   End If
    .Delete
    .Requery
   End If
  End With

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdLast_Click()
  
  If adoguessgame.BOF = True Or adoguessgame.EOF = True Then
    Exit Sub
  Else
    adoguessgame.MoveLast
    txtFWord.Text = ""
    display
  End If

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdNext_Click()
    
    adoguessgame.MoveNext
    display
    txtFWord.Text = ""
  If adoguessgame.EOF = True Then
   adoguessgame.MovePrevious
   display
  ElseIf adoguessgame.EOF = True And adoguessgame.BOF = True Then
   MsgBox "Empty Database", vbInformation, "Guessing Game"
   Exit Sub
  End If
  
End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdPrevious_Click()
    
    adoguessgame.MovePrevious
    display
    txtFWord.Text = ""
  If adoguessgame.BOF = True Then
    adoguessgame.MoveNext
    display
  ElseIf adoguessgame.EOF = True And adoguessgame.BOF = True Then
   MsgBox "Empty Database", vbInformation, "Guessing Game"
   Exit Sub
  End If

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdVCClose_Click()
  
  PicFind2.Visible = False

End Sub

Private Sub cmdPrintCVC_Click()
  
  With adovowelgame
    Set DataReport4.DataSource = adovowelgame
    DataReport4.Show
  End With

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub cmdVClose_Click()
  
  Unload Me
  frmGames.Enabled = True
  frmGames.Show

End Sub




Private Sub DataGrid1_Click()
  
  display

End Sub

'----------------------------'
'Configuration(Guessing Game)'
'----------------------------'
Private Sub Form_Load()

  Set adoguessgame = New ADODB.Recordset
  adoguessgame.Open "Select * from guessinggame order by word ASC", cnn, adOpenStatic, adLockPessimistic
  
  Set adovowelgame = New ADODB.Recordset
  adovowelgame.Open "Select * from vowelsconsonants order by level1 ASC", cnn, adOpenStatic, adLockPessimistic
  
  Set adovalidate = New ADODB.Recordset
  adovalidate.Open "Select * from guessinggame", cnn, adOpenStatic, adLockPessimistic

  Set adoval = New ADODB.Recordset
  adoval.Open "Select * from vowelsconsonants", cnn, adOpenStatic, adLockPessimistic
   
  SSTab1.Tab = 3
  
End Sub

Private Sub optLevel1_Click()
  
  If optLevel1.Value = True Then
    txtVCFLevel1.Enabled = True
    cmdLevel1Find.Enabled = True
    'txtVCFLevel1.SetFocus
    txtVCFLevel1.Locked = False
    txtVCFLevel2.Enabled = False
    cmdLevel2Find.Enabled = False
    txtVCFLevel3.Enabled = False
    cmdLevel3Find.Enabled = False
  End If

End Sub

Private Sub optLevel2_Click()
 
 If optLevel2.Value = True Then
    txtVCFLevel2.Enabled = True
    cmdLevel2Find.Enabled = True
    txtVCFLevel2.SetFocus
    txtVCFLevel2.Locked = False
    txtVCFLevel1.Enabled = False
    cmdLevel1Find.Enabled = False
    txtVCFLevel3.Enabled = False
    cmdLevel3Find.Enabled = False
  End If

End Sub

Private Sub optLevel3_Click()
  
  If optLevel3.Value = True Then
    txtVCFLevel3.Enabled = True
    cmdLevel3Find.Enabled = True
    txtVCFLevel3.SetFocus
    txtVCFLevel3.Locked = False
    txtVCFLevel1.Enabled = False
    cmdLevel1Find.Enabled = False
    txtVCFLevel2.Enabled = False
    cmdLevel2Find.Enabled = False
  End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  
  If SSTab1.Tab = 0 Then
    On Error Resume Next
    ShockwaveFlashA.Movie = Path & ""
    ShockwaveFlashE.Movie = Path & ""
    ShockwaveFlashF.Movie = Path & ""
    
        
    If (cmdVNew.Enabled = False) Then
    MsgBox "Cancel Entry of Data", vbInformation, "Vowel-Consonant Game"
    adovowelgame.CancelUpdate
    cmdGNew.Enabled = True
    End If
    
    
    cmdGNew.Enabled = True
    cmdGFind.Enabled = True
    cmdGClose.Enabled = True
    cmdGCancel.Enabled = False
    cmdGSave.Enabled = False
       
    txtCWord.Locked = True
    txtCDefinition.Locked = True
    txtCWord.Text = ""
    txtCDefinition.Text = ""
         
               
             
  ElseIf SSTab1.Tab = 1 Then
    On Error Resume Next
    ShockwaveFlashA.Movie = Path & ""
    ShockwaveFlashE.Movie = Path & ""
    ShockwaveFlashF.Movie = Path & ""
    
    
    If (cmdGNew.Enabled = False) Then
    MsgBox "Cancel Entry of Data", vbInformation, "Guessing Game"
    adoguessgame.CancelUpdate
    cmdGNew.Enabled = True
    End If
    
    cmdVNew.Enabled = True
    cmdVFind.Enabled = True
    cmdVClose.Enabled = True
    cmdVCancel.Enabled = False
    cmdVSave.Enabled = False
  
    txtVCLevel1.Locked = True
    txtVCLevel2.Locked = True
    txtVCLevel3.Locked = True
    
  ElseIf SSTab1.Tab = 2 Then
    
    ShockwaveFlashA.Movie = Path & "binary.swf"
    ShockwaveFlashA.Play
  
    ShockwaveFlashE.Movie = Path & "jesz-e.swf"
    ShockwaveFlashE.Play
  
    ShockwaveFlashF.Movie = Path & "flag.swf"
    ShockwaveFlashF.Play
    
    If (cmdGNew.Enabled = False) Then
    MsgBox "Cancel Entry of Data", vbInformation, "Guessing Game"
    adoguessgame.CancelUpdate
    cmdGNew.Enabled = True
    ElseIf (cmdVNew.Enabled = False) Then
    MsgBox "Cancel Entry of Data", vbInformation, "Vowel-Consonant Game"
    adovowelgame.CancelUpdate
    cmdGNew.Enabled = True
    End If
    
  ElseIf SSTab1.Tab = 3 Then
    On Error Resume Next
    ShockwaveFlashA.Movie = Path & ""
    ShockwaveFlashE.Movie = Path & ""
    ShockwaveFlashF.Movie = Path & ""
    
    If (cmdGNew.Enabled = False) Then
    MsgBox "Cancel Entry of Data", vbInformation, "Guessing Game"
    adoguessgame.CancelUpdate
    cmdGNew.Enabled = True
    ElseIf (cmdVNew.Enabled = False) Then
    MsgBox "Cancel Entry of Data", vbInformation, "Vowel-Consonant Game"
    adovowelgame.CancelUpdate
    cmdGNew.Enabled = True
    End If
  
  End If

End Sub


Private Sub txtAverage_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
      If KeyAscii = 13 Then
        cmd1_Click
      End If
    If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub txtCWord_KeyPress(KeyAscii As Integer)
  
  'validation -capital letters only
    KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
   If KeyAscii = 13 Then
      
   cmdFindWord_Click
   End If
 
   
   If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
  
End Sub

Private Sub txtDifficult_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
      If KeyAscii = 13 Then
        cmd1_Click
      End If
    If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

Private Sub txtEasy_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
      If KeyAscii = 13 Then
        cmd1_Click
      End If
    If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub txtFWord_KeyPress(KeyAscii As Integer)
    
     KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
  
   If KeyAscii = 13 Then
     cmdFindButton_Click
   End If

     If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
 
End Sub
'----------------------------'
'Configuration(Guessing Game)Find1
'----------------------------'
Private Sub txtFWordC_KeyPress(KeyAscii As Integer)
   
   KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
   If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
  
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub txtVCLevel1_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
      If KeyAscii = 13 Then
        cmd1_Click
      End If
    If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub validatevc1()
   
   On Error Resume Next
   adoval.Requery
    adoval.Find "level1 = " & "'" & abc & "'", , adSearchForward, 1
  If adoval.EOF Then
       
       adoval.Find "level2 = " & "'" & abc & "'", , adSearchForward, 1
          If adoval.EOF Then
               
            adoval.Find "level3 = " & "'" & abc & "'", , adSearchForward, 1
                 If adoval.EOF Then
                   adoval.MoveFirst
                
                 Else
                   MsgBox "Sorry! The word you've entered already exist.", vbInformation, "Vowel-Consonant Game"
                   abc = ""
                   Exit Sub
                 End If
          Else
            MsgBox "Sorry! The word you've entered already exist.", vbInformation, "Vowel-Consonant Game"
            abc = ""
            Exit Sub
          End If
    
    adoval.MoveFirst
  
  Else
    MsgBox "Sorry! The word you've entered already exist.", vbInformation, "Vowel-Consonant Game"
    abc = ""
    Exit Sub
  End If
  
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub txtVCLevel2_KeyPress(KeyAscii As Integer)
   
   KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
      If KeyAscii = 13 Then
        cmd2_Click
      End If
   If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
 
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub txtVCLevel3_KeyPress(KeyAscii As Integer)
   
   KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
    If KeyAscii = 13 Then
      cmd3_Click
    End If
   If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
 
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub cmdVNew_Click()
   
   adovowelgame.AddNew
   LoadDataInControlsvowel
   
   cmdVNew.Enabled = False
   cmdVCancel.Enabled = True
   cmdVClose.Enabled = False
   cmdVFind.Enabled = False
   cmdVSave.Enabled = True
   txtVCLevel1.SetFocus
   txtVCLevel1.Locked = False
   txtVCLevel2.Locked = False
   txtVCLevel3.Locked = False
   
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub LoadDataInControlsvowel()
    
    If adovowelgame.BOF = True Or adovowelgame.EOF = True Then
        Exit Sub
    End If
    
    txtVCLevel1.Text = adovowelgame!level1 & ""
    txtVCLevel2.Text = adovowelgame!level2 & ""
    txtVCLevel3.Text = adovowelgame!level3 & ""

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub cmdVSave_Click()
            
     Dim res As VbMsgBoxResult
      
     res = MsgBox("Save Data?", vbYesNo, "Confirmation")
      If res = vbYes Then
         
      
         If txtVCLevel1.Text = "" Or txtVCLevel2.Text = "" Or txtVCLevel3.Text = "" Then
           MsgBox "Missing Data", vbInformation, "Warning"
           cmdVSave.Enabled = True
           
         Else
                            
           WriteDataFromControlsvowel
           adovowelgame.Update
           
           cmdVNew.Enabled = True
           cmdVCancel.Enabled = False
           cmdVClose.Enabled = True
           cmdVFind.Enabled = True
           cmdVSave.Enabled = False
           txtVCLevel1.Locked = True
           txtVCLevel2.Locked = True
           txtVCLevel3.Locked = True
             
         End If
         
      Else
         adovowelgame.CancelUpdate
         cmdVNew.Enabled = True
         cmdVCancel.Enabled = False
         cmdVClose.Enabled = True
         cmdVFind.Enabled = True
         cmdVSave.Enabled = False
         txtVCLevel1.Text = ""
         txtVCLevel2.Text = ""
         txtVCLevel3.Text = ""
         txtVCLevel1.Locked = True
         txtVCLevel2.Locked = True
         txtVCLevel3.Locked = True
                          
      End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub WriteDataFromControlsvowel()
    
    On Error Resume Next
    adovowelgame!level1 = txtVCLevel1.Text
    adovowelgame!level2 = txtVCLevel2.Text
    adovowelgame!level3 = txtVCLevel3.Text
    
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub cmdVCancel_Click()
     
     adovowelgame.CancelUpdate
     adovowelgame.Requery
     cmdVNew.Enabled = True
     cmdVFind.Enabled = True
     cmdVClose.Enabled = True
     cmdVCancel.Enabled = False
     cmdVSave.Enabled = False
     txtVCLevel1.Locked = True
     txtVCLevel2.Locked = True
     txtVCLevel3.Locked = True
     txtVCLevel1.Text = ""
     txtVCLevel2.Text = ""
     txtVCLevel3.Text = ""
     
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'
Private Sub cmdVFind_Click()
    
    adovowelgame.Requery
    Set DataGrid2.DataSource = adovowelgame
    displayvc
    
    cmdVCFEdit.Enabled = True
    cmdVCFCancel.Enabled = False
    cmdVCFSave.Enabled = False
    cmdVCFDelete.Enabled = True
    cmdVCFClose.Enabled = True
    
    optLevel1.Value = True
    
    Do While PicFind2.Top < 0
    PicFind2.Top = PicFind2.Top + 1
      DoEvents
      PicFind2.Visible = True
   Loop

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)vowelconsonant
'----------------------------'

Private Sub displayvc()
  
  If adovowelgame.BOF = True Or adovowelgame.EOF = True Then
    txtVCFLevel1.Text = ""
    txtVCFLevel2.Text = ""
    txtVCFLevel3.Text = ""
    txtVCFRecordNumber.Text = "Record " & "0" & " of " & adovowelgame.RecordCount
    Exit Sub
  Else
    txtVCFLevel1.Text = adovowelgame.Fields("level1")
    txtVCFLevel2.Text = adovowelgame.Fields("level2")
    txtVCFLevel3.Text = adovowelgame.Fields("level3")
    txtVCFRecordNumber.Text = "Record " & adovowelgame.AbsolutePosition & " of " & adovowelgame.RecordCount
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCMoveFirst_Click()

  If adovowelgame.BOF = True Or adovowelgame.EOF = True Then
    Exit Sub
  Else
    adovowelgame.MoveFirst
    displayvc
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCMovePrevious_Click()
     
     adovowelgame.MovePrevious
     displayvc
  If adovowelgame.BOF = True Then
   adovowelgame.MoveNext
   displayvc
  ElseIf adovowelgame.EOF = True And adovowelgame.BOF = True Then
   MsgBox "Empty Database", vbInformation, , "Vowel-Consonant Game"
   Exit Sub
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCMoveNext_Click()
   
   adovowelgame.MoveNext
   displayvc
  If adovowelgame.EOF = True Then
   adovowelgame.MovePrevious
   displayvc
  ElseIf adovowelgame.EOF = True And adovowelgame.BOF = True Then
   MsgBox "Empty Database", vbInformation, , "Vowel-Consonant Game"
   Exit Sub
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCMoveLast_Click()
  
  If adovowelgame.BOF = True Or adovowelgame.EOF = True Then
    Exit Sub
  Else
    adovowelgame.MoveLast
    displayvc
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdLevel1Find_Click()

  On Error Resume Next
  adovowelgame.Find "level1 = " & "'" & txtVCFLevel1.Text & "'", , adSearchForward, 1
  
  If adovowelgame.EOF Then
    MsgBox "No match", vbInformation, , "Vowel-Consonant Game"
    adovowelgame.MoveFirst
  End If
    displayvc
    txtVCFLevel1_Change

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdLevel2Find_Click()
  
  On Error Resume Next
  adovowelgame.Find "level2 = " & "'" & txtVCFLevel2.Text & "'", , adSearchForward, 1
  txtVCFLevel1_Change
  If adovowelgame.EOF Then
    MsgBox "No match", vbInformation, "Vowel-Consonant Game"
    adovowelgame.MoveFirst
  End If
    displayvc

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdLevel3Find_Click()

On Error Resume Next
  adovowelgame.Find "level3 = " & "'" & txtVCFLevel3.Text & "'", , adSearchForward, 1
  txtVCFLevel1_Change
  If adovowelgame.EOF Then
    MsgBox "No match", vbInformation, "Vowel-Consonant Game"
    adovowelgame.MoveFirst
  End If
    displayvc

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCFEdit_Click()
  
  On Error Resume Next
  
  If adovowelgame.BOF = True And adovowelgame.EOF = True Then
    cmdVCFEdit.Enabled = False
    txtVCFLevel1.Locked = True
    txtVCFLevel2.Locked = True
    txtVCFLevel3.Locked = True
    txtVCFLevel1.Text = ""
    txtVCFLevel2.Text = ""
    txtVCFLevel3.Text = ""
    Exit Sub
  Else
    txtVCFLevel1_Change
    txtVCFLevel1.Visible = False
    txtVCFLevel2.Visible = False
    txtVCFLevel3.Visible = False
    optLevel1.Visible = False
    optLevel2.Visible = False
    optLevel3.Visible = False
    cmdLevel1Find.Visible = False
    cmdLevel2Find.Visible = False
    cmdLevel3Find.Visible = False
   
    txtEasy.Visible = True
    txtAverage.Visible = True
    txtDifficult.Visible = True
    
    lbl1.Visible = True
    lbl2.Visible = True
    lbl3.Visible = True
      
      With adovowelgame
        txtEasy.Text = !level1 & vbNullString
        txtAverage.Text = !level2 & vbNullString
        txtDifficult.Text = !level3 & vbNullString
      End With
     
       cmdVCFEdit.Enabled = False
       cmdVCFCancel.Enabled = True
       cmdVCFSave.Enabled = True
       cmdVCFDelete.Enabled = False
       cmdVCFClose.Enabled = False
       txtEasy.SetFocus
  End If
     
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCFCancel_Click()
    
    txtVCFLevel1.Locked = True
    txtVCFLevel2.Locked = True
    txtVCFLevel3.Locked = True
    
    txtVCFLevel1.Visible = True
    txtVCFLevel2.Visible = True
    txtVCFLevel3.Visible = True
    
    txtEasy.Visible = False
    txtAverage.Visible = False
    txtDifficult.Visible = False
    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False
    
    optLevel1.Visible = True
    optLevel2.Visible = True
    optLevel3.Visible = True
    cmdLevel1Find.Visible = True
    cmdLevel2Find.Visible = True
    cmdLevel3Find.Visible = True
    
    cmdVCFEdit.Enabled = True
    cmdVCFCancel.Enabled = False
    cmdVCFSave.Enabled = False
    cmdVCFDelete.Enabled = True
    cmdVCFClose.Enabled = True

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCFSave_Click()
  
  On Error Resume Next
    txtVCFLevel1.Locked = True
    txtVCFLevel2.Locked = True
    txtVCFLevel3.Locked = True
    
    txtEasy.Visible = False
    txtAverage.Visible = False
    txtDifficult.Visible = False
    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False
    
    txtVCFLevel1.Visible = True
    txtVCFLevel2.Visible = True
    txtVCFLevel3.Visible = True
    optLevel1.Visible = True
    optLevel2.Visible = True
    optLevel3.Visible = True
    cmdLevel1Find.Visible = True
    cmdLevel2Find.Visible = True
    cmdLevel3Find.Visible = True
     
    With adovowelgame
     !level1 = txtEasy.Text
     !level2 = txtAverage.Text
     !level3 = txtDifficult.Text
    End With
    On Error Resume Next
    adovowelgame.UpdateBatch adAffectCurrent
    Set DataGrid2.DataSource = adovowelgame
    On Error Resume Next
    cmdVCFEdit.Enabled = True
    cmdVCFCancel.Enabled = False
    cmdVCFSave.Enabled = False
    cmdVCFDelete.Enabled = True
    cmdVCFClose.Enabled = True
    displayvc

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCFDelete_Click()

   Dim res
    txtVCFLevel1.Locked = True
    txtVCFLevel2.Locked = True
    txtVCFLevel3.Locked = True
  If adovowelgame.BOF = True Or adovowelgame.EOF = True Then
    Exit Sub
  ElseIf adovowelgame.BOF = True And adovowelgame.EOF = True Then
    MsgBox "Empty Database"
    cmdVCFDelete.Enabled = False
    Exit Sub
  Else
 
    res = MsgBox("Are you sure you want to Delete  " & txtVCFLevel1.Text & ", " & txtVCFLevel2.Text & ", " & txtVCFLevel3.Text, vbYesNo, "Confirmation")
      If res = vbYes Then
         adovowelgame.Delete
         adovowelgame.MoveNext
         displayvc
            If adovowelgame.EOF = True Then
               adovowelgame.MovePrevious
            End If
      Else
         Exit Sub
      End If
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmdVCFClose_Click()
   
   txtVCLevel1.Text = ""
   txtVCLevel2.Text = ""
   txtVCLevel3.Text = ""
    Do While PicFind2.Top > 1 - PicFind2.Height
    PicFind2.Top = PicFind2.Top - 1
      DoEvents
   Loop
   PicFind2.Visible = False
   
End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub txtVCFLevel1_KeyPress(KeyAscii As Integer)
  
  KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
    If KeyAscii = 13 Then
      cmdLevel1Find_Click
      txtVCFLevel1_Change
    End If
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub txtVCFLevel2_KeyPress(KeyAscii As Integer)
  
  KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
    If KeyAscii = 13 Then
      cmdLevel2Find_Click
      txtVCFLevel1_Change
    End If
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub txtVCFLevel3_KeyPress(KeyAscii As Integer)
  
  KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase
    If KeyAscii = 13 Then
      cmdLevel3Find_Click
      txtVCFLevel1_Change
    End If
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub txtVCFLevel1_Change()
   
   txtEasy.Text = txtVCFLevel1.Text
   txtAverage.Text = txtVCFLevel2.Text
   txtDifficult.Text = txtVCFLevel3.Text

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmd1_Click()
  
  If txtVCLevel1.Text = "" Then
    MsgBox "Empty Textbox", vbCritical, "Vowel-Consonant Game"
    txtVCLevel1.SetFocus
  Else
    abc = txtVCLevel1.Text
    validatevc1
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmd2_Click()
  
  If txtVCLevel2.Text = "" Then
    MsgBox "Empty Textbox", vbCritical, "Vowel-Consonant Game"
    txtVCLevel2.SetFocus
  Else
    abc = txtVCLevel2.Text
    validatevc1
  End If

End Sub

'----------------------------'
'Configuration(Vowel Consonant Game)Find2
'----------------------------'
Private Sub cmd3_Click()
  
  If txtVCLevel3.Text = "" Then
    MsgBox "Empty Textbox", vbCritical, "Vowel-Consonant Game"
    txtVCLevel3.SetFocus
  Else
    abc = txtVCLevel3.Text
    validatevc1
  End If

End Sub

