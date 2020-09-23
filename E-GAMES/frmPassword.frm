VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdPCancel 
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
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdPOK 
      BackColor       =   &H00FF8080&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblpw 
      BackColor       =   &H00400000&
      Caption         =   "JESSIE"
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   8520
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPassword.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPCancel_Click()
frmGames.Enabled = True
Unload Me
End Sub

Private Sub cmdPOK_Click()
 If txtPassword.Text = lblpw.Caption Then
   frmConfiguration.Show
   Unload Me
 Else
   MsgBox "Invalid Password", vbCritical
   txtPassword.Text = ""
   txtPassword.SetFocus

 End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
    cmdPOK_Click
   End If
End Sub
