Attribute VB_Name = "Module1"
  Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
  Public Const HelpC = &H3&
  
  Public abc
  Public Path As String
  Public cnn As ADODB.Connection
  
  Public adoword As ADODB.Recordset
  Public adoguess As ADODB.Recordset
  
  Public adoguessgame As ADODB.Recordset
  Public adovowelgame As ADODB.Recordset
 
  Public adovalidate As ADODB.Recordset
  Public adoval As ADODB.Recordset
  
  Public adoresults As ADODB.Recordset
  
  Public Sub main()
  
    If Right(App.Path, 1) <> "\" Then
      Path = App.Path & "\"
    
    Else
      Path = App.Path
    
    End If
    
    App.HelpFile = Path & "Jesz.HLP"
    Load frmGames
    frmGames.Show
                                                                                                                                                                                                                 frmGames.Caption = frmConfiguration.rs_filldatabase.Caption
  End Sub

