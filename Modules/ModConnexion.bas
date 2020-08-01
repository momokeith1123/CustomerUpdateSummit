Attribute VB_Name = "ModConnexion"
Option Explicit

'===================================================================
' AUTHOR : Mamadou Oumar Keita - AFDB Apr 2020
' FUNCTION : InitConnection(...)
' DESCRIPTION : Initiliase la connexion à la base de données
' PARAMS : * DSN : Nom du DSN associé à la connexion
' * UserName : Nom de l'utilisateur
' * Password : Mot de passe de l'utilisateur
' VERSION : 1.1
'===================================================================
Public ADOcnx          As ADODB.Connection
Public Function InitConnection(DSN As String, UserName As String, PassWord As String) As Boolean
  
  Dim query           As String
  Dim cnxString       As String
  Dim RequeteOk       As Boolean
  
  Dim mRst As New ADODB.Recordset
  Set ADOcnx = New ADODB.Connection

  InitConnection = False
  'Initialisation de la chaine de connexion
  ADOcnx.ConnectionString = "DSN=" & DSN & ";"

  'Vérifie que la connexion est bien fermée
  If ADOcnx.State = adStateOpen Then
    ADOcnx.Close
  End If
  On Error GoTo BadConnection
  'Connexion à la base de données
  ADOcnx.Open cnxString, UserName, PassWord, adAsyncConnect
  'Attente que la connexion soit établie
  While (ADOcnx.State = adStateConnecting)
    DoEvents
   Wend
  'Vérification des erreurs dans le cas d'une mauvaise connexion
  If ADOcnx.Errors.Count > 0 Then
    'Affichage des erreurs
    MsgBox ADOcnx.Errors.Item(0)
    InitConnection = False
    Exit Function
  Else
    InitConnection = True
   End If
   Exit Function

BadConnection:
If ADOcnx.Errors.Count > 0 Then
    'Affichage des erreurs
    MsgBox ADOcnx.Errors.Item(0)
    InitConnection = False
    Exit Function
Else
    MsgBox err.Description
End If
End Function

'============================================================================='
Public Function ExecSQL(ByRef query As String, ByRef rst As ADODB.Recordset, ByRef cnx As ADODB.Connection) As Boolean
    
  Dim cmd As New ADODB.Command
  
  'cmd.CommandType = adCmdText
  'cmd.CommandText = "alter session set current_schema = V56;"
  'cmd.Execute
  
  'Initialisation du RecordSet
  If rst.State <> adStateClosed Then rst.Close
'  If rstparam.State <> adStateClosed Then rstparam.Close

  'Ouvre une transaction pour ne pas à avoir à réaliser de commit en fin de traitement
  ADOcnx.BeginTrans

  'Positionne le curseur côté client
  rst.CursorLocation = adUseClient
  'Vérifie que la connexion passée est bonne
  Set rst.ActiveConnection = cnx
  
  On Error GoTo ErrHandle
  'Exécute la requête
  'cmd.Execute
  'rstparam.Open "Alter session set current_schema = V56", ADOcnx
  rst.Open query, ADOcnx

  'Valide la transaction
  ADOcnx.CommitTrans
  ExecSQL = True
  Exit Function

ErrHandle:
  ExecSQL = False
  MsgBox "ADOManager.ExecSQL:ErrHandle" & vbCr & vbCr & err.Description, vbCritical
End Function

