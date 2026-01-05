Attribute VB_Name = "Geral"
Option Explicit

'eNDRICK MUDANDO

'Strings da conexão
Public cn As ADODB.Connection
Public strBD As String

'Strings do usuário
Global gintUsuarioID As Integer
Global gstrNome As String
Global gstrDepto As String
Global gstrEMail As String
Global gstrSenha As String

Global vIDusuarioReal As String



'Strings de OS
Global gintOSID As Integer

'Variáveis para criação de usuário
Global gintNovoUsuarioID As Integer

'Variáveis do frmChamados
'Combo status
Global gstrcboStatus As String
Global gblnTelaChamados As Boolean


Public Sub ConectarBD()
    'Conexão com a Telecom
    strBD = "PROVIDER=SQLOLEDB;SERVER=196.200.80.20;DATABASE=HelpDesk;UID=cablena_user;PWD=C@bl3na;"
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open (strBD)
End Sub

Public Function CodDec(ByVal vText As String) As String
On Error Resume Next
Dim strAux As String
Dim intCount As Long
Dim strText As String
    
    strText = vText
    strAux = ""
    For intCount = 1 To Len(strText)
        strAux = Chr((Asc(Mid(strText, intCount, 1)) Xor 255)) & strAux
    Next
    CodDec = strAux
End Function


