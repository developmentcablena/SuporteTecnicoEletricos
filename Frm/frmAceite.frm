VERSION 5.00
Begin VB.Form frmAceite 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Aceite Usuário"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAceite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNaoValidar 
      Caption         =   "&Não Validar"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      MaxLength       =   5
      TabIndex        =   4
      Top             =   150
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceite 
      Caption         =   "&Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº OS:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   270
      Width           =   705
   End
   Begin VB.Label lblComentario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Comentário/ Motivo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2010
   End
End
Attribute VB_Name = "frmAceite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private strOcorrencia As String
Private strDataOcorrencia As String

Private Sub cmdAceite_Click()
    Call suCadastrarAceite(Trim(txtOSID.Text), Trim(txtComentario.Text))
    Call Unload(Me)
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdNaoValidar_Click()
    If Len(Trim(txtComentario.Text)) > 0 Then
        Call suNaoValidar(Trim(txtOSID.Text), Trim(txtComentario.Text))
        Call Unload(Me)
    Else
        MsgBox "Digite o motivo pela não validação do serviço!", vbOKOnly + vbExclamation, "Suporte Técnico"
    End If
End Sub

Private Sub Form_Load()
    With Me
        .txtOSID.Text = Format(gintOSID, "0000")
    End With
End Sub

Private Sub suCadastrarAceite(ByVal vOSID As Integer, ByVal vComentario As String)
Call ConectarBD
    strSQL = "UPDATE tb_OS SET DataAceite = '" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "',Comentario='" & vComentario & "', Status = 3 " & _
             "WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Sub suNaoValidar(ByVal vOSID As Integer, ByVal vMotivo As String)
Call ConectarBD
Dim strDataOSNaoValidada As String
    
    strDataOSNaoValidada = Format(Now, "yyyy/MM/dd HH:mm:ss")
    
    If fnPesquisarOcorrencia(vOSID) = False Then
        'Call suGravarOcorrencia(vOSID, vMotivo, strDataOSNaoValidada)
        strSQL = "UPDATE tb_OS SET DataOSNaoValidada = '" & strDataOSNaoValidada & "',MotivoOSNaoValidada='" & vMotivo & "', Status = 6 " & _
                 "WHERE OSID = " & vOSID
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        Set rs = Nothing
    Else
        'Call suGravarOcorrencia(vOSID, strOcorrencia, strDataOcorrencia)
        strSQL = "UPDATE tb_OS SET DataOSNaoValidada = '" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "',MotivoOSNaoValidada='" & vMotivo & "', Status = 6 " & _
                 "WHERE OSID = " & vOSID
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        Set rs = Nothing
    End If
End Sub

Private Sub suGravarOcorrencia(ByVal vOSID As Integer, ByVal vMotivo As String, ByVal vDataOcorrencia As String)
Call ConectarBD
    strSQL = "INSERT INTO tb_Ocorrencias (Ocorrencia,DataOcorrencia, Status,OSID) " & _
             "VALUES ('" & vMotivo & "','" & CDate(vDataOcorrencia) & "',1," & vOSID & " )"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Function fnPesquisarOcorrencia(ByVal vOSID As Integer) As Boolean
Call ConectarBD
    fnPesquisarOcorrencia = False
    strOcorrencia = ""
    strDataOcorrencia = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " and  DataOSNaoValidada IS NOT NULL"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strOcorrencia = rs!MotivoOSNaoValidada & ""
        strDataOcorrencia = rs!DataOSNaoValidada & ""
        fnPesquisarOcorrencia = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function
