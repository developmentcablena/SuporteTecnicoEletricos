VERSION 5.00
Begin VB.Form frmSituacao 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Situação OS"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   4455
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "&Cadastrar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox txtDataAtual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   3
      Top             =   150
      Width           =   1575
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      MaxLength       =   5
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblComentario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Comentário"
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
      TabIndex        =   6
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label lblDataAtual 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Atualização:"
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
      Height          =   240
      Left            =   2280
      TabIndex        =   4
      Top             =   270
      Width           =   1920
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº OS:"
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
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   270
      Width           =   705
   End
End
Attribute VB_Name = "frmSituacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private Sub cmdAtualizar_Click()
    If Len(Trim(txtOSID.Text)) = 0 Then
        MsgBox "OS não encontrada!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    If cmdAtualizar.Caption = "&Cadastrar" Then
        If fnCadastrarSituacao(txtOSID.Text, txtComentario.Text) = True Then
            MsgBox "Situação cadastrada com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
        End If
    ElseIf cmdAtualizar.Caption = "&Adicionar" Then
        If fnLogarSituacao(txtOSID.Text) = True Then
            If fnCadastrarSituacao(txtOSID.Text, txtComentario.Text) = True Then
                MsgBox "Situação adicionada com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
            End If
            cmdAtualizar.Caption = "&Novo"
        End If
    Else
        txtComentario.Text = ""
        txtDataAtual.Text = ""
        cmdAtualizar.Caption = "&Adicionar"
        Exit Sub
    End If
    
    If fnEnviarEmail(txtOSID.Text, txtComentario.Text, fnCapturarEMail(txtOSID.Text), fnCapturarUsuario(txtOSID.Text), fnCapturarEmailAtendente(txtOSID.Text)) = True Then
        MsgBox "Email enviado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
    End If

    Call Unload(Me)
End Sub

Private Function fnLogarSituacao(ByVal vOSID As Integer) As Boolean
On Error GoTo Erro
Dim rs1 As ADODB.Recordset
    
    fnLogarSituacao = False
    
    strSQL = "SELECT Situacao, DataSituacao FROM tb_Os WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strSQL = "INSERT INTO tb_Log_Situacoes (Situacao,DataSituacao,OSID) VALUES ('" & rs!Situacao & "','" & Format(rs!DataSituacao, "yyyy-MM-dd HH:mm") & "'," & vOSID & ")"
        Set rs1 = New ADODB.Recordset
        rs1.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        
        Set rs1 = Nothing
        fnLogarSituacao = True
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
Erro:
    rs.Close
    Set rs = Nothing
    MsgBox "Erro " & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"

End Function

Private Function fnCadastrarSituacao(ByVal vOSID As Integer, ByVal vSituacao As String) As Boolean
On Error GoTo Erro

    fnCadastrarSituacao = False
    
    strSQL = "UPDATE tb_OS SET Situacao = '" & vSituacao & "', DataSituacao='" & Format(Now, "yyyy-MM-dd HH:mm") & "' WHERE OSID=" & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        
    Set rs = Nothing
    fnCadastrarSituacao = True
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
End Function

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Function fnEnviarEmail(ByVal vOSID As Integer, ByVal vSituacao As String, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEMailAtendente As String) As Boolean
On Error GoTo Erro
Dim poSendMail As vbSendMail.clsSendMail
    
    fnEnviarEmail = False
    Me.Enabled = False
    Screen.MousePointer = 11
    
    DoEvents
    
    Set poSendMail = New vbSendMail.clsSendMail
    poSendMail.SMTPHost = "email-ssl.com.br"
    poSendMail.SMTPPort = "587"
    poSendMail.UseAuthentication = True
    poSendMail.UserName = fnContaSMTP
    poSendMail.Password = fnSenhaSMTP
    poSendMail.From = fnContaSMTP
    poSendMail.FromDisplayName = "ADM Suporte Técnico"
    poSendMail.Recipient = vEMail
    poSendMail.RecipientDisplayName = vUsuario
    poSendMail.CcRecipient = vEMailAtendente
    If gstrcboStatus = "Em Análise" Then
        poSendMail.Subject = "OS " & Format(vOSID, "0000") & " em Análise"
    ElseIf gstrcboStatus = "Em Atendimento" Then
        poSendMail.Subject = "OS " & Format(vOSID, "0000") & " em Atendimento"
    End If
    poSendMail.Priority = HIGH_PRIORITY
    Call suRelatorio(vOSID)
    poSendMail.Message = strRelatorio
    poSendMail.Send

    Set poSendMail = Nothing
    fnEnviarEmail = True
    Screen.MousePointer = 0
    Me.Enabled = True
    Exit Function

Erro:
    Screen.MousePointer = 0
    Me.Enabled = True
    MsgBox "Erro: " & Err.Description, vbOKOnly = vbCritical, "Suporte Técnico"
    Set poSendMail = Nothing

End Function

Private Function fnContaSMTP() As String
    fnContaSMTP = ""
    
    strSQL = "SELECT * FROM dbo.tb_Configuracoes WHERE ID = 1"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnContaSMTP = Trim(rs!valor)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnSenhaSMTP() As String
    fnSenhaSMTP = ""
    
    strSQL = "SELECT * FROM dbo.tb_Configuracoes WHERE ID = 2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnSenhaSMTP = Trim(rs!valor)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suRelatorio(ByVal vOSID As Long)
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strRelatorio = ""
        strRelatorio = String(100, "=") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Nº OS: " & Format(rs!OSID, "0000") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Tipo: " & rs!Divisao & String(10, " ") & "Caract.: " & rs!Tipo & String(10, " ") & "Especificação: " & rs!Especificacao & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Reporte Usuário: " & rs!DescricaoServico & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "SITUAÇÃO: " & rs!Situacao & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Data Cadastro: " & Format(rs!DataSituacao, "dd/MM/yy HH:mm") & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        'strRelatorio = strRelatorio & "*ATENÇÃO: Utilize também a ferramenta SuporteWEB para visualizar o andamento das Ordens de Serviço." & vbCrLf & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "**OBSERVAÇÃO: Esta mensagem é gerada automaticamente pelo sistema." & vbCrLf & "POR FAVOR NÃO RESPONDA ESTA MENSAGEM." & vbCrLf & vbCrLf & vbCrLf & "Suporte Técnico" & vbCrLf & "Cablena do Brasil" & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(100, "=")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function fnCapturarEmailAtendente(ByVal vOSID As Long) As String
Dim rs1 As New ADODB.Recordset
    
    fnCapturarEmailAtendente = ""
    
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strSQL = "SELECT * FROM vw_Usuarios WHERE Nome = '" & rs!Atendente & "'"
        rs1.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
        
        If rs1.EOF = False Then
            fnCapturarEmailAtendente = rs1!EMail & ""
        End If
        
        rs1.Close
        Set rs1 = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnCapturarUsuario(ByVal vOSID As Long) As String
    fnCapturarUsuario = ""
    
    strSQL = "SELECT Nome FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnCapturarUsuario = rs!Nome & ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnCapturarEMail(ByVal vOSID As Long) As String
    fnCapturarEMail = ""
    
    strSQL = "SELECT EMail FROM vw_Chamados WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnCapturarEMail = rs!EMail & ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function

