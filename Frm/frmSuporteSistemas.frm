VERSION 5.00
Begin VB.Form frmSuporteSistemas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suporte Técnico - Sistemas"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuporteSistemas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtReporteTecnico 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3960
      Width           =   7335
   End
   Begin VB.CheckBox chkPrioridade 
      Appearance      =   0  'Flat
      Caption         =   "Prioridade"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1635
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   5865
      Width           =   1335
   End
   Begin VB.TextBox txtPrevisao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   5955
      Width           =   1335
   End
   Begin VB.TextBox txtPrazo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   5955
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   5865
      Width           =   1335
   End
   Begin VB.CommandButton cmdCadastrar 
      Appearance      =   0  'Flat
      Caption         =   "&Cadastrar"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5865
      Width           =   1335
   End
   Begin VB.TextBox txtObservacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   1695
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   7335
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   7335
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblLimite 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Limite de caracteres: "
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
      Left            =   4320
      TabIndex        =   19
      Top             =   1680
      Width           =   2130
   End
   Begin VB.Label lblQtdeCaracteres 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6480
      TabIndex        =   18
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label lblReporteTecnico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Técnico:"
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
      TabIndex        =   17
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label lblPrevisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Necessidade"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   5715
      Width           =   1230
   End
   Begin VB.Label lblPrazo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Prazo (dias)"
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
      TabIndex        =   11
      Top             =   5715
      Width           =   1200
   End
   Begin VB.Label lblDescricaoServico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Usuário:"
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
      TabIndex        =   9
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Label lblEspecificacao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Especificação"
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
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Característica"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmSuporteSistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private Sub cboDivisao_Click()
    cboTipo.Clear
    cboEspecificacao.Clear

    If Len(Trim(cboDivisao.Text)) > 0 Then
        Call suListarTipo(Left(cboDivisao.Text, 4))
    End If
End Sub

Private Sub cboTipo_Click()
    cboEspecificacao.Clear
    
    If Len(Trim(cboTipo.Text)) > 0 Then
        Call suListarEspecificacao(Left(cboTipo.Text, 4))
    End If
End Sub

Private Sub chkPrioridade_Click()
    If chkPrioridade.Value = 1 Then
        chkPrioridade.ForeColor = &HFF&
    Else
        chkPrioridade.ForeColor = &H80000008
    End If
End Sub

Private Sub cmdCancelar_Click()
    If cmdCadastrar.Caption = "&Cadastrar" Then
        Call suLimparCampos
    Else
        Call Unload(Me)
    End If
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    With Me
        .lblReporteTecnico.Enabled = True
        .txtReporteTecnico.Enabled = False
    End With
    Call suListarDivisao
End Sub

Private Sub lblQtdeCaracteres_Change()
    If lblQtdeCaracteres.Caption = 0 Then
        lblQtdeCaracteres.ForeColor = vbRed
    Else
        lblQtdeCaracteres.ForeColor = vbBlue
    End If
End Sub

Private Sub txtObservacao_Change()
    lblQtdeCaracteres.Caption = 500 - Len(txtObservacao.Text)
End Sub

Private Sub txtPrazo_LostFocus()
    If Len(Trim(txtPrazo.Text)) > 0 Then
        If IsNumeric(txtPrazo.Text) = True Then
            If txtPrazo.Text >= 0 Then
                txtPrevisao.Text = Format(Now + Trim(txtPrazo.Text), "dd/MM/yyyy")
            Else
                MsgBox "Prazo não pode ser menor que 0!", vbOKOnly + vbExclamation, "Suporte Técnico"
                txtPrazo.Text = ""
                txtPrazo.SetFocus
            End If
        Else
            MsgBox "Digite a quantidade de dias!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtPrazo.Text = ""
            txtPrazo.SetFocus
        End If
    Else
        MsgBox "Digite um prazo para o serviço!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtPrazo.SetFocus
    End If
End Sub

Private Sub cmdCadastrar_Click()
    If Len(Trim(cboDivisao.Text)) = 0 Then
        MsgBox "Selecione uma divisão!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboDivisao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cboTipo.Text)) = 0 Then
        MsgBox "Selecione um tipo!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboTipo.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cboEspecificacao.Text)) = 0 Then
        MsgBox "Selecione uma especificação!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboEspecificacao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtObservacao.Text)) = 0 Then
        MsgBox "Digite a descrição do serviço à ser realizado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtObservacao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtPrazo.Text)) = 0 Then
        MsgBox "Digite um prazo para o serviço!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtPrazo.SetFocus
        Exit Sub
    End If
    
    If cmdCadastrar.Caption = "&Cadastrar" Then
        If fnCadastrarOS(Left(cboDivisao.Text, 4), Left(cboTipo.Text, 4), Left(cboEspecificacao.Text, 4), gintUsuarioID, gstrEMail, Trim(txtObservacao.Text), chkPrioridade, Trim(txtPrazo.Text), Trim(txtPrevisao.Text)) = True Then
            MsgBox "Ordem de serviço nº " & Format(gintOSID, "0000") & " gerada com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
            Call suLimparCampos
        End If
    Else
        If fnAlterarOS(Left(cboDivisao.Text, 4), Left(cboTipo.Text, 4), Left(cboEspecificacao.Text, 4), Trim(txtObservacao.Text), chkPrioridade, Trim(txtPrazo.Text), Trim(txtPrevisao.Text)) = True Then
            MsgBox "Ordem de serviço nº " & Format(gintOSID, "0000") & " alterada com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
            Call Unload(Me)
        End If
    End If
End Sub

Private Function fnAlterarOS(ByVal vDivisaoID As Integer, ByVal vTipoID As Integer, ByVal vEspecificacaoID As Integer, ByVal vDescricao As String, ByRef vPrioridade As CheckBox, ByVal vPrazo As Integer, ByVal vPrevisao As Variant) As Boolean
On Error GoTo Erro
    fnAlterarOS = False
    
    strSQL = "UPDATE tb_OS SET DivisaoID = " & vDivisaoID & ",TipoID = " & vTipoID & ",EspecificacaoID = " & vEspecificacaoID & ",DescricaoServico = '" & vDescricao & "',Prioridade = " & vPrioridade & ",Prazo = " & vPrazo & ",Previsao = '" & Format(vPrevisao, "yyyy-MM-dd") & "' WHERE OSID = " & gintOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnAlterarOS = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
End Function

Private Function fnCadastrarOS(ByVal vDivisaoID As Integer, ByVal vTipoID As Integer, ByVal vEspecificacaoID As Integer, ByVal vUsuarioID As Integer, ByVal vEMail As String, ByVal vDescricao As String, ByRef vPrioridade As CheckBox, ByVal vPrazo As Integer, ByVal vPrevisao As Variant) As Boolean
On Error GoTo Erro
    fnCadastrarOS = False
    gintOSID = 0
    
    strSQL = "SELECT * FROM tb_OS WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!DivisaoID = 0 & vDivisaoID
        rs!TipoID = 0 & vTipoID
        rs!EspecificacaoID = 0 & vEspecificacaoID
        rs!usuarioID = vUsuarioID
        rs!EMail = vEMail & ""
        rs!DescricaoServico = vDescricao & ""
        rs!Prioridade = vPrioridade
        rs!Prazo = 0 & vPrazo
        rs!Previsao = vPrevisao
        rs!Datacadastro = Format(Now, "dd/MM/yyyy HH:mm:ss")
        rs!Status = 0
        rs.Update
        gintOSID = rs!OSID
        fnCadastrarOS = True
        
        If fnEnviarEmail(gintOSID, gstrEMail, gstrNome) = False Then
            MsgBox "Erro ao enviar o e-mail!", vbOKOnly + vbCritical, "Suporte Técnico"
            Exit Function
        End If
    End If
    
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
End Function

Private Function fnEnviarEmail(ByVal vOSID As Integer, ByVal vEMail As String, ByVal vUsuario As String) As Boolean
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
    poSendMail.FromDisplayName = vUsuario
    poSendMail.Recipient = "ti@cablena.com.br" 'colocar email ti
    poSendMail.RecipientDisplayName = "ADM Suporte Técnico"
    poSendMail.Subject = "Nova OS TELECOM " & Format(vOSID, "0000")
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
        strRelatorio = strRelatorio & "Data Cadastro: " & Format(rs!Datacadastro, "dd/MM/yy HH:mm") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(100, "=")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suLimparCampos()
    cboDivisao.Clear
    Call suListarDivisao
    Call cboDivisao_Click
    txtObservacao.Text = ""
    chkPrioridade.Value = 0
    txtPrazo.Text = ""
    txtPrevisao.Text = ""
End Sub

Private Sub suListarDivisao()
    
    If gblnTelaChamados = True Then
        strSQL = "SELECT * FROM tb_Divisao ORDER BY DivisaoID"
    Else
        strSQL = "SELECT * FROM tb_Divisao WHERE Inativo=0 ORDER BY DivisaoID"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboDivisao.AddItem Format(rs!DivisaoID, "0000") & " - " & rs!Divisao
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarTipo(ByVal vDivisaoID As Integer)

    If gblnTelaChamados = True Then
        strSQL = "SELECT * FROM tb_Tipos WHERE DivisaoID = " & vDivisaoID
    Else
        strSQL = "SELECT * FROM tb_Tipos WHERE DivisaoID = " & vDivisaoID & " AND Inativo=0"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboTipo.AddItem Format(rs!TipoID, "0000") & " - " & rs!Tipo
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarEspecificacao(ByVal vTipoID As Integer)

    If gblnTelaChamados = True Then
        strSQL = "SELECT * FROM tb_Especificacoes WHERE TipoID = " & vTipoID
    Else
        strSQL = "SELECT * FROM tb_Especificacoes WHERE TipoID = " & vTipoID & " AND Inativo=0"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboEspecificacao.AddItem Format(rs!EspecificacaoID, "0000") & " - " & rs!Especificacao
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub
