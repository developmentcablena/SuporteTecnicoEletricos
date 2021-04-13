VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmChamados 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Verificar Chamados"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChamados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSituacao 
      Appearance      =   0  'Flat
      Caption         =   "&Situação"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkHistorico 
      Appearance      =   0  'Flat
      Caption         =   "Visualizar Ocorrências"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   435
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnalisar 
      Appearance      =   0  'Flat
      Caption         =   "A&nalisar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CheckBox chkComentario 
      Appearance      =   0  'Flat
      Caption         =   "Visualizar Comentário"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   435
      Width           =   1455
   End
   Begin VB.CommandButton cmdAtender 
      Appearance      =   0  'Flat
      Caption         =   "&Atender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      Caption         =   "&Pesquisar"
      Height          =   300
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   495
      Width           =   1335
   End
   Begin VB.CommandButton cmdBaixarOS 
      Appearance      =   0  'Flat
      Caption         =   "&Finalizar OS"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   6840
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin GridEX20.GridEX gexOS 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9975
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   11
      Column(1)       =   "frmChamados.frx":1CCA
      Column(2)       =   "frmChamados.frx":1EC2
      Column(3)       =   "frmChamados.frx":202E
      Column(4)       =   "frmChamados.frx":21BA
      Column(5)       =   "frmChamados.frx":245E
      Column(6)       =   "frmChamados.frx":2636
      Column(7)       =   "frmChamados.frx":284E
      Column(8)       =   "frmChamados.frx":2A52
      Column(9)       =   "frmChamados.frx":2BAE
      Column(10)      =   "frmChamados.frx":2D96
      Column(11)      =   "frmChamados.frx":2F96
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmChamados.frx":319A
      FormatStyle(2)  =   "frmChamados.frx":32C6
      FormatStyle(3)  =   "frmChamados.frx":3376
      FormatStyle(4)  =   "frmChamados.frx":342A
      FormatStyle(5)  =   "frmChamados.frx":3502
      ImageCount      =   0
      PrinterProperties=   "frmChamados.frx":35BA
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº OS"
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
      Left            =   2880
      TabIndex        =   8
      Top             =   255
      Width           =   555
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmChamados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strPrevisaoSistemas As String
Private strSituacao As String
Private strDataSituacao As String
Private strRelatorio As String

Private Sub cboStatus_Click()
    
    Select Case cboStatus.Text
        Case Is = "Em Aberto"
            cmdAtender.Enabled = True
            cmdAnalisar.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Urgente"
            cmdAtender.Enabled = True
            cmdAnalisar.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbRed
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Em Análise"
            cmdAtender.Enabled = True
            cmdAnalisar.Enabled = False
            cmdBaixarOS.Enabled = False
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = True
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Em Atendimento"
            cmdAtender.Enabled = False
            cmdAnalisar.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = True
            cmdSituacao.Visible = True
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Aguardando Aceite"
            cmdAtender.Enabled = False
            cmdAnalisar.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = False
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Finalizada"
            cmdAtender.Enabled = False
            cmdAnalisar.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Comentário"
            chkComentario.Value = 0
            chkComentario.Enabled = True
            gexOS.ForeColor = vbBlue
            chkHistorico.Enabled = True
            chkHistorico.Value = 0
        Case Is = "Cancelada"
            cmdAtender.Enabled = False
            cmdAnalisar.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = False
            cmdBaixarOS.Enabled = False
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Motivo"
            chkComentario.Value = 0
            chkComentario.Enabled = True
            gexOS.ForeColor = vbBlack
            chkHistorico.Enabled = False
            chkHistorico.Value = 0
        Case Is = "Não Validada"
            cmdAtender.Enabled = False
            cmdAnalisar.Enabled = False
            cmdBaixarOS.Enabled = True
            cmdImprimir.Enabled = True
            cmdBaixarOS.Enabled = True
            cmdSituacao.Visible = False
            chkComentario.Caption = "Visualizar Motivo"
            chkComentario.Value = 0
            chkComentario.Enabled = True
            gexOS.ForeColor = vbRed
            chkHistorico.Enabled = True
            chkHistorico.Value = 0
    End Select
        
    Call suListarChamados(cboStatus.Text)
End Sub

Private Sub chkComentario_Click()
    If chkComentario.Value = 1 Then
        chkHistorico.Value = 0
    End If
End Sub

Private Sub chkHistorico_Click()
    If chkHistorico.Value = 1 Then
        chkComentario.Value = 0
    End If
End Sub

Private Sub cmdAnalisar_Click()
Dim i As Integer

    gintOSID = 0
    strPrevisaoSistemas = ""
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            
            If MsgBox("Tem certeza que deseja analisar a OS " & Format(gintOSID, "0000") & "?", vbYesNo + vbQuestion, "Suporte Técnico") = vbYes Then
                If fnStatusAnalise(gintOSID, gstrNome) = True Then
                    Call suListarChamados(cboStatus.Text)
                End If
            Else
                Exit Sub
            End If
        End If
    Next

End Sub

Private Function fnStatusAnalise(ByVal vOSID As Integer, ByVal vAtendente As String) As Boolean
On Error GoTo Erro

    fnStatusAnalise = False
    strSQL = "UPDATE tb_OS SET Status = 7,Atendente = '" & vAtendente & "' WHERE OSID=" & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    fnStatusAnalise = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
End Function

Private Sub cmdAtender_Click()
Dim i As Integer

    gintOSID = 0
    strPrevisaoSistemas = ""
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            
            If MsgBox("Tem certeza que deseja atender a OS " & Format(gintOSID, "0000") & "?", vbYesNo + vbQuestion, "Suporte Técnico") = vbYes Then
                
                strPrevisaoSistemas = InputBox("Digite a data de previsão para conclusão do serviço!", "Suporte Técnico")
                                
                If IsDate(strPrevisaoSistemas) = False Then
                    MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Técnico"
                    Exit Sub
                End If
                
                If fnCadastrarAtendente(gintOSID, gstrNome, CDate(strPrevisaoSistemas)) = True Then
                    Call suListarChamados(cboStatus.Text)
                End If
                If fnEnviarEmail(gintOSID, fnCapturarEMail(gintOSID), fnCapturarUsuario(gintOSID), fnCapturarEmailAtendente(gstrNome)) = True Then
                    MsgBox "Email enviado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
                End If
            Else
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Function fnEnviarEmail(ByVal vOSID As Integer, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEMailAtendente As String) As Boolean
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
    poSendMail.Subject = "OS " & Format(vOSID, "0000") & " em Atendimento"
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
        fnContaSMTP = Trim(rs!Valor)
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
        fnSenhaSMTP = Trim(rs!Valor)
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
        strRelatorio = strRelatorio & "Situação: Em atendimento" & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Data Previsão: " & Format(rs!PrevisaoSistemas, "dd/MM/yyyy") & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        'strRelatorio = strRelatorio & "*ATENÇÃO: Utilize também a ferramenta SuporteWEB para visualizar o andamento das Ordens de Serviço." & vbCrLf & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "**OBSERVAÇÃO: Esta mensagem é gerada automaticamente pelo sistema." & vbCrLf & "POR FAVOR NÃO RESPONDA ESTA MENSAGEM." & vbCrLf & vbCrLf & vbCrLf & "Suporte Técnico" & vbCrLf & "Cablena do Brasil" & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(100, "=")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function fnCapturarEmailAtendente(ByVal vAtendente As String) As String
Dim rs1 As New ADODB.Recordset
    
    fnCapturarEmailAtendente = ""
    
    strSQL = "SELECT Nome,Email FROM tb_Usuarios WHERE Nome = '" & vAtendente & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnCapturarEmailAtendente = rs!Email & ""
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
        fnCapturarEMail = rs!Email & ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function


Private Function fnCadastrarAtendente(ByVal vOSID As Integer, ByVal vAtendente As String, ByVal vPrevisaoSistemas As Variant) As Boolean
On Error GoTo Erro
    
    fnCadastrarAtendente = False
        
    strSQL = "UPDATE tb_OS SET Atendente = '" & vAtendente & "',PrevisaoSistemas='" & Format(vPrevisaoSistemas, "yyyy-MM-dd") & "', Status = 1 WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnCadastrarAtendente = True
    Set rs = Nothing
    Exit Function
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
End Function

Private Sub cmdBaixarOS_Click()
Dim i As Integer

    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call frmReporteTecnico.Show(vbModal)
        End If
    Next
    
    Call cboStatus_Click
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer
    
    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call acrOS.Show(vbModal)
        End If
    Next

End Sub

Private Sub cmdSituacao_Click()
Dim i As Integer

    gintOSID = 0
    gstrcboStatus = ""

    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            gstrcboStatus = cboStatus.Text
            With frmSituacao
                .txtOSID.Text = Format(gintOSID, "0000")
                .txtOSID.Enabled = False
                .txtDataAtual.Enabled = False
                Call suPesquisarSituacao(gintOSID)
                .txtComentario.Text = strSituacao
                .txtDataAtual.Text = strDataSituacao
                If Len(Trim(.txtComentario.Text)) > 0 Then
                    .cmdAtualizar.Caption = "&Novo"
                End If
            End With
            Call frmSituacao.Show(vbModal)
            Call cboStatus_Click
        End If
    Next
End Sub

Private Sub suPesquisarSituacao(ByVal vOSID As Integer)
    strSQL = "SELECT Situacao,DataSituacao FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        strSituacao = rs!Situacao & ""
        strDataSituacao = Format(rs!DataSituacao, "dd/MM/yyyy") & ""
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Call suListarStatus
End Sub

Private Sub cmdPesquisar_Click()
    If Len(Trim(cboStatus.Text)) = 0 Then
        MsgBox "Selecione um status!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboStatus.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtOSID.Text)) = 0 Then
        MsgBox "Digite o Nº da OS!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtOSID.SetFocus
        Exit Sub
    Else
        If IsNumeric(Trim(txtOSID.Text)) = True Then
            If fnPesquisarOSID(Trim(txtOSID.Text), cboStatus.Text) = False Then
                MsgBox "OS não localizada!", vbOKOnly + vbExclamation, "Suporte Técnico"
                Exit Sub
            End If
        Else
            MsgBox "Digite um valor numérico!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtOSID.Text = ""
            txtOSID.SetFocus
        End If
    End If
End Sub

Private Function fnPesquisarOSID(ByVal vOSID As Integer, ByVal vStatus As String) As Boolean
    fnPesquisarOSID = False
    
    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 0 AND Prioridade = 0 ORDER BY OSID"
        Case Is = "Urgente"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 0 AND Prioridade = 1 ORDER BY OSID"
        Case Is = "Em Análise"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 7 ORDER BY OSID"
        Case Is = "Em Atendimento"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 1 ORDER BY OSID"
        Case Is = "Aguardando Aceite"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 2 ORDER BY OSID"
        Case Is = "Finalizada"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 3 ORDER BY OSID"
        Case Is = "Cancelada"
            strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND Status = 4 ORDER BY OSID"
    End Select
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = False Then
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
        fnPesquisarOSID = True
    Else
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suListarChamados(ByVal vStatus As String)
    
    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 AND Prioridade = 0 ORDER BY OSID"
        Case Is = "Urgente"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 AND Prioridade = 1 ORDER BY OSID"
        Case Is = "Em Análise"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 7 ORDER BY OSID"
        Case Is = "Em Atendimento"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 1 ORDER BY OSID"
        Case Is = "Aguardando Aceite"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 2 ORDER BY OSID"
        Case Is = "Finalizada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 3 ORDER BY OSID"
        Case Is = "Cancelada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 4 ORDER BY OSID"
        Case Is = "Não Validada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 6 ORDER BY OSID"
    End Select
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = False Then
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
    Else
        gexOS.HoldFields
        Set gexOS.ADORecordset = rs
    End If
    
    Set rs = Nothing
End Sub

Private Sub suListarStatus()
    cboStatus.Clear
    cboStatus.AddItem "Em Aberto"
    cboStatus.AddItem "Urgente"
    cboStatus.AddItem "Em Análise"
    cboStatus.AddItem "Em Atendimento"
    cboStatus.AddItem "Aguardando Aceite"
    cboStatus.AddItem "Finalizada"
    cboStatus.AddItem "Cancelada"
    cboStatus.AddItem "Não Validada"
End Sub

Private Sub gexOS_DblClick()
Dim i As Integer
    
    gintOSID = 0
    gblnTelaChamados = False
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    If chkHistorico.Value = 1 Then
        For i = 1 To gexOS.RowCount
            If gexOS.RowSelected(i) = True Then
                gintOSID = gexOS.Value(1)
                Call suMostrarOcorrencia(gintOSID)
                Call frmOcorrencias.Show(vbModal)
            End If
        Next
        Exit Sub
    End If
    
    If chkComentario.Value = 0 Then
        For i = 1 To gexOS.RowCount
            If gexOS.RowSelected(i) = True Then
                gintOSID = gexOS.Value(1)
                gblnTelaChamados = True
                Call suMostrarOS(gintOSID)
                Call frmSuporteSistemas.Show(vbModal)
            End If
        Next
    Else
        For i = 1 To gexOS.RowCount
            If gexOS.RowSelected(i) = True Then
                gintOSID = gexOS.Value(1)
                If cboStatus.Text = "Finalizada" Then
                    Call suMostrarComentario(gintOSID)
                    Call frmAceite.Show(vbModal)
                ElseIf cboStatus.Text = "Cancelada" Then
                    Call suMostrarMotivo(gintOSID)
                    Call frmCancelarOS.Show(vbModal)
                ElseIf cboStatus.Text = "Não Validada" Then
                    Call suMostrarNaoValidada(gintOSID)
                    Call frmAceite.Show(vbModal)
                End If
            End If
        Next
    End If
    
    gblnTelaChamados = False
End Sub

Private Sub suMostrarOcorrencia(ByVal vOSID As Long)
    strSQL = "SELECT * FROM tb_Ocorrencias WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmOcorrencias
            .gexOcorrencias.HoldFields
            Set .gexOcorrencias.ADORecordset = rs
        End With
    End If
    
    Set rs = Nothing
End Sub


Private Sub suMostrarNaoValidada(ByVal vOSID As Long)
    strSQL = "SELECT OSID,MotivoOSNaoValidada FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmAceite
            .txtOSID.Text = Format(rs!OSID, "0000")
            .txtComentario.Text = rs!MotivoOSNaoValidada & ""
            .cmdAceite.Enabled = False
            .cmdNaoValidar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suMostrarMotivo(ByVal vOSID As Long)
    strSQL = "SELECT OSID,Comentario FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmCancelarOS
            .txtOSID.Text = Format(rs!OSID, "0000")
            .txtMotivo.Text = rs!Comentario & ""
            .cmdCancelar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suMostrarComentario(ByVal vOSID As Long)
    strSQL = "SELECT OSID,Comentario FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmAceite
            .txtOSID.Text = Format(rs!OSID, "0000")
            .txtComentario.Text = rs!Comentario & ""
            .cmdAceite.Enabled = False
            .cmdNaoValidar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suMostrarOS(ByVal vOSID As Long)
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With frmSuporteSistemas
            .cboDivisao.Enabled = False
            .cboDivisao.Text = Format(rs!DivisaoID, "0000") & " - " & rs!Divisao
            .cboTipo.Enabled = False
            .cboTipo.Text = Format(rs!TipoID, "0000") & " - " & rs!Tipo
            .cboEspecificacao.Enabled = False
            .cboEspecificacao.Text = Format(rs!EspecificacaoID, "0000") & " - " & rs!Especificacao
            .txtObservacao.Enabled = True
            .txtObservacao.Locked = True
            .txtObservacao.Text = rs!DescricaoServico
            .txtReporteTecnico.Enabled = True
            .txtReporteTecnico.Locked = True
            .txtReporteTecnico.Text = rs!ReporteTecnico & ""
            .chkPrioridade.Enabled = False
            .chkPrioridade.Value = IIf(rs!Prioridade = True, 1, 0)
            .txtPrazo.Enabled = False
            .txtPrazo.Text = rs!Prazo
            .txtPrevisao.Enabled = False
            .txtPrevisao.Text = rs!Previsao
            .lblQtdeCaracteres.Enabled = False
            .cmdCadastrar.Enabled = False
            .cmdCancelar.Enabled = False
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub txtOSID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtOSID.Text)) > 0 Then
            Call cmdPesquisar_Click
        End If
    End If
End Sub
