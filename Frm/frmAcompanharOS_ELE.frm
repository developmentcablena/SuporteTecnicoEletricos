VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcompanharOS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Acompanhar OS"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcompanharOS_ELE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwChamados 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSituacao 
      Appearance      =   0  'Flat
      Caption         =   "&Situação"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4455
      TabIndex        =   8
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceite 
      Appearance      =   0  'Flat
      Caption         =   "&Aceite"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   300
      Left            =   2985
      MaxLength       =   5
      TabIndex        =   5
      Top             =   420
      Width           =   1485
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      Caption         =   "&Pesquisar"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   405
      Width           =   1095
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   2775
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
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
      Left            =   2985
      TabIndex        =   6
      Top             =   180
      Width           =   555
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   3
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "frmAcompanharOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strSituacao As String
Private strDataSituacao As String

Private Sub cboStatus_Click()
    
    Select Case cboStatus.Text
        Case Is = "Em Aberto"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = False
            cmdCancelar.Enabled = True
            cmdImprimir.Enabled = True
        Case Is = "Urgente"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = False
            cmdCancelar.Enabled = True
            cmdImprimir.Enabled = True
        Case Is = "Em Análise"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = True
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Em Atendimento"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = True
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Aguardando Aceite"
            cmdAceite.Enabled = True
            cmdSituacao.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Finalizada"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Cancelada"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
        Case Is = "Não Validada"
            cmdAceite.Enabled = False
            cmdSituacao.Enabled = False
            cmdCancelar.Enabled = False
            cmdImprimir.Enabled = False
    End Select
    
    Call suListarOS(gintUsuarioID, cboStatus.Text)
End Sub

Private Sub cmdAceite_Click()
Dim itemSelecionado As ListItem
Dim i As Integer
Dim ValueOS2 As Long

    ValueOS2 = 0

    If lvwChamados.ListItems.Count = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Tecnico"
        Exit Sub
    End If
    
    
    Set itemSelecionado = Me.lvwChamados.selectedItem
    
    If itemSelecionado Is Nothing Then
        MsgBox "Selecione uma OS para continuar!", vbOKOnly + vbExclamation, "Suporte Manutenção"
        Exit Sub
    End If
    
    ValueOS2 = CLng(itemSelecionado.Text)
    Call frmAceite.Show(vbModal)
    Call cboStatus_Click
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer

    gintOSID = 0
    
    If lvwChamados.ListItems.Count = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To lvwChamados.ListItems.Count
        If lvwChamados.ListItems(i).Selected = True Then
            gintOSID = CLng(lvwChamados.ListItems(i).Text)
            If MsgBox("Tem certeza que deseja cancelar a OS " & Format(gintOSID, "0000") & "?", vbYesNo + vbQuestion, "Suporte Técnico") = vbYes Then
                'Call suCancelarOS(gintOSID)
                Call frmCancelarOS.Show(vbModal)
                Call suListarOS(gintUsuarioID, cboStatus.Text)
            End If
        End If
    Next

End Sub

Private Sub cmdSituacao_Click()
Dim i As Integer

    gintOSID = 0

    If lvwChamados.ListItems.Count = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To lvwChamados.ListItems.Count
        If lvwChamados.ListItems(i).Selected = True Then
            gintOSID = CLng(lvwChamados.ListItems(i).Text)
            With frmSituacao
                .txtOSID.Text = Format(gintOSID, "0000")
                .txtOSID.Enabled = False
                Call suPesquisarSituacao(gintOSID)
                .txtComentario.Text = strSituacao
                .txtDataAtual.Text = strDataSituacao
                .txtDataAtual.Enabled = False
                .txtComentario.Locked = True
                .cmdAtualizar.Visible = False
            End With
            Call frmSituacao.Show(vbModal)
            Call cboStatus_Click
        End If
    Next

End Sub

Private Sub suPesquisarSituacao(ByVal vOSID As Integer)
Call ConectarBD
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

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer
    
    gintOSID = 0
    
    If lvwChamados.ListItems.Count = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To lvwChamados.ListItems.Count
        If lvwChamados.ListItems(i).Selected = True Then
            gintOSID = CLng(lvwChamados.ListItems(i).Text)  ' Primeira coluna é o Text
            Call acrOS.Show(vbModal)
        End If
    Next
End Sub

Private Sub cmdPesquisar_Click()
    If Len(Trim(cboStatus.Text)) = 0 Then
        MsgBox "Selecione um status!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboStatus.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtOSID.Text)) > 0 Then
        If IsNumeric(txtOSID.Text) = True Then
            If fnPesquisarOSID(Trim(txtOSID.Text), gintUsuarioID, cboStatus.Text) = False Then
                MsgBox "Ordem de serviço não localizada!", vbOKOnly + vbExclamation, "Suporte Técnico"
            End If
        Else
            MsgBox "Digite um valor numérico!", vbOKOnly + vbExclamation, "Suporte Técnico"
        End If
    Else
        MsgBox "Digite o código da ordem de serviço!", vbOKOnly + vbExclamation, "Suporte Técnico"
    End If
End Sub

Private Sub Form_Load()
     ' Definir cabeçalhos do ListView
With Me.lvwChamados
    .ColumnHeaders.Add , , "OS ID", Width:=700
    .ColumnHeaders.Add , , "Usuario", Width:=2500
    .ColumnHeaders.Add , , "Depto.", Width:=2500
    .ColumnHeaders.Add , , "Prior.", Width:=730
    .ColumnHeaders.Add , , "Data Cadastro", Width:=2500
    .ColumnHeaders.Add , , "Necessidade", Width:=2500
    .ColumnHeaders.Add , , "Prev. Sitemas", Width:=2500
    .ColumnHeaders.Add , , "Atendente", Width:=2500
    .ColumnHeaders.Add , , "Data Final", Width:=2500
    .ColumnHeaders.Add , , "Data Aceite", Width:=2000
    .ColumnHeaders.Add , , "Data Cancel", Width:=2000
    .LabelEdit = lvwManual
End With
    'São as lista Ex: Em aberto - em atendimento
    Call suListarStatus
   
    
    
End Sub






Private Sub suCancelarOS(ByVal vOSID As Long)
Call ConectarBD
On Error GoTo Erro

    strSQL = "UPDATE tb_OS SET DataCancelamento = '" & Format(Now, "dd/MM/yyyy HH:mm:ss") & "',Status = 4 WHERE OSID = " & vOSID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
    MsgBox "OS " & Format(vOSID, "0000") & " cancelada com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
    Exit Sub
    
Erro:
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"

End Sub

Private Sub suMostrarOS(ByVal vOSID As Integer, ByVal vStatus As String)
Call ConectarBD
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        If vStatus = "Em Aberto" Then
            With frmSuporteSistemas
                .cboDivisao.Enabled = True
                .cboDivisao.Text = Format(rs!DivisaoID, "0000") & " - " & rs!Divisao
                .cboTipo.Enabled = True
                .cboTipo.Text = Format(rs!TipoID, "0000") & " - " & rs!Tipo
                .cboEspecificacao.Enabled = True
                .cboEspecificacao.Text = Format(rs!EspecificacaoID, "0000") & " - " & rs!Especificacao
                .txtObservacao.Enabled = True
                .txtObservacao.Text = rs!DescricaoServico
                '.txtReporteTecnico.Text = rs!ReporteTecnico & ""
                .chkPrioridade.Enabled = True
                .chkPrioridade.Value = IIf(rs!Prioridade = True, 1, 0)
                .txtPrazo.Enabled = True
                .txtPrazo.Text = rs!Prazo
                .txtPrevisao.Enabled = False
                .txtPrevisao.Text = rs!Previsao
                .lblQtdeCaracteres.Enabled = True
                .cmdCadastrar.Caption = "&Alterar"
                .cmdCadastrar.Enabled = True
                .cmdCancelar.Enabled = True
            End With
        Else
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
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function fnPesquisarOSID(ByVal vOSID As Integer, ByVal vUsuarioID As Integer, ByVal vStatus As String) As Boolean
Call ConectarBD
    fnPesquisarOSID = False
    
    If vStatus = "Em Aberto" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND UsuarioID = " & vUsuarioID & " AND Status = 0 ORDER BY OSID"
    ElseIf vStatus = "Em Atendimento" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND UsuarioID = " & vUsuarioID & " AND Status = 1 ORDER BY OSID"
    ElseIf vStatus = "Em Análise" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND UsuarioID = " & vUsuarioID & " AND Status = 7 ORDER BY OSID"
    ElseIf vStatus = "Aguardando Aceite" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND UsuarioID = " & vUsuarioID & " AND Status = 2 ORDER BY OSID"
    ElseIf vStatus = "Finalizada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND UsuarioID = " & vUsuarioID & " AND Status = 3 ORDER BY OSID"
    ElseIf vStatus = "Cancelada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID & " AND UsuarioID = " & vUsuarioID & " AND Status = 4 ORDER BY OSID"
    End If
    
    Call ConectarBD
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    lvwChamados.ListItems.Clear
    
    If rs.EOF = False Then
        With lvwChamados.ListItems.Add(, , rs!OSID)
            .SubItems(1) = rs!Nome
            .SubItems(2) = rs!Departamento
            If rs!Prioridade = False Then
                .SubItems(3) = "Não"
            Else
                .SubItems(3) = "Sim"
            End If
         
            .SubItems(4) = rs!Datacadastro
            .SubItems(5) = rs!Previsao
            If IsNull(rs!PrevisaoSistemas) Then
                .SubItems(6) = ""
            Else
                .SubItems(6) = rs!PrevisaoSitemas
            End If
            
           If IsNull(rs!Atendente) Then
                .SubItems(7) = ""
            Else
                .SubItems(7) = rs!Atendente
            End If
           
            If IsNull(rs!DataBaixa) Then
                .SubItems(8) = ""
            Else
                .SubItems(8) = rs!DataBaixa
            End If
        
            If IsNull(rs!DataAceite) Then
                .SubItems(9) = ""
            Else
                .SubItems(9) = rs!DataAceite
            End If
            If IsNull(rs!DataCancelamento) Then
                .SubItems(10) = ""
            Else
                .SubItems(10) = rs!DataCancelamento
            End If
        End With
        fnPesquisarOSID = True
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Function

Private Sub suListarOS(ByVal vUsuarioID As Integer, ByVal vStatus As String)

Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim lvwItem As ListItem
    Dim i As Integer
    
    If vStatus = "Em Aberto" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 0 ORDER BY OSID"
    ElseIf vStatus = "Em Atendimento" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 1 ORDER BY OSID"
    ElseIf vStatus = "Em Análise" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 7 ORDER BY OSID"
    ElseIf vStatus = "Aguardando Aceite" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 2 ORDER BY OSID"
    ElseIf vStatus = "Finalizada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 3 ORDER BY OSID"
    ElseIf vStatus = "Cancelada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 4 ORDER BY OSID"
    ElseIf vStatus = "Não Validada" Then
        strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 6 ORDER BY OSID"
    End If
        
     ' Inicializar conexão com o banco de dados
    Call ConectarBD
    
    strBD = "PROVIDER=SQLOLEDB;SERVER=196.200.80.20;DATABASE=HelpDesk;UID=cablena_user;PWD=C@bl3na;"
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open (strBD)
    
    ' Inicializar o Recordset
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic

    ' Limpar o ListView antes de adicionar novos itens
    Me.lvwChamados.ListItems.Clear
    
   ' Verificação se há registros retornados pela consulta
    If Not rs.EOF Then
        ' Loop para adicionar cada registro ao ListView
        Do While Not rs.EOF
            ' Criar um novo item de ListView
            Set lvwItem = Me.lvwChamados.ListItems.Add(, , rs!OSID)
            lvwItem.SubItems(1) = rs!Nome
            lvwItem.SubItems(2) = rs!Departamento
            If rs!Prioridade = False Then
                lvwItem.SubItems(3) = "Não"
            Else
                lvwItem.SubItems(3) = "Sim"
            End If
            lvwItem.SubItems(4) = rs!Datacadastro
            lvwItem.SubItems(5) = rs!Previsao
            If IsNull(rs!Atendente) Then
                lvwItem.SubItems(6) = ""
            Else
                lvwItem.SubItems(6) = rs!Atendente
            End If
            
            If IsNull(rs!DataBaixa) Then
                lvwItem.SubItems(7) = ""
            Else
                lvwItem.SubItems(7) = rs!DataBaixa
            End If
            
            If IsNull(rs!DataAceite) Then
                lvwItem.SubItems(8) = ""
            Else
                lvwItem.SubItems(8) = rs!DataAceite
            End If
            If IsNull(rs!DataCancelamento) Then
                lvwItem.SubItems(9) = ""
            Else
                lvwItem.SubItems(9) = rs!DataCancelamento
            End If
            
            gintOSID = rs!OSID

            rs.MoveNext
        Loop
    Else
        MsgBox "Nenhum registro encontrado!", vbOKOnly + vbInformation, "Consulta"
    End If

    ' Fechar o Recordset e a Conexão
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    Exit Sub

Erro:
    ' Tratamento de erro na abertura do Recordset
    MsgBox "Erro na consulta: " & Err.Description, vbOKOnly + vbCritical, "Erro"
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub



Private Sub suListarStatus()
    cboStatus.AddItem "Em Aberto"
    cboStatus.AddItem "Em Análise"
    cboStatus.AddItem "Em Atendimento"
    cboStatus.AddItem "Aguardando Aceite"
    cboStatus.AddItem "Finalizada"
    cboStatus.AddItem "Cancelada"
    cboStatus.AddItem "Não Validada"
End Sub

Private Sub lvwChamados_DblClick()
Dim i As Integer
    
    gintOSID = 0
    gblnTelaChamados = False
    
    If lvwChamados.ListItems.Count = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To lvwChamados.ListItems.Count
        If lvwChamados.ListItems(i).Selected = True Then
            gintOSID = CLng(lvwChamados.ListItems(i).Text)
            gblnTelaChamados = True
            Call suMostrarOS(gintOSID, cboStatus.Text)
            Call frmSuporteSistemas.Show(vbModal)
        End If
    Next
    
    gblnTelaChamados = False

End Sub

Private Sub txtOSID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtOSID.Text)) > 0 Then
            Call cmdPesquisar_Click
        End If
    End If
End Sub
