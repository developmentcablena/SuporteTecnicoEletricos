VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAcompanharOS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Acompanhar OS"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcompanharOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSituacao 
      Appearance      =   0  'Flat
      Caption         =   "&Situação"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4455
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceite 
      Appearance      =   0  'Flat
      Caption         =   "&Aceite"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6480
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin GridEX20.GridEX gexOS 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
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
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmAcompanharOS.frx":058A
      Column(2)       =   "frmAcompanharOS.frx":075E
      Column(3)       =   "frmAcompanharOS.frx":0952
      Column(4)       =   "frmAcompanharOS.frx":0B2A
      Column(5)       =   "frmAcompanharOS.frx":0D2A
      Column(6)       =   "frmAcompanharOS.frx":0E86
      Column(7)       =   "frmAcompanharOS.frx":1066
      Column(8)       =   "frmAcompanharOS.frx":124E
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmAcompanharOS.frx":1452
      FormatStyle(2)  =   "frmAcompanharOS.frx":157E
      FormatStyle(3)  =   "frmAcompanharOS.frx":162E
      FormatStyle(4)  =   "frmAcompanharOS.frx":16E2
      FormatStyle(5)  =   "frmAcompanharOS.frx":17BA
      ImageCount      =   0
      PrinterProperties=   "frmAcompanharOS.frx":1872
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
Dim i As Integer

    gintOSID = 0

    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If

    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            Call frmAceite.Show(vbModal)
            Call cboStatus_Click
        End If
    Next

End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer

    gintOSID = 0
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
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

    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhuma OS na lista!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
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

Private Sub gexOS_DblClick()
Dim i As Integer
    
    gintOSID = 0
    gblnTelaChamados = False
    
    If gexOS.RowCount = 0 Then
        MsgBox "Não há nenhum chamado listado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexOS.RowCount
        If gexOS.RowSelected(i) = True Then
            gintOSID = gexOS.Value(1)
            gblnTelaChamados = True
            Call suMostrarOS(gintOSID, cboStatus.Text)
            Call frmSuporteSistemas.Show(vbModal)
        End If
    Next
    
    gblnTelaChamados = False
End Sub

Private Sub Form_Load()
    Call suListarStatus
End Sub

Private Sub suCancelarOS(ByVal vOSID As Long)
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
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
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

Private Sub suListarOS(ByVal vUsuarioID As Integer, ByVal vStatus As String)
    
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
    cboStatus.AddItem "Em Aberto"
    cboStatus.AddItem "Em Análise"
    cboStatus.AddItem "Em Atendimento"
    cboStatus.AddItem "Aguardando Aceite"
    cboStatus.AddItem "Finalizada"
    cboStatus.AddItem "Cancelada"
    cboStatus.AddItem "Não Validada"
End Sub

Private Sub txtOSID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtOSID.Text)) > 0 Then
            Call cmdPesquisar_Click
        End If
    End If
End Sub
