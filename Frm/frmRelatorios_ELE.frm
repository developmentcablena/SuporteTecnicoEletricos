VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form frmRelatorios 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Relatórios"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorios_ELE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwRelatorio 
      Height          =   3135
      Left            =   120
      TabIndex        =   25
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
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
   Begin VB.ComboBox cboRelOS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "frmRelatorios_ELE.frx":030A
      Left            =   3825
      List            =   "frmRelatorios_ELE.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5235
      Width           =   2295
   End
   Begin VB.CheckBox chkStatus 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   750
      Width           =   975
   End
   Begin VB.CheckBox chkUsuario 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   750
      Width           =   975
   End
   Begin VB.CheckBox chkEspec 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   105
      Width           =   975
   End
   Begin VB.CheckBox chkCaract 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   105
      Width           =   975
   End
   Begin VB.CheckBox chkTipo 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   105
      Width           =   975
   End
   Begin VB.ComboBox cboUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1020
      Width           =   2295
   End
   Begin VB.TextBox txtDataCadAte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox txtDataCadDe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   1035
      Width           =   1215
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cdlSalvarArquivo 
      Left            =   9720
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pasta de Trabalho do Microsoft Office Excel (*.xls) | *.xls"
   End
   Begin MSComctlLib.ProgressBar pgbProgresso 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar XLS"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "&Gerar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1020
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      Left            =   3360
      TabIndex        =   8
      Top             =   5280
      Width           =   420
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
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
      TabIndex        =   18
      Top             =   780
      Width           =   750
   End
   Begin VB.Label lblAte 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "até"
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
      Left            =   6225
      TabIndex        =   17
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label lblDataCadastro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cadastro"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   795
      Width           =   1380
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caract."
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
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   690
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
      Left            =   4920
      TabIndex        =   14
      Top             =   120
      Width           =   1335
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
      TabIndex        =   13
      Top             =   120
      Width           =   420
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
      Left            =   2520
      TabIndex        =   9
      Top             =   780
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private Sub cboDivisao_Click()
   
    
    pgbProgresso.Value = 0
    cboTipo.Clear
    cboEspecificacao.Clear
    
    If chkCaract.Value = 1 Then
        Exit Sub
    End If
    
    If Len(Trim(cboDivisao.Text)) > 0 Then
        Call suListarTipo(Left(cboDivisao.Text, 4))
    End If
End Sub

Private Sub cboEspecificacao_Click()
    
    pgbProgresso.Value = 0
    
    
End Sub

Private Sub cboStatus_Click()
    
    If cboStatus.Text = "Relatório Semanal" Then
        cboRelOS.ListIndex = 2
    Else
       
        cboRelOS.ListIndex = 0
    End If
    
    pgbProgresso.Value = 0
End Sub

Private Sub cboTipo_Click()
    
    
    pgbProgresso.Value = 0
    cboEspecificacao.Clear
    
    If chkEspec.Value = 1 Then
        Exit Sub
    End If
    
    If Len(Trim(cboTipo.Text)) > 0 Then
        Call suListarEspecificacao(Left(cboTipo.Text, 4))
    End If
End Sub

Private Sub cboUsuario_Click()
    
    pgbProgresso.Value = 0
    
   
End Sub

Private Sub chkCaract_Click()
    If chkCaract.Value = 1 Then
        cboTipo.Clear
    Else
        Call cboDivisao_Click
    End If
End Sub

Private Sub chkEspec_Click()
    If chkEspec.Value = 1 Then
        cboEspecificacao.Clear
    Else
        Call cboTipo_Click
    End If
End Sub

Private Sub chkStatus_Click()
    If chkStatus.Value = 1 Then
        cboStatus.Clear
    Else
        Call suListarStatus
    End If
End Sub

Private Sub chkTipo_Click()
    If chkTipo.Value = 1 Then
        cboDivisao.Clear
    Else
        Call suListarDivisao
    End If
End Sub

Private Sub chkUsuario_Click()
    If chkUsuario.Value = 1 Then
        cboUsuario.Clear
    Else
        Call suListarUsuarios
    End If
End Sub

Private Sub cmdExportar_Click()
    On Error GoTo Erro

    ' Configurações do CommonDialog para abrir o explorador de arquivos
    With CommonDialog1
        .CancelError = True
        .Filter = "Pasta de Trabalho do Microsoft Office Excel (*.xls)|*.xls|All Files (*.*)|*.*" 'Excel Files
        .FilterIndex = 1
        .Flags = cdlOFNOverwritePrompt
        On Error Resume Next
        .ShowSave
        If Err.Number <> 0 Then
            MsgBox "Cancelado", vbInformation, "Suporte Manutenção"
            Err.Clear
            Exit Sub
        End If
    End With

    ' Captura o caminho do arquivo selecionado pelo usuário
    Dim filePath As String
    filePath = CommonDialog1.filename

    ' Verifica se o usuário selecionou um arquivo
    If filePath <> "" Then
        ' Chama a função para salvar o ListView no Excel
        Me.MousePointer = vbHourglass
        Call SaveListViewToExcel(filePath)
         Me.MousePointer = vbDefault
    End If
    Exit Sub

Erro:
    MsgBox "Erro: " & Err.Description, vbCritical, "Suporte Manutenção"
    On Error Resume Next
End Sub

Private Sub SaveListViewToExcel(filePath As String)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim j As Integer
    Dim totalItems As Integer
    Dim currentItem As Integer
    ' Cria uma nova instância do Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    ' Copia os cabeçalhos das colunas do ListView para a planilha
    For i = 1 To Me.lvwRelatorio.ColumnHeaders.Count
        xlSheet.Cells(1, i).Value = Me.lvwRelatorio.ColumnHeaders(i).Text
    Next i
    ' Formata os cabeçalhos em negrito e centralizado
    With xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, Me.lvwRelatorio.ColumnHeaders.Count))
        .Font.Bold = True
        .HorizontalAlignment = -4108 ' xlCenter
        .Interior.Color = RGB(222, 212, 27)
        .WrapText = True
    End With
    ' Obtém o total de itens para a barra de progresso
    totalItems = Me.lvwRelatorio.ListItems.Count

    ' Copia os itens do ListView para a planilha
    For i = 1 To totalItems
        xlSheet.Cells(i + 1, 1).Value = Me.lvwRelatorio.ListItems(i).Text ' Primeiro subitem
        For j = 1 To Me.lvwRelatorio.ListItems(i).ListSubItems.Count
            xlSheet.Cells(i + 1, j + 1).Value = Me.lvwRelatorio.ListItems(i).SubItems(j)
        Next j
    Next i
    
    ' Ajusta a largura das colunas automaticamente
    xlSheet.Columns.AutoFit
    ' Salva o arquivo Excel
    xlBook.SaveAs filePath
    
    xlBook.Close False
    xlApp.Quit
    
    ' Libera a memória
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    ' Exibe a mensagem de sucesso
    MsgBox "Os dados foram salvos com sucesso em " & filePath, vbInformation, "Suporte Manutenção"
End Sub


Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdGerar_Click()
    
    If Len(cboDivisao.Text) = 0 And Len(cboTipo.Text) = 0 And Len(cboEspecificacao.Text) = 0 And Len(cboUsuario.Text) = 0 And Len(cboStatus.Text) = 0 And Len(txtDataCadDe.Text) = 0 And Len(txtDataCadAte.Text) = 0 Then
        MsgBox "Nenhum critério foi informado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    If IsDate(txtDataCadDe.Text) = True And IsDate(txtDataCadAte.Text) = True Then
        If CDate(txtDataCadDe.Text) > CDate(txtDataCadAte.Text) Then
            MsgBox "DATA INICIAL não pode ser maior que a DATA FINAL!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtDataCadDe.Text = ""
            txtDataCadAte.Text = ""
            txtDataCadDe.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(txtDataCadDe.Text) > 0 Then
        If IsDate(txtDataCadDe.Text) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtDataCadDe.Text = ""
            txtDataCadDe.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(txtDataCadAte.Text) > 0 Then
        If IsDate(txtDataCadAte.Text) = False Then
            MsgBox "Data inválida!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtDataCadAte.Text = ""
            txtDataCadAte.SetFocus
            Exit Sub
        End If
    End If
        
    If cboRelOS.Text = "" Then
        MsgBox "Selecione o tipo de relatório!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboRelOS.SetFocus
        Exit Sub
    End If
    
    If cboRelOS.Text = "Semanal" Then
        Call suGerarRelatorioDetalhado(Left(cboDivisao.Text, 4), Left(cboTipo.Text, 4), Left(cboEspecificacao.Text, 4), Left(cboUsuario.Text, 4), cboStatus.Text, txtDataCadDe.Text, txtDataCadAte.Text)
    Else
        MsgBox "Relatório Indisponivel!!", vbOKOnly + vbInformation, "Suporte Técnico"
        'Call suGerarRelatorioLista(Left(cboDivisao.Text, 4), Left(cboTipo.Text, 4), Left(cboEspecificacao.Text, 4), Left(cboUsuario.Text, 4), cboStatus.Text, txtDataCadDe.Text, txtDataCadAte.Text)
    End If
End Sub

Private Sub Form_Load()

With Me.lvwRelatorio
    .ListItems.Clear
    .ColumnHeaders.Add , , "SOLICITANTE", 2000     '0
    .ColumnHeaders.Add , , "DEPTO.", 2000    '1
    .ColumnHeaders.Add , , "OS", 2000
    .ColumnHeaders.Add , , "STATUS", 2000
    .ColumnHeaders.Add , , "TIPO", 2000
    .ColumnHeaders.Add , , "CARACTERISTICA", 2000
    .ColumnHeaders.Add , , "ESPECIFICAÇÃO", 2000
    .ColumnHeaders.Add , , "OBSERVAÇÃO", 2000
    .ColumnHeaders.Add , , "DATA CADASTRO", 2000
    .ColumnHeaders.Add , , "NECESSIDADE", 2000
    .ColumnHeaders.Add , , "PREV. SISTEMAS", 2000
    .ColumnHeaders.Add , , "ATENDENTE", 2000
    .ColumnHeaders.Add , , "DATA FINAL.", 2000
    .ColumnHeaders.Add , , "PRAZO", 2000
    .ColumnHeaders.Add , , "REPORTE TECNICO", 2000
    .ColumnHeaders.Add , , "DATA ACEITE", 2000
    .ColumnHeaders.Add , , "COMENTÁRIO", 2000
    .ColumnHeaders.Add , , "DATA CANCEL.", 2000
    .ColumnHeaders.Add , , "DATA OS NÃO VALIDADA", 2000
    .ColumnHeaders.Add , , "MOTIVO OS NÃO VALIDADA", 2000
    .ColumnHeaders.Add , , "SITUAÇÃO", 2000
    .ColumnHeaders.Add , , "DATA SITUAÇÃO", 2000
    .ColumnHeaders.Add , , "ATENDIDO", 2000
   

End With


    Call suListarDivisao
    Call suListarStatus
    Call suListarUsuarios
End Sub

Private Sub suGerarRelatorioLista(ByVal vTipoID As String, ByVal vCaractID As String, ByVal vEspecID As String, ByVal vUsuarioID As String, ByVal vStatus As String, ByVal vDataCadDe As Variant, ByVal vDataCadAte As Variant)

End Sub


Private Sub suGerarRelatorioDetalhado(ByVal vTipoID As String, ByVal vCaractID As String, ByVal vEspecID As String, ByVal vUsuarioID As String, ByVal vStatus As String, ByVal vDataCadDe As Variant, ByVal vDataCadAte As Variant)
Me.lvwRelatorio.ListItems.Clear
Me.pgbProgresso.Value = 0
On Error GoTo Erro

Dim blnWhere As Boolean
Dim intStatus As Integer
Dim blnPrioridade As Boolean
Dim strStatus As String
Dim strPrazo As String
Dim blnRelSemanal As Boolean
    
    blnPrioridade = False
    blnRelSemanal = False
    
    strSQL = "SELECT * FROM vw_Chamados"
    blnWhere = False
    
    If Len(vTipoID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DivisaoID = " & CInt(vTipoID) & " "
        blnWhere = True
    End If
            
    If Len(vCaractID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "TipoID = " & CInt(vCaractID) & " "
        blnWhere = True
    End If

    If Len(vEspecID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "EspecificacaoID = " & CInt(vEspecID) & " "
        blnWhere = True
    End If
    
    If Len(vUsuarioID) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "UsuarioID = " & CInt(vUsuarioID) & " "
        blnWhere = True
    End If
    
    If Len(vStatus) > 0 Then

        Select Case vStatus
            Case Is = "Em Aberto"
                intStatus = 0
            Case Is = "Urgente"
                blnPrioridade = True
            Case Is = "Em Análise"
                intStatus = 7
            Case Is = "Em Atendimento"
                intStatus = 1
            Case Is = "Aguardando Aceite"
                intStatus = 2
            Case Is = "Finalizada"
                intStatus = 3
            Case Is = "Cancelada"
                intStatus = 4
            Case Is = "Não Validada"
                intStatus = 6
            Case Is = "Relatório Semanal"
                blnRelSemanal = True
        End Select
                    
        If blnPrioridade = True Then
            strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Prioridade = '" & IIf(blnPrioridade = True, 1, 0) & "' "
        Else
            If blnRelSemanal = False Then
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Status = " & intStatus & " "
            Else
                strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Status IN (0,7,1,6) "
            End If
        End If
        
        blnWhere = True
    End If
    
    If Len(vDataCadDe) > 0 And Len(vDataCadAte) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro >= '" & Format(vDataCadDe, "yyyy-MM-dd") & " 00:00:00' AND DataCadastro <= '" & Format(vDataCadAte, "yyyy-MM-dd") & " 23:59:59' "
        blnWhere = True
    ElseIf Len(vDataCadDe) > 0 And Len(vDataCadAte) = 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro BETWEEN '" & Format(vDataCadDe, "yyyy-MM-dd") & " 00:00:00' AND '" & Format(vDataCadDe, "yyyy-MM-dd") & " 23:59:59' "
        blnWhere = True
    ElseIf Len(vDataCadDe) = 0 And Len(vDataCadAte) > 0 Then
        strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "DataCadastro <= '" & Format(vDataCadAte, "yyyy-MM-dd") & " 23:59:59' "
        blnWhere = True
    End If
    
    strSQL = strSQL & "ORDER BY Nome,OSID"
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        pgbProgresso.Min = 0
        pgbProgresso.Value = 0
        pgbProgresso.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
        Screen.MousePointer = 11
    Else
        MsgBox "Nenhuma OS foi localizada!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    Do While Not rs.EOF
        Dim itmX As ListItem
        Set itmX = lvwRelatorio.ListItems.Add(, , rs!Nome & "")    'SOLICITANTE
        itmX.SubItems(1) = rs!Departamento & ""                     'DEPARTAMENTO
        itmX.SubItems(2) = CStr(Format(rs!OSID, "0000")) & ""       'OS
        
        Select Case rs!Status
            Case 0
                strStatus = "Em Aberto"
            Case 1
                strStatus = "Em Atendimento"
            Case 2
                strStatus = "Aguardando Aceite"
            Case 3
                strStatus = "OS Finalizada"
            Case 4
                strStatus = "OS Cancelada"
            Case 6
                strStatus = "Não Validada"
        End Select
        
        itmX.SubItems(3) = strStatus & ""           'STATUS
        itmX.SubItems(4) = rs!Divisao & ""          'TIPO
        itmX.SubItems(5) = rs!Tipo & ""             'caracteristica
        itmX.SubItems(6) = rs!Especificacao & ""       'especificação
        itmX.SubItems(7) = rs!DescricaoServico & ""    'observação
        itmX.SubItems(8) = Format(rs!Datacadastro, "dd/MM/yy HH:mm")  'data cadastro
        itmX.SubItems(9) = Format(rs!Previsao, "dd/MM/yy HH:mm")            'necessidade
        itmX.SubItems(10) = IIf(IsNull(rs!PrevisaoSistemas), "", Format(rs!PrevisaoSistemas, "dd/MM/yy HH:mm")) 'PREV. SISTEMA
        If IsNull(rs!Atendente) Then
            itmX.SubItems(11) = ""
        Else
            itmX.SubItems(11) = rs!Atendente   'ATENDENTE
        End If
        
        itmX.SubItems(12) = IIf(IsNull(rs!DataBaixa), "", Format(rs!DataBaixa, "dd/MM/yy HH:mm"))       'DATA FINAL.
        
        Select Case rs!Prazo
            Case 0
                strPrazo = "No prazo"
            Case 1
                strPrazo = "Atrasado"
        End Select
        
        itmX.SubItems(13) = strPrazo                                                                                'PRAZO
        If IsNull(rs!ReporteTecnico) Then
            itmX.SubItems(14) = ""
        Else
            itmX.SubItems(14) = rs!ReporteTecnico                                                  'REPORTE TECNICO
        End If
        
        itmX.SubItems(15) = IIf(IsNull(rs!DataAceite), "", Format(rs!DataAceite, "dd/MM/yy HH:mm"))                     'DATA ACEITE
        
        If IsNull(rs!Comentario) Then
            itmX.SubItems(16) = ""
        Else
            itmX.SubItems(16) = rs!Comentario                                                        'COMENTARIO
        End If
        
        itmX.SubItems(17) = IIf(IsNull(rs!DataCancelamento), "", Format(rs!DataCancelamento, "dd/MM/yy HH:mm")) ' DATA CANCELAMENTO
        itmX.SubItems(18) = IIf(IsNull(rs!DataOSNaoValidada), "", Format(rs!DataOSNaoValidada, "dd/MM/yy HH:mm")) ' DATA OS NÃO VALIDADA
        
        If IsNull(rs!MotivoOSNaoValidada) Then
            itmX.SubItems(19) = ""
        Else
            itmX.SubItems(19) = rs!MotivoOSNaoValidada                                                        'COMENTARIO
        End If
        
        If IsNull(rs!Situacao) Then
            itmX.SubItems(20) = ""
        Else
            itmX.SubItems(20) = rs!Situacao                                                        'COMENTARIO
        End If
        
        itmX.SubItems(21) = IIf(IsNull(rs!DataSituacao), "", Format(rs!DataSituacao, "dd/MM/yy HH:mm")) ' DATA SITUACAO
                                                                                                        'ATENDIDO
        
               
            pgbProgresso.Value = rs.AbsolutePosition
            rs.MoveNext
        Loop
        
        
    
    
   rs.Close
    Set rs = Nothing
    Screen.MousePointer = 0
    MsgBox "Relatório gerado com sucesso!", vbOKOnly + vbInformation, "Suporte Tecnico"
    Exit Sub
    
Erro:
   Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Erro"
    On Error Resume Next
    'rs.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub



Private Function fnPrazo(ByVal vDataFinal As String, ByVal vDataPrev As String) As String
    If vDataFinal = "" And vDataPrev = "" Then
        Exit Function
    End If

    If vDataFinal = "" And CDate(vDataPrev) < Format(Now, "dd/MM/yyyy") Then
        fnPrazo = "Atrasado"
    ElseIf vDataFinal = "" And CDate(vDataPrev) >= Format(Now, "dd/MM/yyyy") Then
        fnPrazo = "No Prazo"
    ElseIf CDate(vDataFinal) <= CDate(vDataPrev) Then
        fnPrazo = "No Prazo"
    ElseIf CDate(vDataFinal) > CDate(vDataPrev) Then
        fnPrazo = "Atrasado"
    End If
End Function

Private Sub suListarUsuarios()
    strSQL = "SELECT UsuarioID,Nome FROM vw_Usuarios ORDER BY Nome"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboUsuario.AddItem Format(rs!usuarioID, "0000") & " - " & rs!Nome
        rs.MoveNext
    Loop
    
    Set rs = Nothing
End Sub

Private Sub suListarEspecificacao(ByVal vTipoID As Integer)
    strSQL = "SELECT * FROM tb_Especificacoes WHERE TipoID = " & vTipoID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboEspecificacao.AddItem Format(rs!EspecificacaoID, "0000") & " - " & rs!Especificacao
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarTipo(ByVal vDivisaoID As Integer)
    strSQL = "SELECT * FROM tb_Tipos WHERE DivisaoID = " & vDivisaoID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboTipo.AddItem Format(rs!TipoID, "0000") & " - " & rs!Tipo
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarDivisao()
Call ConectarBD
    strSQL = "SELECT * FROM tb_Divisao ORDER BY DivisaoID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboDivisao.AddItem Format(rs!DivisaoID, "0000") & " - " & rs!Divisao
        rs.MoveNext
    Loop
    
    rs.Close
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
    cboStatus.AddItem "Relatório Semanal"
End Sub




