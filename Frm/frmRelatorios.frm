VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmRelatorios 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Relatórios"
   ClientHeight    =   5760
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
   Icon            =   "frmRelatorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboRelOS 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmRelatorios.frx":030A
      Left            =   3825
      List            =   "frmRelatorios.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5355
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
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1020
      Width           =   2295
   End
   Begin VB.TextBox txtDataCadAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox txtDataCadDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   1035
      Width           =   1215
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cdlSalvarArquivo 
      Left            =   8160
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pasta de Trabalho do Microsoft Office Excel (*.xls) | *.xls"
   End
   Begin MSComctlLib.ProgressBar pgbProgresso 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   7695
      _ExtentX        =   13573
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
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "&Gerar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1020
      Width           =   2295
   End
   Begin FPSpreadADO.fpSpread fpsRelatorio 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   7695
      _Version        =   458752
      _ExtentX        =   13573
      _ExtentY        =   5741
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxRows         =   65000
      OperationMode   =   1
      SpreadDesigner  =   "frmRelatorios.frx":032D
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
      TabIndex        =   25
      Top             =   5400
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
    With fpsRelatorio
        .Reset
    End With
    
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
    
    With fpsRelatorio
        .Reset
    End With
End Sub

Private Sub cboStatus_Click()
    
    If cboStatus.Text = "Relatório Semanal" Then
        cboRelOS.ListIndex = 2
    Else
        With fpsRelatorio
            .Reset
        End With
        cboRelOS.ListIndex = 0
    End If
    
    pgbProgresso.Value = 0
End Sub

Private Sub cboTipo_Click()
    With fpsRelatorio
        .Reset
    End With
    
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
    
    With fpsRelatorio
        .Reset
    End With
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
Dim blnExportar As Boolean
    
    If fpsRelatorio.SheetName = "Relatório OS" Then
        
        cdlSalvarArquivo.FileName = ""
        
        Call cdlSalvarArquivo.ShowSave
        
        If cdlSalvarArquivo.FileName <> "" Then
            fpsRelatorio.Protect = False
            blnExportar = fpsRelatorio.ExportToExcel(cdlSalvarArquivo.FileName, fpsRelatorio.SheetName, "")
            fpsRelatorio.Protect = True
        Else
            Exit Sub
        End If
        
        If blnExportar = True Then
            MsgBox "Arquivo exportado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
        Else
            MsgBox "Erro ao exportar o arquivo!", vbOKOnly + vbCritical, "Suporte Técnico"
        End If
    Else
        MsgBox "Nenhum relatório foi gerado!", vbOKOnly + vbExclamation, "Suporte Técnico"
    End If
    
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
        Call suGerarRelatorioLista(Left(cboDivisao.Text, 4), Left(cboTipo.Text, 4), Left(cboEspecificacao.Text, 4), Left(cboUsuario.Text, 4), cboStatus.Text, txtDataCadDe.Text, txtDataCadAte.Text)
    End If
End Sub

Private Sub Form_Load()
    Call suListarDivisao
    Call suListarStatus
    Call suListarUsuarios
End Sub

Private Sub suGerarRelatorioLista(ByVal vTipoID As String, ByVal vCaractID As String, ByVal vEspecID As String, ByVal vUsuarioID As String, ByVal vStatus As String, ByVal vDataCadDe As Variant, ByVal vDataCadAte As Variant)
On Error GoTo Erro

Dim blnWhere As Boolean
Dim intStatus As Integer
Dim blnPrioridade As Boolean
Dim strStatus As String
    
    blnPrioridade = False
    
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
        End Select
                    
        If blnPrioridade = True Then
            strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Prioridade = '" & IIf(blnPrioridade = True, 1, 0) & "' "
        Else
            strSQL = strSQL & IIf(blnWhere = True, " AND ", " WHERE ") & "Status = " & intStatus & " "
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
    
    With fpsRelatorio
        .Reset
        .Row = 1
        .FontSize = 8
        .FontName = "Verdana"
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        .MaxRows = 65000
        
        .Col = 1
        .ColWidth(1) = 15
        .Text = "SOLICITANTE"
        
        .Col = 2
        .ColWidth(2) = 15
        .Text = "DEPTO."
        
        .Col = 3
        .ColWidth(3) = 5
        .Text = "OS"
        
        .Col = 4
        .ColWidth(4) = 15
        .Text = "STATUS"
        
        .Col = 5
        .ColWidth(5) = 15
        .Text = "TIPO"
        
        .Col = 6
        .ColWidth(6) = 15
        .Text = "CARACTERÍSTICA"
        
        .Col = 7
        .ColWidth(7) = 30
        .Text = "ESPECIFICAÇÃO"
        
        .Col = 8
        .ColWidth(8) = 30
        .Text = "OBSERVAÇÃO"
        
        .Col = 9
        .ColWidth(9) = 15
        .Text = "DATA CADASTRO"
        
        .Col = 10
        .ColWidth(10) = 15
        .Text = "NECESSIDADE"
        
        .Col = 11
        .ColWidth(11) = 15
        .Text = "PREV. SISTEMAS"
        
        .Col = 12
        .ColWidth(12) = 15
        .Text = "ATENDENTE"
        
        .Col = 13
        .ColWidth(13) = 15
        .Text = "DATA FINAL."
        
        .Col = 14
        .ColWidth(14) = 15
        .Text = "PRAZO"

        .Col = 15
        .ColWidth(15) = 30
        .Text = "REPORTE TÉCNICO"
    
        .Col = 16
        .ColWidth(16) = 15
        .Text = "DATA ACEITE"
                
        .Col = 17
        .ColWidth(17) = 30
        .Text = "COMENTÁRIO"
    
        .Col = 18
        .ColWidth(18) = 15
        .Text = "DATA CANCEL."

        .Col = 19
        .ColWidth(19) = 15
        .Text = "DATA OS NÃO VALIDADA"

        .Col = 20
        .ColWidth(20) = 30
        .Text = "MOTIVO OS NÃO VALIDADA"
        
        .Col = 21
        .ColWidth(21) = 30
        .Text = "SITUAÇÃO"

        .Col = 22
        .ColWidth(22) = 15
        .Text = "DATA SITUAÇÃO"
        
        .Col = 23
        .ColWidth(23) = 15
        .Text = "ATENDIDO"
        
        Do While Not rs.EOF
        DoEvents
            .Row = .Row + 1

            .Col = 1
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Nome & ""
            
            .Col = 2
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Departamento & ""
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = CStr(Format(rs!OSID, "0000")) & ""
            
            Select Case rs!Status
                Case 0
                    strStatus = "Em Aberto"
                Case 7
                    strStatus = "Em Análise"
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
            
            .Col = 4
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = strStatus & ""
            
            .Col = 5
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Divisao & ""
            
            .Col = 6
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Tipo & ""
            
            .Col = 7
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Especificacao & ""
            
            .Col = 8
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!DescricaoServico & ""
            
            .Col = 9
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCadastro, "dd/MM/yy HH:mm")
            
            .Col = 10
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!Previsao, "dd/MM/yyyy")
                    
            .Col = 11
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!PrevisaoSistemas, "dd/MM/yyyy")
            
            .Col = 12
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Atendente & ""
        
            .Col = 13
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            
            .Col = 14
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = fnPrazo(IIf(IsNull(Format(rs!DataBaixa, "dd/MM/yyyy")) = True, "", Format(rs!DataBaixa, "dd/MM/yyyy")), IIf(IsNull(rs!PrevisaoSistemas) = True, "", rs!PrevisaoSistemas))
            
            .Col = 15
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!ReporteTecnico & ""
            
            .Col = 16
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataAceite, "dd/MM/yy HH:mm")
            
            .Col = 17
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Comentario & ""
            
            .Col = 18
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCancelamento, "dd/MM/yy HH:mm")
            
            .Col = 19
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataOSNaoValidada, "dd/MM/yy HH:mm")
            
            .Col = 20
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!MotivoOSNaoValidada & ""
            
            .Col = 21
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Situacao & ""
            
            .Col = 22
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataSituacao, "dd/MM/yy HH:mm")
            
            .Col = 23
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = IIf(IsNull(rs!DataBaixa) = False, 1, 0)
        
            pgbProgresso.Value = rs.AbsolutePosition
            rs.MoveNext
        Loop
        
        .SheetName = "Relatório OS"
        .OperationMode = OperationModeRead
        
    End With
    
    Set rs = Nothing
    MsgBox "Relatório gerado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
    Screen.MousePointer = 0
End Sub


Private Sub suGerarRelatorioDetalhado(ByVal vTipoID As String, ByVal vCaractID As String, ByVal vEspecID As String, ByVal vUsuarioID As String, ByVal vStatus As String, ByVal vDataCadDe As Variant, ByVal vDataCadAte As Variant)
On Error GoTo Erro

Dim blnWhere As Boolean
Dim intStatus As Integer
Dim blnPrioridade As Boolean
Dim strStatus As String
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
    
    With fpsRelatorio
        .Reset
        .Row = 1
        .FontSize = 8
        .FontName = "Verdana"
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        .MaxRows = 65000
        
        .Col = 1
        .ColWidth(1) = 15
        .Text = "SOLICITANTE"
        
        .Col = 2
        .ColWidth(2) = 15
        .Text = "DEPTO."
        
        .Col = 3
        .ColWidth(3) = 5
        .Text = "OS"
        
        .Col = 4
        .ColWidth(4) = 15
        .Text = "STATUS"
        
        .Col = 5
        .ColWidth(5) = 15
        .Text = "TIPO"
        
        .Col = 6
        .ColWidth(6) = 15
        .Text = "CARACTERÍSTICA"
        
        .Col = 7
        .ColWidth(7) = 30
        .Text = "ESPECIFICAÇÃO"
        
        .Col = 8
        .ColWidth(8) = 30
        .Text = "OBSERVAÇÃO"
        
        .Col = 9
        .ColWidth(9) = 15
        .Text = "DATA CADASTRO"
        
        .Col = 10
        .ColWidth(10) = 15
        .Text = "NECESSIDADE"
        
        .Col = 11
        .ColWidth(11) = 15
        .Text = "PREV. SISTEMAS"
        
        .Col = 12
        .ColWidth(12) = 15
        .Text = "ATENDENTE"
        
        .Col = 13
        .ColWidth(13) = 15
        .Text = "DATA FINAL."
        
        .Col = 14
        .ColWidth(14) = 15
        .Text = "PRAZO"

        .Col = 15
        .ColWidth(15) = 30
        .Text = "REPORTE TÉCNICO"
    
        .Col = 16
        .ColWidth(16) = 15
        .Text = "DATA ACEITE"
                
        .Col = 17
        .ColWidth(17) = 30
        .Text = "COMENTÁRIO"
    
        .Col = 18
        .ColWidth(18) = 15
        .Text = "DATA CANCEL."

        .Col = 19
        .ColWidth(19) = 15
        .Text = "DATA OS NÃO VALIDADA"

        .Col = 20
        .ColWidth(20) = 30
        .Text = "MOTIVO OS NÃO VALIDADA"
        
        .Col = 21
        .ColWidth(21) = 30
        .Text = "SITUAÇÃO"

        .Col = 22
        .ColWidth(22) = 15
        .Text = "DATA SITUAÇÃO"
        
        .Col = 23
        .ColWidth(23) = 15
        .Text = "ATENDIDO"
        
        Call suColorirSubTotais(.Row)
        
        Do While Not rs.EOF
        DoEvents
            .Row = .Row + 1
            
            .Col = 1
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Nome & ""
            
            Dim tmpNome As String
            Dim tmpNomeAnt As String
            Dim intCont As Integer
            
            If .Row >= 3 Then
                tmpNomeAnt = .Text
                If tmpNome <> tmpNomeAnt Then
                    .Col = 1
                    .FontSize = 8
                    .FontName = "Verdana"
                    .FontBold = True
                    .TypeHAlign = TypeHAlignLeft
                    .Text = tmpNome
                    
                    .Col = 3
                    .FontSize = 8
                    .FontName = "Verdana"
                    .FontBold = True
                    .TypeHAlign = TypeHAlignRight
                    .CellType = CellTypeNumber
                    .TypeNumberDecPlaces = 0
                    .Text = intCont
                            
                    Call suColorirSubTotais(.Row)
                                    
                End If
                
                If .Text <> tmpNomeAnt Then
                    intCont = 0
                    .Row = .Row + 1
                    .Col = 1
                    .FontSize = 8
                    .FontName = "Verdana"
                    .TypeHAlign = TypeHAlignLeft
                    .Text = tmpNomeAnt
                    tmpNome = .Text
                    intCont = intCont + 1
                Else
                    intCont = intCont + 1
                End If
            Else
                tmpNome = .Text
                intCont = intCont + 1
            End If
            
            .Col = 2
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Departamento & ""
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = CStr(Format(rs!OSID, "0000")) & ""
            
            Select Case rs!Status
                Case 0
                    strStatus = "Em Aberto"
                Case 7
                    strStatus = "Em Análise"
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
            
            .Col = 4
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = strStatus & ""
            
            .Col = 5
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Divisao & ""
            
            .Col = 6
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Tipo & ""
            
            .Col = 7
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Especificacao & ""
            
            .Col = 8
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!DescricaoServico & ""
            
            .Col = 9
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCadastro, "dd/MM/yy HH:mm")
            
            .Col = 10
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!Previsao, "dd/MM/yyyy")
                    
            .Col = 11
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!PrevisaoSistemas, "dd/MM/yyyy")
            
            .Col = 12
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Atendente & ""
        
            .Col = 13
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            
            .Col = 14
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = fnPrazo(IIf(IsNull(Format(rs!DataBaixa, "dd/MM/yyyy")) = True, "", Format(rs!DataBaixa, "dd/MM/yyyy")), IIf(IsNull(rs!PrevisaoSistemas) = True, "", rs!PrevisaoSistemas))
            
            .Col = 15
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!ReporteTecnico & ""
            
            .Col = 16
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataAceite, "dd/MM/yy HH:mm")
            
            .Col = 17
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Comentario & ""
            
            .Col = 18
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCancelamento, "dd/MM/yy HH:mm")
            
            .Col = 19
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataOSNaoValidada, "dd/MM/yy HH:mm")
            
            .Col = 20
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!MotivoOSNaoValidada & ""
            
            .Col = 21
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Situacao & ""
            
            .Col = 22
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataSituacao, "dd/MM/yy HH:mm")
            
            .Col = 23
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = IIf(IsNull(rs!DataBaixa) = False, 1, 0)
        
            pgbProgresso.Value = rs.AbsolutePosition
            rs.MoveNext
        Loop
        
        If rs.EOF = True Then
            .Col = 1
            .Row = .Row + 1
            .FontSize = 8
            .FontName = "Verdana"
            .FontBold = True
            .TypeHAlign = TypeHAlignLeft
            .Text = tmpNome
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .FontBold = True
            .TypeHAlign = TypeHAlignRight
            .CellType = CellTypeNumber
            .TypeNumberDecPlaces = 0
            .Text = intCont
            
             Call suColorirSubTotais(.Row)
        End If
        
        .SheetName = "Relatório OS"
        .OperationMode = OperationModeRead
        
    End With
    
    Set rs = Nothing
    MsgBox "Relatório gerado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub suColorirSubTotais(ByVal vRow As Long)
Dim lngRow As Long
    
    lngRow = vRow
    
    With fpsRelatorio
        .Col = 1
        Do While .Col <= 23
            .Row = lngRow
            .BackColor = &HC0FFFF
            .SetSelection 1, .Row, .Col, .Row
            .BackColor = &HC0FFFF
            .Col = .Col + 1
        Loop
    End With
    
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
        cboUsuario.AddItem Format(rs!UsuarioID, "0000") & " - " & rs!Nome
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

Private Sub suGerarRelatorio(ByVal vStatus As String)
On Error GoTo Erro

    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0 ORDER BY OSID"
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
    
    With fpsRelatorio
        .Reset
        .Row = 1
        .FontSize = 8
        .FontName = "Verdana"
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 1
        .ColWidth(1) = 15
        .Text = "SOLICITANTE"
        
        .Col = 2
        .ColWidth(2) = 5
        .Text = "OS"
        
        .Col = 3
        .ColWidth(3) = 15
        .Text = "TIPO"
        
        .Col = 4
        .ColWidth(4) = 15
        .Text = "CARACTERÍSTICA"
        
        .Col = 5
        .ColWidth(5) = 30
        .Text = "ESPECIFICAÇÃO"
        
        .Col = 6
        .ColWidth(6) = 30
        .Text = "OBSERVAÇÃO"
        
        .Col = 7
        .ColWidth(7) = 15
        .Text = "DATA CADASTRO"
        
        .Col = 8
        .ColWidth(8) = 15
        .Text = "NECESSIDADE"
        
        Select Case vStatus
            Case Is = "Em Atendimento"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "PREV. SISTEMAS"
                
                .Col = 10
                .ColWidth(10) = 15
                .Text = "ATENDENTE"
            Case Is = "Aguardando Aceite"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "PREV. SISTEMAS"
                
                .Col = 10
                .ColWidth(10) = 15
                .Text = "ATENDENTE"
                
                .Col = 11
                .ColWidth(11) = 15
                .Text = "DATA FINALIZAÇÃO"
            Case Is = "Finalizada"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "PREV. SISTEMAS"
                
                .Col = 10
                .ColWidth(10) = 15
                .Text = "ATENDENTE"
                
                .Col = 11
                .ColWidth(11) = 15
                .Text = "DATA FINAL."
            
                .Col = 12
                .ColWidth(12) = 15
                .Text = "DATA ACEITE"
            Case Is = "Cancelada"
                .Col = 9
                .ColWidth(9) = 15
                .Text = "DATA CANCEL."
        End Select
                
        Do While Not rs.EOF
        DoEvents
            .Row = .Row + 1
            .Col = 1
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Nome & ""
            
            .Col = 2
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignRight
            .Text = CStr(Format(rs!OSID, "0000")) & ""
            
            .Col = 3
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Divisao & ""
            
            .Col = 4
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Tipo & ""
            
            .Col = 5
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!Especificacao & ""
            
            .Col = 6
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignLeft
            .Text = rs!DescricaoServico & ""
            
            .Col = 7
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!DataCadastro, "dd/MM/yy HH:mm")
            
            .Col = 8
            .FontSize = 8
            .FontName = "Verdana"
            .TypeHAlign = TypeHAlignCenter
            .Text = Format(rs!Previsao, "dd/MM/yyyy")
                    
        Select Case vStatus
            Case Is = "Em Atendimento"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!PrevisaoSistemas, "dd/MM/yyyy")
                
                .Col = 10
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = rs!Atendente & ""
            Case Is = "Aguardando Aceite"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!PrevisaoSistemas, "dd/MM/yyyy")
                
                .Col = 10
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = rs!Atendente & ""
            
                .Col = 11
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            Case Is = "Finalizada"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!PrevisaoSistemas, "dd/MM/yyyy")
                
                .Col = 10
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignLeft
                .Text = rs!Atendente & ""
            
                .Col = 11
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!DataBaixa, "dd/MM/yy HH:mm")
            
                .Col = 12
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!DataAceite, "dd/MM/yy HH:mm")
            Case Is = "Cancelada"
                .Col = 9
                .FontSize = 8
                .FontName = "Verdana"
                .TypeHAlign = TypeHAlignCenter
                .Text = Format(rs!DataCancelamento, "dd/MM/yy HH:mm")
        End Select
            
            pgbProgresso.Value = rs.AbsolutePosition
            rs.MoveNext
        Loop
        
        .OperationMode = OperationModeRead
        
    End With
    
    Set rs = Nothing
    MsgBox "Relatório gerado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
    Set rs = Nothing
    Screen.MousePointer = 0
End Sub


