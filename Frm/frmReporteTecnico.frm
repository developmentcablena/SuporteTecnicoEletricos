VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteTecnico 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Reporte"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReporteTecnico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBaixar 
      Appearance      =   0  'Flat
      Caption         =   "&Finalizar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3780
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mskDataBaixa 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtReporteTecnico 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CommandButton cmdValidar 
      Appearance      =   0  'Flat
      Caption         =   "&Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3780
      Visible         =   0   'False
      Width           =   1335
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
      TabIndex        =   6
      Top             =   390
      Width           =   705
   End
   Begin VB.Label lblDataBaixa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data/ Hora Final."
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
      Top             =   3600
      Width           =   1695
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
      TabIndex        =   4
      Top             =   840
      Width           =   1635
   End
End
Attribute VB_Name = "frmReporteTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private strMotivoOSNaoValidada As String
Private strDataOSNaoValidada As String
Private strReporte As String
Private strDataReporte As String

Private Sub cmdBaixar_Click()
    If Len(Trim(txtOSID.Text)) = 0 Then
        MsgBox "Nenhuma OS foi selecionada!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    If Len(Trim(txtReporteTecnico.Text)) = 0 Then
        MsgBox "Digite o reporte técnico!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtReporteTecnico.SetFocus
        Exit Sub
    End If
    
    If IsDate(mskDataBaixa.Text) = False Then
        MsgBox "Digite a data/ hora de finalização do serviço!", vbOKOnly + vbExclamation, "Suporte Técnico"
        mskDataBaixa.SetFocus
        Exit Sub
    Else
        If CDate(mskDataBaixa.Text) > Now Then
            MsgBox "Data/ hora de finalização não pode ser maior que a data atual!", vbOKOnly + vbExclamation, "Suporte Técnico"
            mskDataBaixa.SetFocus
            Exit Sub
        End If
    End If
    
    Me.Enabled = False
    If fnPesquisarBaixa(Trim(txtOSID.Text)) = False Then
        Call suBaixarOS(gintOSID, Trim(txtReporteTecnico.Text), CDate(mskDataBaixa.Text), fnCapturarEMail(gintOSID), fnCapturarUsuario(gintOSID), fnCapturarEmailAtendente(gintOSID))
    Else
        Call suInserirReporteOcorrencia(gintOSID, strReporte, strDataReporte, strMotivoOSNaoValidada, strDataOSNaoValidada)
        Call suBaixarOS(gintOSID, Trim(txtReporteTecnico.Text), CDate(mskDataBaixa.Text), fnCapturarEMail(gintOSID), fnCapturarUsuario(gintOSID), fnCapturarEmailAtendente(gintOSID))
    End If
    Me.Enabled = True
    Call Unload(Me)
End Sub

Private Sub suInserirReporteOcorrencia(ByVal vOSID As Integer, ByVal vReporte As String, ByVal vDataReporte As String, ByVal vMotivo As String, ByVal vDataMotivo As String)
    strSQL = "INSERT INTO tb_Ocorrencias " & _
             "(ReporteTecnico,DataReporte,Ocorrencia,DataOcorrencia,Status,OSID) " & _
             "VALUES ('" & vReporte & "','" & Format(CDate(vDataReporte), "yyyy-MM-dd HH:mm") & "','" & vMotivo & "','" & Format(CDate(vDataMotivo), "yyyy-MM-dd HH:mm") & "',1," & vOSID & ")"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

'Private Sub suGravarReporteOcorrencia(ByVal vOSID As Integer, ByVal vReporte As String, ByVal vDataReporte As String, ByVal vMotivo As String, ByVal vDataMotivo As String)
'    strSQL = "UPDATE tb_Ocorrencias " & _
'             "SET ReporteTecnico = '" & vReporte & "',DataReporte = '" & CDate(vDataReporte) & "', " & _
'             "Ocorrencia= '" & vMotivo & "',DataOcorrencia = '" & vDataMotivo & "' " & _
'             "WHERE OSID = " & vOSID & " AND ReporteTecnico IS NULL AND DataReporte IS NULL"
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
'
'    Set rs = Nothing
'End Sub

Private Function fnPesquisarBaixa(ByVal vOSID As Integer) As Boolean
    
    strMotivoOSNaoValidada = ""
    strDataOSNaoValidada = ""
    strReporte = ""
    strDataReporte = ""
    fnPesquisarBaixa = False
    
    strSQL = "SELECT * FROM tb_OS WHERE OSID=" & vOSID & " AND DataBaixa IS NOT NULL"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strMotivoOSNaoValidada = rs!MotivoOSNaoValidada
        strDataOSNaoValidada = rs!DataOSNaoValidada
        strReporte = rs!ReporteTecnico
        strDataReporte = rs!DataBaixa
        fnPesquisarBaixa = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    With Me
        .txtOSID.Text = Format(gintOSID, "0000")
    End With
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
            fnCapturarEmailAtendente = rs1!Email & ""
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
        fnCapturarEMail = rs!Email & ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suBaixarOS(ByVal vOSID As Long, ByVal vReporteTecnico As String, ByVal vDataBaixa As Date, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEMailAtendente As String)
On Error GoTo Erro

    strSQL = "UPDATE tb_OS SET ReporteTecnico = '" & vReporteTecnico & "', DataBaixa = '" & Format(vDataBaixa, "yyyy-MM-dd HH:mm") & "', Status = 2 WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If fnEnviarEmail(vOSID, vDataBaixa, vReporteTecnico, vEMail, vUsuario, vEMailAtendente) = True Then
        MsgBox "OS finalizada com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
        Exit Sub
    Else
        MsgBox "Erro ao enviar e-mail!", vbOKOnly + vbCritical, "Suporte Técnico"
        Exit Sub
    End If
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
End Sub

Private Function fnEnviarEmail(ByVal vOSID As Integer, ByVal vDataBaixa As Date, ByVal vReporte As String, ByVal vEMail As String, ByVal vUsuario As String, ByVal vEMailAtendente As String) As Boolean
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
    poSendMail.Subject = "Finalização da OS " & Format(vOSID, "0000")
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
        strRelatorio = strRelatorio & "Reporte Técnico: " & rs!ReporteTecnico & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Data Finalização: " & Format(rs!DataBaixa, "dd/MM/yy HH:mm") & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "*ATENÇÃO: Favor registrar o ACEITE no sistema." & vbCrLf & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "**OBSERVAÇÃO: Esta mensagem é gerada automaticamente pelo sistema." & vbCrLf & "POR FAVOR NÃO RESPONDA ESTA MENSAGEM." & vbCrLf & vbCrLf & vbCrLf & "Suporte Técnico" & vbCrLf & "Cablena do Brasil" & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(100, "=")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub
