VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acrOS 
   Caption         =   "Suporte Técnico - Ordem de Serviço"
   ClientHeight    =   12090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   Icon            =   "acrOS_ELE.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21167
   _ExtentY        =   21325
   SectionData     =   "acrOS_ELE.dsx":076A
End
Attribute VB_Name = "acrOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private Sub suGerarOS(ByVal vOSID As Integer)
Call ConectarBD
    strSQL = "SELECT * FROM vw_Chamados WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        With acrOS
            .fldOSID = Format(rs!OSID, "0000")
            .fldDataCadastro.Text = Format(rs!Datacadastro, "dd/MM/yyyy HH:mm:ss")
            .fldDataPrevisao.Text = Format(rs!Previsao, "dd/MM/yyyy")
            .fldSolicitante.Text = rs!Nome
            .fldDepartamento.Text = rs!Departamento
            .fldPrioridade.Text = IIf(rs!Prioridade = True, "Sim", "Não")
            .fldDivisao.Text = rs!Divisao
            .fldTipo.Text = rs!Tipo
            .fldEspecificacao.Text = rs!Especificacao
            .fldDescricao.Text = rs!DescricaoServico
            .fldReporte.Text = rs!ReporteTecnico & ""
        End With
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ActiveReport_Activate()
    Call suGerarOS(gintOSID)
End Sub


