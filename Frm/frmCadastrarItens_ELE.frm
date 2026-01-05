VERSION 5.00
Begin VB.Form frmCadastrarItens 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Cadastrar Itens"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadastrarItens_ELE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cboEspecificacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4335
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboDivisao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblEspecificacao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Especificação"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Característica"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmCadastrarItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private intDivisaoID As Integer
Private intTipoID As Integer
Private intEspecificacaoID As Integer

Private Sub Label1_Click()

End Sub

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

Private Sub cmdCadastrar_Click()

    If cboDivisao.Text = "" And cboTipo.Text = "" And cboEspecificacao.Text = "" Then
        MsgBox "Nenhum campo foi preenchido!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    
'    If Len(Trim(cboDivisao.Text)) > 0 Then
        If IsNumeric(Left(Trim(cboDivisao.Text), 4)) = True Then
            intDivisaoID = Left(cboDivisao.Text, 4)
        Else
            intDivisaoID = 0
        End If
'    Else
'        MsgBox "Selecione ou digite uma nova divisão!", vbOKOnly + vbExclamation, "Suporte Técnico"
'        cboDivisao.SetFocus
'        Exit Sub
'    End If
    
'    If Len(Trim(cboTipo.Text)) > 0 Then
        If IsNumeric(Left(Trim(cboTipo.Text), 4)) = True Then
            intTipoID = Left(cboTipo.Text, 4)
        Else
            intTipoID = 0
        End If
'    Else
'        MsgBox "Selecione ou digite um novo tipo!", vbOKOnly + vbExclamation, "Suporte Técnico"
'        cboTipo.SetFocus
'        Exit Sub
'    End If
    
'    If Len(Trim(cboEspecificacao.Text)) > 0 Then
        If IsNumeric(Left(Trim(cboEspecificacao.Text), 4)) = True Then
            intEspecificacaoID = Left(cboEspecificacao.Text, 4)
        Else
            intEspecificacaoID = 0
        End If
'    Else
'        MsgBox "Selecione ou digite uma nova especificação!", vbOKOnly + vbExclamation, "Suporte Técnico"
'        cboEspecificacao.SetFocus
'        Exit Sub
'    End If
    
    If intDivisaoID > 0 And intTipoID > 0 And intEspecificacaoID > 0 Then
        MsgBox "Item já cadastrado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    ElseIf intDivisaoID = 0 And intTipoID = 0 And intEspecificacaoID = 0 Then
        Call suCadastrarDivisao(Trim(cboDivisao.Text))
    ElseIf intDivisaoID > 0 And intTipoID = 0 And intEspecificacaoID = 0 Then
        Call suCadastrarTipo(Left(Trim(cboDivisao.Text), 4), Trim(cboTipo.Text))
    ElseIf intDivisaoID > 0 And intTipoID > 0 And intEspecificacaoID = 0 Then
        Call suCadastrarEspecificacao(Left(Trim(cboTipo.Text), 4), Trim(cboEspecificacao.Text))
    End If
    
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Call suListarDivisoes
End Sub

Private Sub suCadastrarEspecificacao(ByVal vTipoID As Integer, ByVal vEspecificacao As String)
    strSQL = "SELECT * FROM tb_Especificacoes WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!Especificacao = vEspecificacao & ""
        rs!TipoID = vTipoID
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suCadastrarTipo(ByVal vDivisaoID As Integer, ByVal vTipo As String)
    strSQL = "SELECT * FROM tb_Tipos WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!Tipo = vTipo & ""
        rs!DivisaoID = vDivisaoID
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suCadastrarDivisao(ByVal vDivisao As String)
    strSQL = "SELECT * FROM tb_Divisao WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!Divisao = vDivisao & ""
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarDivisoes()
Call ConectarBD
    strSQL = "SELECT * FROM tb_Divisao WHERE Inativo=0 ORDER BY DivisaoID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboDivisao.AddItem Format(rs!DivisaoID, "0000") & " - " & rs!Divisao
        rs.MoveNext
    Loop
    
    Set rs = Nothing
End Sub

Private Sub suListarTipo(ByVal vDivisaoID As Integer)
    strSQL = "SELECT * FROM tb_Tipos WHERE DivisaoID = " & vDivisaoID & " AND Inativo=0"
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
    strSQL = "SELECT * FROM tb_Especificacoes WHERE TipoID = " & vTipoID & " AND Inativo=0"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboEspecificacao.AddItem Format(rs!EspecificacaoID, "0000") & " - " & rs!Especificacao
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

