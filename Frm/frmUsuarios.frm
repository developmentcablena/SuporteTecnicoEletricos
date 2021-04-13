VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmUsuarios 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Usuários"
   ClientHeight    =   5745
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
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemover 
      Caption         =   "-"
      Height          =   300
      Left            =   6855
      TabIndex        =   26
      Top             =   3105
      Width           =   615
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "+"
      Height          =   300
      Left            =   6255
      TabIndex        =   25
      Top             =   3105
      Width           =   615
   End
   Begin VB.ComboBox cboModulos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame fraDadosUsuario 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   4260
      Begin VB.CheckBox chkEmBranco 
         Appearance      =   0  'Flat
         Caption         =   "Em branco"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkInativo 
         Appearance      =   0  'Flat
         Caption         =   "Inativo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2205
         TabIndex        =   19
         Top             =   1905
         Width           =   975
      End
      Begin VB.CheckBox chkAlterar 
         Appearance      =   0  'Flat
         Caption         =   "Alterar Senha"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1905
         Width           =   1575
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "&Salvar"
         Height          =   375
         Left            =   2190
         TabIndex        =   17
         Top             =   2235
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3165
         TabIndex        =   16
         Top             =   2235
         Width           =   975
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2235
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   1095
         TabIndex        =   14
         Top             =   2235
         Width           =   945
      End
      Begin VB.TextBox txtConfirmar 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2205
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1545
         Width           =   1935
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1545
         Width           =   1920
      End
      Begin VB.TextBox txtEMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         TabIndex        =   8
         Top             =   945
         Width           =   1935
      End
      Begin VB.TextBox txtDepto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   945
         Width           =   1935
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblConfirmar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Confirmar Senha"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2205
         TabIndex        =   13
         Top             =   1320
         Width           =   1470
      End
      Begin VB.Label lblSenha 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblEMail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2205
         TabIndex        =   9
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblDepto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Depto."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblUsuario 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2205
         TabIndex        =   5
         Top             =   135
         Width           =   645
      End
      Begin VB.Label lblNome 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   135
         Width           =   495
      End
   End
   Begin GridEX20.GridEX gexUsuarios 
      Height          =   4965
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8758
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmUsuarios.frx":08CA
      Column(2)       =   "frmUsuarios.frx":0A82
      Column(3)       =   "frmUsuarios.frx":0BCA
      FmtConditionsCount=   1
      FmtCondition(1) =   "frmUsuarios.frx":0D0A
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmUsuarios.frx":0ED6
      FormatStyle(2)  =   "frmUsuarios.frx":1002
      FormatStyle(3)  =   "frmUsuarios.frx":10B2
      FormatStyle(4)  =   "frmUsuarios.frx":1166
      FormatStyle(5)  =   "frmUsuarios.frx":123E
      ImageCount      =   0
      PrinterProperties=   "frmUsuarios.frx":12F6
   End
   Begin GridEX20.GridEX gexPermissoes 
      Height          =   1695
      Left            =   3240
      TabIndex        =   20
      Top             =   3480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmUsuarios.frx":14C6
      Column(2)       =   "frmUsuarios.frx":16B6
      Column(3)       =   "frmUsuarios.frx":182A
      Column(4)       =   "frmUsuarios.frx":19C2
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmUsuarios.frx":1B8A
      FormatStyle(2)  =   "frmUsuarios.frx":1CB6
      FormatStyle(3)  =   "frmUsuarios.frx":1D66
      FormatStyle(4)  =   "frmUsuarios.frx":1E1A
      FormatStyle(5)  =   "frmUsuarios.frx":1EF2
      ImageCount      =   0
      PrinterProperties=   "frmUsuarios.frx":1FAA
   End
   Begin VB.Label lblPermissoes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Permissões"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   24
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private Sub cmdAdicionar_Click()
    If Len(Trim(cboModulos.Text)) = 0 Then
        MsgBox "Selecione um módulo!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cboModulos.SetFocus
        Exit Sub
    End If
    
    If cmdNovo.Enabled = True Then
        Call suAdicionarModulo(gintNovoUsuarioID, cboModulos.Text)
        Call suMostrarPermissoes(gintNovoUsuarioID)
    ElseIf cmdEditar.Enabled = True Then
        Call suAdicionarModulo(gexUsuarios.Value(1), cboModulos.Text)
        Call suMostrarPermissoes(gexUsuarios.Value(1))
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call suLimparDados
    Call suLimparGridPermissoes
    Call suHabilitarDados(False)
    cmdNovo.Enabled = True
    cmdEditar.Enabled = True
    If cmdSalvar.Caption = "&Atualizar" Then
        cmdSalvar.Caption = "&Salvar"
    End If
End Sub

Private Sub cmdEditar_Click()
    If Len(Trim(txtNome.Text)) > 0 And Len(Trim(txtUsuario.Text)) > 0 Then
        Call suHabilitarDados(True)
        Call suHabilitarPermissoes(True)
        cmdNovo.Enabled = False
        cmdSalvar.Caption = "&Atualizar"
    Else
        MsgBox "Nenhum usuário foi selecionado!", vbOKOnly + vbExclamation, "Suporte Técnico"
        cmdNovo.Enabled = True
        cmdSalvar.Caption = "&Salvar"
    End If
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdNovo_Click()
    Call suLimparDados
    Call suLimparGridPermissoes
    Call suHabilitarDados(True)
    cmdEditar.Enabled = False
    txtNome.SetFocus
End Sub

Private Sub cmdRemover_Click()
Dim i As Integer
    
    If gexPermissoes.RowCount = 0 Then
        MsgBox "Usuário não possue nenhuma permissão!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    For i = 1 To gexPermissoes.RowCount
        If gexPermissoes.RowSelected(i) = True Then
            Call suRemoverModulo(gexPermissoes.Value(1), gexUsuarios.Value(1))
            Call suMostrarPermissoes(gexUsuarios.Value(1))
        End If
    Next
    
End Sub

Private Sub cmdSalvar_Click()
    If Len(Trim(txtNome.Text)) = 0 Then
        MsgBox "Digite o nome do usuário!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtNome.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtUsuario.Text)) = 0 Then
        MsgBox "Digite o login do usuário!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtUsuario.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtDepto.Text)) = 0 Then
        MsgBox "Digite o departamento do usuário!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtDepto.SetFocus
        Exit Sub
    End If
    
    If chkEmBranco.Value = 0 Then
        If Len(Trim(txtSenha.Text)) = 0 Then
            MsgBox "Digite uma senha!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtSenha.SetFocus
            Exit Sub
        End If
        
        If Len(Trim(txtConfirmar.Text)) = 0 Then
            MsgBox "Digite a confirmação da senha!", vbOKOnly + vbExclamation, "Suporte Técnico"
            txtConfirmar.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtSenha.Text) <> Trim(txtConfirmar.Text) Then
        MsgBox "Senhas não conferem!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtSenha.Text = ""
        txtConfirmar.Text = ""
        txtSenha.SetFocus
        Exit Sub
    Else
        If cmdSalvar.Caption = "&Salvar" Then
            If fnCadastrarUsuario(Trim(txtNome.Text), Trim(txtUsuario.Text), CodDec(Trim(txtSenha.Text)), Trim(txtDepto.Text), Trim(txtEMail.Text), chkInativo, chkAlterar) = True Then
                MsgBox "Usuário cadastrado com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
                Call suListarUsuarios
                Call suHabilitarDados(False)
                chkEmBranco.Value = 0
                Call suHabilitarPermissoes(True)
            End If
        Else
            If fnAtualizar(gexUsuarios.Value(1), Trim(txtDepto.Text), Trim(txtEMail.Text), CodDec(Trim(txtSenha.Text)), chkInativo, chkAlterar) = True Then
                MsgBox "Atualizações efetuadas com sucesso!", vbOKOnly + vbInformation, "Suporte Técnico"
                Call suListarUsuarios
                Call suHabilitarDados(False)
                Call suHabilitarPermissoes(False)
                chkEmBranco.Value = 0
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call suListarUsuarios
    Call suHabilitarDados(False)
    Call suHabilitarPermissoes(False)
    Call suListarModulos
End Sub

Private Sub gexUsuarios_Click()
Dim i As Integer

    If gexUsuarios.RowCount > 0 Then
        For i = 1 To gexUsuarios.RowCount
            If gexUsuarios.RowSelected(i) = True Then
                Call suLimparDados
                Call suLimparGridPermissoes
                Call suHabilitarDados(False)
                Call suHabilitarPermissoes(False)
                Call suMostrarDados(gexUsuarios.Value(1))
                Call suMostrarPermissoes(gexUsuarios.Value(1))
                cmdNovo.Enabled = True
                cmdEditar.Enabled = True
                cmdSalvar.Caption = "&Salvar"
            End If
        Next
    End If
    
End Sub

Private Sub suHabilitarPermissoes(ByVal vBoolean As Boolean)
    lblPermissoes.Enabled = vBoolean
    cboModulos.Enabled = vBoolean
    cmdAdicionar.Enabled = vBoolean
    cmdRemover.Enabled = vBoolean
    gexPermissoes.Enabled = vBoolean
End Sub

Private Sub suRemoverModulo(ByVal vPermissaoID As Integer, ByVal vUsuarioID As Integer)
    strSQL = "DELETE FROM tb_Permissoes WHERE PermissaoID = " & vPermissaoID & " AND UsuarioID = " & vUsuarioID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Sub suAdicionarModulo(ByVal vUsuarioID As Integer, ByVal vModulo As String)
    strSQL = "INSERT INTO tb_Permissoes " & _
             "(Modulo,Permissao,UsuarioID) " & _
             "VALUES ('" & vModulo & "',0," & vUsuarioID & ")"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub

Private Sub suListarModulos()
    strSQL = "SELECT * FROM vw_Modulos ORDER BY Modulo"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        cboModulos.AddItem rs!Modulo
        rs.MoveNext
    Loop
        
    rs.Close
    Set rs = Nothing
End Sub

Private Function fnAtualizar(ByVal vUsuarioID As Integer, ByVal vDepto As String, ByVal vEMail As String, ByVal vSenha As String, ByRef vInativo As CheckBox, ByVal vAlterarSenha As CheckBox) As Boolean
On Error GoTo Erro
    
    fnAtualizar = False
    
    strSQL = "UPDATE tb_Usuarios SET Departamento = '" & vDepto & "',EMail = '" & vEMail & "',Senha = '" & vSenha & "',Inativo = " & vInativo & ",AlterarSenha = " & vAlterarSenha & " WHERE UsuarioID = " & vUsuarioID & " "
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnAtualizar = True
    Set rs = Nothing
    Exit Function

Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
End Function

Private Function fnCadastrarUsuario(ByVal vNome As String, ByVal vUsuario As String, ByVal vSenha As String, ByVal vDepto As String, ByVal vEMail As String, ByRef vInativo As CheckBox, ByRef vAlterarSenha As CheckBox) As Boolean
On Error GoTo Erro

    fnCadastrarUsuario = False
    gintNovoUsuarioID = 0
    
'    strSQL = "INSERT INTO tb_Usuarios " & _
'             "(Nome,Usuario,Senha,Departamento,EMail,Inativo,AlterarSenha) " & _
'             "VALUES ('" & vNome & "','" & vUsuario & "','" & vSenha & "','" & vDepto & "','" & vEMail & "'," & vInativo & "," & vAlterarSenha & ")"
    
    strSQL = "SELECT * FROM tb_Usuarios WHERE Usuario = '" & vUsuario & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        rs!Nome = vNome & ""
        rs!Usuario = vUsuario & ""
        rs!Senha = vSenha & ""
        rs!Departamento = vDepto & ""
        rs!EMail = vEMail & ""
        rs!Inativo = vInativo
        rs!AlterarSenha = vAlterarSenha
        rs.Update
        gintNovoUsuarioID = rs!UsuarioID
        
        fnCadastrarUsuario = True
        Set rs = Nothing
        Exit Function
    End If
    
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Técnico"
End Function

Private Sub suLimparDados()
    txtNome.Text = ""
    txtUsuario.Text = ""
    txtDepto.Text = ""
    txtEMail.Text = ""
    txtSenha.Text = ""
    txtConfirmar.Text = ""
    chkAlterar.Value = 0
    chkInativo.Value = 0
    cboModulos.Clear
    Call suListarModulos
End Sub

Private Sub suLimparGridPermissoes()
    strSQL = "SELECT * FROM vw_Permissoes WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = True Then
        gexPermissoes.HoldFields
        Set gexPermissoes.ADORecordset = rs
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suHabilitarDados(ByVal vBoolean As Boolean)
    txtNome.Enabled = vBoolean
    txtUsuario.Enabled = vBoolean
    txtDepto.Enabled = vBoolean
    txtEMail.Enabled = vBoolean
    txtSenha.Enabled = vBoolean
    txtConfirmar.Enabled = vBoolean
    chkAlterar.Enabled = vBoolean
    chkInativo.Enabled = vBoolean
    cmdSalvar.Enabled = vBoolean
    cmdCancelar.Enabled = vBoolean
    chkEmBranco.Enabled = vBoolean
End Sub

Private Sub suMostrarPermissoes(ByVal vUsuarioID As Integer)
    strSQL = "SELECT * FROM vw_Permissoes WHERE UsuarioID = " & vUsuarioID & " ORDER BY Modulo "
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = False Then
        gexPermissoes.HoldFields
        Set gexPermissoes.ADORecordset = rs
    Else
        gexPermissoes.HoldFields
        Set gexPermissoes.ADORecordset = rs
    End If
    
    Set rs = Nothing
End Sub

Private Sub suMostrarDados(ByVal vUsuarioID As Integer)
    strSQL = "SELECT * FROM vw_Usuarios WHERE UsuarioID = " & vUsuarioID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        txtNome.Text = rs!Nome & ""
        txtUsuario.Text = rs!Usuario & ""
        txtDepto.Text = rs!Departamento & ""
        txtEMail.Text = rs!EMail & ""
        chkInativo.Value = IIf(rs!Inativo = True, 1, 0)
        chkAlterar.Value = IIf(rs!AlterarSenha = True, 1, 0)
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub suListarUsuarios()
    strSQL = "SELECT * FROM vw_Usuarios ORDER BY Nome"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        gexUsuarios.HoldFields
        Set gexUsuarios.ADORecordset = rs
    Else
        gexUsuarios.HoldFields
        Set gexUsuarios.ADORecordset = rs
    End If
    
    Set rs = Nothing
End Sub

