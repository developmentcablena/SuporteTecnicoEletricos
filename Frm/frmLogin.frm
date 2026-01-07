VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Validar Usuário"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
<<<<<<< HEAD
      BackColor       =   &H80000018&
=======
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
<<<<<<< HEAD
      BackColor       =   &H80000018&
=======
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
      Height          =   285
      Left            =   120
      MaxLength       =   15
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblSenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
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
      TabIndex        =   2
      Top             =   840
      Width           =   600
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
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
   Begin VB.Image imgSeguranca 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2280
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private blnAlterarSenha As Boolean

Private Sub cmdCancelar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdValidar_Click()
    If Len(Trim(txtUsuario.Text)) = 0 Then
        MsgBox "Digite o usuário!", vbOKOnly + vbExclamation, "Suporte Técnico"
        txtUsuario.SetFocus
        Exit Sub
    End If
    
'    If Len(Trim(txtSenha.Text)) = 0 Then
'        MsgBox "Digite a senha!", vbOKOnly + vbExclamation, "Suporte Técnico"
'        txtSenha.SetFocus
'        Exit Sub
'    End If

    If fnValidarUsuario(Trim(txtUsuario.Text), CodDec(Trim(txtSenha.Text))) = True Then
        Call Unload(Me)
        Call frmPrincipal.Show
    ElseIf blnAlterarSenha = True Then
        Call Unload(Me)
        Call frmAlterarSenha.Show
    Else
        MsgBox "Usuário ou senha inválida!", vbOKOnly + vbCritical, "Suporte Técnico"
        txtUsuario.Text = ""
        txtSenha.Text = ""
        txtUsuario.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call ConectarBD
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cn.Close
    Set cn = Nothing
End Sub

Private Function fnValidarUsuario(ByVal vUsuario As String, ByVal vSenha As String) As Boolean
    fnValidarUsuario = False
    gintUsuarioID = 0
    gstrNome = ""
    gstrDepto = ""
    gstrEMail = ""
    gstrSenha = ""
    blnAlterarSenha = False
    
    strSQL = "SELECT * FROM vw_Usuarios WHERE Usuario = '" & vUsuario & "' AND Senha = '" & vSenha & "' AND Inativo = 0"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
<<<<<<< HEAD
        gintUsuarioID = rs!usuarioID
=======
        gintUsuarioID = rs!UsuarioID
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
        gstrNome = rs!Nome
        gstrDepto = rs!Departamento
        gstrEMail = rs!EMail
        gstrSenha = rs!Senha
        
        If rs!AlterarSenha = True Then
            blnAlterarSenha = True
        Else
            fnValidarUsuario = True
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdValidar_Click
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtUsuario.Text)) > 0 Then
            txtSenha.SetFocus
        End If
    End If
End Sub
