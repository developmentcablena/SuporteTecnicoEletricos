VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrincipal 
   Appearance      =   0  'Flat
   Caption         =   "Suporte Técnico - ELÉTRICOS"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8220
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3969
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "05/01/2026"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "14:54"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   10320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1350
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":3C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":4556
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":4AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":4F42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   3240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   5715
      ButtonWidth     =   3122
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "imlMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Suporte Sistemas"
            Key             =   "Sistemas"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acompanhar OS"
            Key             =   "Acompanhar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Suporte Manutenção"
            Key             =   "Manutencao"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Verificar Chamados"
            Key             =   "Chamados"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar Relatórios"
            Key             =   "Relatorios"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Usuários"
            Key             =   "Usuarios"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cadastrar Itens"
            Key             =   "Cadastrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inventário"
            Key             =   "Inventario"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCablena 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   8880
      Picture         =   "frmPrincipal.frx":5C1C
      Top             =   6360
      Width           =   1875
   End
   Begin VB.Shape shpPrincipal 
      BackStyle       =   1  'Opaque
      Height          =   8175
      Left            =   0
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private intTotalOS As Integer

Private Sub Form_Load()

    With Me
        .tlbMenu.Buttons(1).Visible = False
        .tlbMenu.Buttons(2).Visible = False
        .tlbMenu.Buttons(3).Visible = False
        .tlbMenu.Buttons(4).Visible = False
        .tlbMenu.Buttons(5).Visible = False
        .tlbMenu.Buttons(6).Visible = False
        .tlbMenu.Buttons(7).Visible = False
        .tlbMenu.Buttons(8).Visible = False
    End With

    Call ConectarBD
    Call suVerificarPermissoes(gintUsuarioID)
    
    If fnVerificarPendenciasAceite(gintUsuarioID) > 0 Then
        MsgBox "Você tem [" & intTotalOS & "] OS sem registrar o ACEITE!" & vbCrLf & "FAVOR REGISTRAR O ACEITE NO SISTEMA!", vbOKOnly + vbCritical, "Suporte Técnico"
    End If
End Sub

Private Sub Form_Resize()
    shpPrincipal.Width = Me.Width
    shpPrincipal.Height = Me.Height
    stbPrincipal.Panels(1).Text = "Usuário: " & gstrNome
    stbPrincipal.Panels(2).Text = "Departamento: " & gstrDepto
    stbPrincipal.Panels(5).Text = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    imgCablena.Left = Me.Width - 2400
    imgCablena.Top = Me.Height - 2700
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cn.Close
    Set cn = Nothing
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case Is = "Sistemas"
            Call frmSuporteSistemas.Show(vbModal)
        Case Is = "Chamados"
            Call frmChamados.Show(vbModal)
        Case Is = "Usuarios"
            Call frmUsuarios.Show(vbModal)
        Case Is = "Acompanhar"
            Call frmAcompanharOS.Show(vbModal)
        Case Is = "Cadastrar"
            Call frmCadastrarItens.Show(vbModal)
        Case Is = "Relatorios"
            Call frmRelatorios.Show(vbModal)
        Case Is = "Inventario"
            Call frmInventario.Show(vbModal)
        Case Is = "Sair"
            Call Unload(Me)
    End Select
End Sub

Private Function fnVerificarPendenciasAceite(ByVal vUsuarioID As Integer) As Integer
    intTotalOS = 0
    
    strSQL = "SELECT * FROM vw_Chamados WHERE UsuarioID = " & vUsuarioID & " AND Status = 2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        intTotalOS = rs.RecordCount
        fnVerificarPendenciasAceite = intTotalOS
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub suVerificarPermissoes(ByVal vUsuarioID As Integer)
    strSQL = "SELECT * FROM vw_Permissoes WHERE UsuarioID = " & vUsuarioID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
    
        Select Case rs!Modulo
            Case Is = "Suporte Sistemas"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(1).Visible = True
                Else
                    tlbMenu.Buttons(1).Visible = False
                End If
            Case Is = "Acompanhar OS"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(2).Visible = True
                Else
                    tlbMenu.Buttons(2).Visible = False
                End If
            Case Is = "Suporte Manutenção"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(3).Visible = True
                Else
                    tlbMenu.Buttons(3).Visible = False
                End If
            Case Is = "Verificar Chamados"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(4).Visible = True
                Else
                    tlbMenu.Buttons(4).Visible = False
                End If
            Case Is = "Gerar Relatórios"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(5).Visible = True
                Else
                    tlbMenu.Buttons(5).Visible = False
                End If
            Case Is = "Usuários"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(6).Visible = True
                Else
                    tlbMenu.Buttons(6).Visible = False
                End If
            Case Is = "Cadastrar Itens"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(7).Visible = True
                Else
                    tlbMenu.Buttons(7).Visible = False
                End If
            Case Is = "Cadastrar Inventario"
                If rs!Permissao = True Then
                    tlbMenu.Buttons(8).Visible = True
                Else
                    tlbMenu.Buttons(8).Visible = False
                End If
        End Select
        
        rs.MoveNext
        
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub
