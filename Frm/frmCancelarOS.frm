VERSION 5.00
Begin VB.Form frmCancelarOS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Cancelar OS"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCancelarOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   345
      Left            =   960
      MaxLength       =   5
      TabIndex        =   4
      Top             =   150
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   4215
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
      TabIndex        =   5
      Top             =   270
      Width           =   705
   End
   Begin VB.Label lblMotivo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Motivo:"
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
      TabIndex        =   1
      Top             =   720
      Width           =   705
   End
End
Attribute VB_Name = "frmCancelarOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String


Private Sub cmdCancelar_Click()
    If Len(txtMotivo.Text) = 0 Then
        MsgBox "Digite o motivo pelo qual está cancelando a OS " & Format(gintOSID, "0000") & "!", vbOKOnly + vbExclamation, "Suporte Técnico"
        Exit Sub
    End If
    
    Call suCancelarOS(Trim(txtOSID.Text), Trim(txtMotivo.Text))
    Call Unload(Me)
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    With Me
        .txtOSID.Text = Format(gintOSID, "0000")
    End With
End Sub

Private Sub suCancelarOS(ByVal vOSID As Integer, ByVal vMotivo As String)
Call ConectarBD
    strSQL = "UPDATE tb_OS SET DataCancelamento = '" & Format(Now, "yyyy-MM-dd HH:mm") & "',Comentario='" & vMotivo & "', Status = 4 " & _
             "WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
End Sub
