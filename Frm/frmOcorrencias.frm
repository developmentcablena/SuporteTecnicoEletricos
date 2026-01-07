VERSION 5.00
<<<<<<< HEAD
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
=======
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
Begin VB.Form frmOcorrencias 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Técnico - Ocorrências"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOcorrencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
<<<<<<< HEAD
   Begin MSComctlLib.ListView lvwOcorrencia 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9551
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
=======
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
<<<<<<< HEAD
=======
   Begin GridEX20.GridEX gexOcorrencias 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8916
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmOcorrencias.frx":0CCA
      Column(2)       =   "frmOcorrencias.frx":0E76
      Column(3)       =   "frmOcorrencias.frx":101E
      Column(4)       =   "frmOcorrencias.frx":11BE
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmOcorrencias.frx":1352
      FormatStyle(2)  =   "frmOcorrencias.frx":147E
      FormatStyle(3)  =   "frmOcorrencias.frx":152E
      FormatStyle(4)  =   "frmOcorrencias.frx":15E2
      FormatStyle(5)  =   "frmOcorrencias.frx":16BA
      ImageCount      =   0
      PrinterProperties=   "frmOcorrencias.frx":1772
   End
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
End
Attribute VB_Name = "frmOcorrencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub
<<<<<<< HEAD


Private Sub Form_Load()
With Me.lvwOcorrencia
    .ColumnHeaders.Add , , "Reporte Técnico", Width:=2500
    .ColumnHeaders.Add , , "Data Reporte", Width:=2500
    .ColumnHeaders.Add , , "Ocorrência", Width:=4500
    .ColumnHeaders.Add , , "Data Ocorrência", Width:=2500
End With
Call suMostrarOcorrencia(gintOSID)
End Sub


Private Sub suMostrarOcorrencia(ByVal vOSID As Long)
Dim strSQL As String
Dim lvwItem As ListItem
Dim rs As ADODB.Recordset
Dim frm As New frmChamados
Dim selectedItem As ListItem
Call ConectarBD




    strSQL = "SELECT * FROM tb_Ocorrencias WHERE OSID = " & vOSID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly

    Me.lvwOcorrencia.ListItems.Clear

    If rs.EOF = False Then
        Do While Not rs.EOF
            Set lvwItem = Me.lvwOcorrencia.ListItems.Add(, , rs!ReporteTecnico)
            lvwItem.SubItems(1) = rs!DataReporte
            lvwItem.SubItems(2) = rs!Ocorrencia
            lvwItem.SubItems(3) = rs!DataOcorrencia
            rs.MoveNext
        Loop
    End If

    rs.Clone
    Set rs = Nothing
    Set cn = Nothing
End Sub

=======
>>>>>>> 8c6a2da482b88bea820591297e72d3467bc38812
