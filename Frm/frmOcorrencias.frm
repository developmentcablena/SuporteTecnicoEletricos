VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
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
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
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
