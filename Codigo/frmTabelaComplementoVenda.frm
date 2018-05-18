VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmDadosGerais 
   Caption         =   "Dados Gerais"
   ClientHeight    =   3900
   ClientLeft      =   10200
   ClientTop       =   3315
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   4830
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   195
      OleObjectBlob   =   "frmTabelaComplementoVenda.frx":0000
      Top             =   3540
   End
   Begin VB.CommandButton cmdRetorna 
      Caption         =   "Retorna"
      Height          =   390
      Left            =   3255
      TabIndex        =   4
      Top             =   3450
      Width           =   1455
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   3450
      Width           =   1455
   End
   Begin VB.Frame fraDados 
      Height          =   825
      Left            =   105
      TabIndex        =   5
      Top             =   -15
      Width           =   4635
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   195
         Left            =   1530
         OleObjectBlob   =   "frmTabelaComplementoVenda.frx":0234
         TabIndex        =   9
         Top             =   135
         Width           =   1530
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   195
         Left            =   810
         OleObjectBlob   =   "frmTabelaComplementoVenda.frx":02A4
         TabIndex        =   8
         Top             =   135
         Width           =   525
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Left            =   120
         OleObjectBlob   =   "frmTabelaComplementoVenda.frx":030A
         TabIndex        =   7
         Top             =   135
         Width           =   510
      End
      Begin VB.TextBox txtSequencia 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   765
         TabIndex        =   1
         Top             =   345
         Width           =   645
      End
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1485
         MaxLength       =   20
         TabIndex        =   2
         Top             =   345
         Width           =   3030
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   105
         TabIndex        =   0
         Top             =   345
         Width           =   585
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItens 
      Height          =   2550
      Left            =   105
      TabIndex        =   6
      Top             =   855
      Width           =   4605
      _cx             =   8123
      _cy             =   4498
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16761024
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12582912
      ForeColorSel    =   65535
      BackColorBkg    =   16761024
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTabelaComplementoVenda.frx":0374
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmDadosGerais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim rsDadosGerais As rdoResultset
Dim SQL As String

Private Sub cmdGravar_Click()
  Call GravaDadosGerais
  Call CarregaGride
  txtCodigo.Text = ""
  txtSequencia.Text = ""
  txtDescricao.Text = ""
  txtCodigo.SetFocus
End Sub

Private Sub cmdRetorna_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2

 ' Skin1.LoadSkin "C:\WINDOWS\system\skin.skn"
 ' Skin1.ApplySkin Me.hwnd

 
 Call CarregaGride
End Sub
Private Sub GravaDadosGerais()

On Error GoTo erronaInclusao
  adoCNLoja.BeginTrans
  Screen.MousePointer = vbHourglass
  SQL = "Insert Into DadosGerais (DAG_Codigo,DAG_Sequencia,DAG_NomeCampo)" _
      & "Values (" & txtCodigo.Text & "," & txtSequencia.Text & ",'" & txtDescricao.Text & "')"
  adoCNLoja.Execute (SQL)
  Screen.MousePointer = vbNormal
  adoCNLoja.CommitTrans
  Exit Sub
  
erronaInclusao:
MsgBox "Erro na Inclusão de Dados Gerais " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
  
End Sub
Private Sub CarregaGride()

  grdItens.Rows = 1
  SQL = "Select * from DadosGerais"
 
  rsDadosGerais.CursorLocation = adUseClient
  rsDadosGerais.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  If Not rsDadosGerais.EOF Then
     Do While Not rsDadosGerais.EOF
        grdItens.AddItem rsDadosGerais("DAG_Codigo") & Chr(9) _
        & rsDadosGerais("DAG_Sequencia") & Chr(9) & rsDadosGerais("DAG_NomeCampo")
        rsDadosGerais.MoveNext
     Loop
  End If
  rsDadosGerais.Close
End Sub

Private Sub grdItens_DblClick()

 If MsgBox("Deseja Excluir o Item = " & grdItens.TextMatrix(grdItens.Row, 2), vbYesNo + vbQuestion, "Atenção") = vbYes Then
    On Error GoTo ErronaDelecao
    adoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    SQL = "Delete DadosGerais Where DAG_Codigo = " & grdItens.TextMatrix(grdItens.Row, 0) _
        & " and DAG_Sequencia = " & grdItens.TextMatrix(grdItens.Row, 1)
    adoCNLoja.Execute (SQL)
    Screen.MousePointer = vbNormal
    adoCNLoja.CommitTrans
    Call CarregaGride
    Exit Sub
 End If
 
ErronaDelecao:
MsgBox "Erro na Deleção de Dados Gerais " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
End Sub
