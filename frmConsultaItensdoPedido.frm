VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmConsultaItensdoPedido 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Compras"
   ClientHeight    =   5880
   ClientLeft      =   7485
   ClientTop       =   2460
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   3
      Top             =   5235
      Width           =   6165
   End
   Begin VB.TextBox txtDescricao2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensProduto 
      Height          =   4260
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Se desejar excluir um item, clique duas vezes sobre o item a ser excluido."
      Top             =   150
      Width           =   6165
      _cx             =   10874
      _cy             =   7514
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColor       =   14737632
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      BackColorAlternate=   12632256
      GridColor       =   14737632
      GridColorFixed  =   8421504
      TreeColor       =   8421504
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsultaItensdoPedido.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   2
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   330
      TabIndex        =   1
      Text            =   "0"
      Top             =   3495
      Visible         =   0   'False
      Width           =   270
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   525
      OleObjectBlob   =   "frmConsultaItensdoPedido.frx":00C3
      Top             =   1830
   End
   Begin VB.Label txtDescricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do item selecionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   150
      TabIndex        =   4
      Top             =   4515
      Width           =   3240
   End
End
Attribute VB_Name = "frmConsultaItensdoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim rsItensVenda As rdoResultset
Dim SQL As String
Dim wTotalitens As Integer
Private Sub cmdRetorna_Click()
  Unload Me
  frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub cmdRetorna_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Activate()
  txtpedido.Text = frmPedido.txtpedido.Text
  Call CarregaItens

End Sub

Private Sub Form_Load()
   Call AjustaTela(frmConsultaItensdoPedido)
End Sub
Private Sub CarregaItens()

  grdItensProduto.Rows = 1
  
  SQL = "Select NFItens.*, PR_Descricao From NFItens,ProdutoLoja " _
        & "Where Referencia = PR_Referencia and NumeroPed = " & txtpedido.Text _
        & " Order By Item"
  rsItensVenda.CursorLocation = adUseClient
  rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsItensVenda.EOF Then
     Do While Not rsItensVenda.EOF
        grdItensProduto.AddItem rsItensVenda("Item") & Chr(9) & rsItensVenda("Referencia") _
        & Chr(9) & rsItensVenda("PR_Descricao") & Chr(9) & rsItensVenda("Qtde") _
        & Chr(9) & Format(rsItensVenda("vlUnit"), "###,###,##0.00") & Chr(9) _
        & Format((rsItensVenda("vlUnit") * rsItensVenda("qtde")), "###,###,##0.00")
        rsItensVenda.MoveNext
     Loop
     grdItensProduto.Enabled = True
     grdItensProduto.Editable = flexEDNone
     grdItensProduto.Row = 1
     grdItensProduto.SetFocus
  End If
  rsItensVenda.Close
End Sub


Private Sub grdItensProduto_DblClick()

Dim i As Byte

On Error GoTo erronaUpdate

 If MsgBox("Deseja Excluir o Item = " & grdItensProduto.TextMatrix(grdItensProduto.Row, 0), _
           vbYesNo + vbQuestion, "Atenção") = vbYes Then
    adoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    
'*************************** ItensVenda TraderCaixa

         
    SQL = "Delete NFItens Where NumeroPed = " & txtpedido.Text & " and " _
          & "Item = " & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & " and " _
          & "Referencia = '" & Mid(grdItensProduto.TextMatrix(grdItensProduto.Row, 1), 1, 7) & "'"
         
         adoCNLoja.Execute (SQL)
         
         Screen.MousePointer = vbNormal
         adoCNLoja.CommitTrans
    
    Call SomaItensVenda
    Call CarregaItens
    
        For i = grdItensProduto.FixedRows To grdItensProduto.Rows - 1
            SQL = "update NFItens set item = " & i & " Where NumeroPed = " & txtpedido.Text & " and " _
            & "Referencia = '" & Mid(grdItensProduto.TextMatrix(i, 1), 1, 7) & "'"
            grdItensProduto.TextMatrix(i, 0) = i
            adoCNLoja.Execute (SQL)
        Next i
    
    If wTotalitens = 0 Then
       frmPedido.cmdTotalPedido.Caption = Format(0, "0.00") '+ "           "
       Unload Me
       frmPedido.txtPesquisar.SetFocus
       Exit Sub
    End If
      
    Exit Sub
 Else
    Exit Sub
 End If

erronaUpdate:
MsgBox "Erro na Exclusão Item " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal

End Sub


Private Sub SomaItensVenda()

SQL = ""
SQL = "Select Sum(VLUNIT * Qtde) as TotalVenda, Count(*) as TotalItens From NFItens " _
      & "Where Numeroped = " & txtpedido.Text & " and TipoNota = 'PD'"
 
  rsItensVenda.CursorLocation = adUseClient
  rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  frmPedido.cmdTotalPedido.Caption = Format(rsItensVenda("TotalVenda") + GBL_Frete, "###,###,##0.00")
  frmPedido.cmdQtdeItens.Caption = rsItensVenda("TotalItens")
  wTotalitens = rsItensVenda("Totalitens")
  
     'Do While Len(frmPedido.cmdTotalPedido.Caption) <= 12
         'frmPedido.cmdTotalPedido.Caption = frmPedido.cmdTotalPedido.Caption '+ " "
     'Loop
  
     'Do While Len(frmPedido.cmdQtdeItens.Caption) <= 5
         'frmPedido.cmdQtdeItens.Caption = frmPedido.cmdQtdeItens.Caption + " "
     'Loop

  rsItensVenda.Close
End Sub


Private Sub grdItensProduto_EnterCell()
     txtDescricao.Caption = " " & grdItensProduto.TextMatrix(grdItensProduto.Row, 2)
  
End Sub

Private Sub grdItensProduto_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
' frmPedido.picQuadroGeral.Width = 9975
  frmPedido.txtPesquisar.SetFocus
End If

End Sub

Private Sub lblDescricaoLBL_Click()

End Sub

Private Sub lblDescricao_Change()

End Sub

