VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTrocaModalidadeVenda 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Modalidade"
   ClientHeight    =   6180
   ClientLeft      =   5520
   ClientTop       =   2835
   ClientWidth     =   6570
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   1
      Top             =   4920
      Width           =   6165
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdPrecos 
      Height          =   3735
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      _cx             =   7223
      _cy             =   6588
      _ConvInfo       =   1
      Appearance      =   0
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
      FloodColor      =   3421236
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   13
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTrocaModalidadeVenda.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin Project1.chameleonButton cmbGravar 
      Height          =   405
      Left            =   5175
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Grava"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmTrocaModalidadeVenda.frx":00E5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdModalidade 
      Height          =   4035
      Left            =   120
      TabIndex        =   3
      Top             =   630
      Width           =   1995
      _cx             =   3519
      _cy             =   7117
      _ConvInfo       =   1
      Appearance      =   0
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
      FloodColor      =   3421236
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   13
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTrocaModalidadeVenda.frx":0101
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin MSMask.MaskEdBox mskDataInf 
      Height          =   255
      Left            =   3525
      TabIndex        =   5
      Tag             =   "DataMaiorCmp"
      Top             =   4395
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCondicaoFaturado 
      BackColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label lblPagamento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modalidade Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   6165
   End
End
Attribute VB_Name = "frmTrocaModalidadeVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String
Dim wValorVenda As Double
Dim wIndicePreco As String * 1
Dim wPrecoCalculado As Double
Dim wValorTotalCalculado As Double
Dim NomeParcela As String
Dim QtdeParcelas As Integer
Dim wTotalPedido As Double
Dim wGuardaLinha As Long
Dim wFormaPgto   As String
Dim wNroPedido As Double
Dim wCodigo As String
Dim wTipo As String
Dim wTipoCondicao As String
Dim wdata As String

'Dim wPagamento As String
'Dim wPagamentototal As String

Private Sub cmbGravar_Click()

    Dim financiado As Boolean
    Dim codigo As Integer

      If grdModalidade.TextMatrix(grdModalidade.Row, 0) Like "Finan*" Then
        codigo = wCodigo
        grdModalidade.Row = 1
        grdPrecos.Row = 1
        grdPrecos_Click
        financiado = True
        
      End If

      grdPrecos_Click
      wPagamento = grdPrecos.TextMatrix(grdPrecos.RowSel, 0)
      wPagamentototal = grdPrecos.TextMatrix(grdPrecos.RowSel, 2)
      
      Sql = "Exec SP_Atualiza_Modalidade_Venda_Pedido " & wNroPedido & ",'" & LTrim(RTrim(wTipo)) & "','" _
             & LTrim(RTrim(wCodigo)) & "','" & LTrim(RTrim(wTipoCondicao)) & "'"
      adoCNLoja.Execute Sql
      frmPedido.cmdTotalPedido.Caption = Format(wTotalPedido, "###,###,##0.00")
      wGravaModalidade = True

      If financiado Then
            Sql = "update nfcapa set CONDPAG = '3', ModalidadeVenda = 'FI', Parcelas=1 where numeroped = '" & wNroPedido & "'"
            adoCNLoja.Execute Sql
      End If

'If Not mskDataInf.Visible Then
If wdata <> "" Then
 Sql = "Update   nfcapa set DataPag='" & Format(wdata, "mm/dd/yyyy") & "' Where NumeroPed = " & wNroPedido & ""
      adoCNLoja.Execute Sql
End If
      
      Unload Me

'Else
        'MsgBox "Digitte uma Data"
'End If

End Sub


Private Sub cmbGravar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmbGravar_Click
    ElseIf KeyCode = 27 Then
        cmdRetorna_Click
    End If
End Sub

Private Sub cmdRetorna_Click()
 Unload Me
End Sub

Private Sub Form_Activate()

wPagamento = ""
wPagamentototal = ""
'Limpa tabelas

    If wLiberaBloqueioPreco Then
    
        Sql = "update nfitens set VLUNIT = PrecoUnitAlternativa,  VLTOTITEM = (PrecoUnitAlternativa * QTDE) " & _
              "from produtoloja " & _
              "where numeroped = " & frmPedido.txtpedido.Text & " and pr_referencia = referencia"
        adoCNLoja.Execute Sql
        
        Sql = "update nfcapa set condpag = 1 where numeroped = " & frmPedido.txtpedido.Text & ""
        adoCNLoja.Execute Sql
    
    Else
    
        Sql = "update nfitens set VLUNIT = PR_PrecoVenda1, VLTOTITEM = (PR_PrecoVenda1 * QTDE) " & _
              "from produtoloja " & _
              "where numeroped = " & frmPedido.txtpedido.Text & " and pr_referencia = referencia"
        adoCNLoja.Execute Sql
        
        Sql = "update nfcapa set condpag = 1 where numeroped = " & frmPedido.txtpedido.Text & ""
        adoCNLoja.Execute Sql
    
    End If
    
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

 wGravaModalidade = False
 wValorVenda = frmPedido.cmdTotalPedido
 wNroPedido = frmPedido.txtpedido.Text

 grdModalidade.Rows = 1
 grdModalidade.AddItem "A Vista"
 grdModalidade.AddItem "Cartão"
 grdModalidade.AddItem "Faturado"
 grdModalidade.AddItem "Financiado"
 grdModalidade.AddItem "Finan. / Cheque"
 
 grdModalidade.Row = 1
' grdModalidade_Click
 'Call MontaPrecos("AV")
 grdPrecos.Row = 1
 grdPrecos_Click
 grdModalidade.SetFocus
 'grdPrecos.Select 1, 1
 
End Sub

Private Sub Form_Load()

    Call AjustaTela(Me)
    
End Sub

Private Sub MontaPrecos(CodigoCrediario As String)

     grdPrecos.Rows = 1
     
     Sql = "select pr_indicePreco from produtoloja, nfitens where numeroped = '" & wNroPedido & "' and pr_referencia = referencia and pr_indicePreco = 8"
     rsCondicaoFaturadoCentavos.CursorLocation = adUseClient
     rsCondicaoFaturadoCentavos.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     If wLiberaBloqueioPreco Then
     
         If Not rsCondicaoFaturadoCentavos.EOF Then
     
            Sql = "Select cp_tipo,CP_Codigo,CP_Condicao,CP_TipoCondicao,cp_parcelas,Floor(SUM(PrecoUnitAlternativa * qtde), 1)," _
                & " Floor (sum(((PrecoUnitAlternativa * qtde) * cp_coeficiente)), 1) as PrecoID," _
                & " Floor (sum((((PrecoUnitAlternativa * qtde) * cp_coeficiente)/cp_parcelas)),1) as ValorParcela " _
                & " From produtoloja, CondicaoPagamento, nfitens " _
                & " where PR_IndicePreco=CP_ID and CP_Tipo='" & CodigoCrediario _
                & "' and PR_Referencia = REFERENCIA and NUMEROPED =" & wNroPedido _
                & " group by cp_tipo,CP_Codigo,CP_TipoCondicao,cp_parcelas,CP_Condicao"
          
         Else
         
            Sql = "Select cp_tipo,CP_Codigo,CP_Condicao,CP_TipoCondicao,cp_parcelas,round(SUM(PrecoUnitAlternativa * qtde), 1)," _
                & " round (sum(((PrecoUnitAlternativa * qtde) * cp_coeficiente)), 1) as PrecoID," _
                & " round (sum((((PrecoUnitAlternativa * qtde) * cp_coeficiente)/cp_parcelas)),1) as ValorParcela " _
                & " From produtoloja, CondicaoPagamento, nfitens " _
                & " where PR_IndicePreco=CP_ID and CP_Tipo='" & CodigoCrediario _
                & "' and PR_Referencia = REFERENCIA and NUMEROPED =" & wNroPedido _
                & " group by cp_tipo,CP_Codigo,CP_TipoCondicao,cp_parcelas,CP_Condicao"
         
         End If
          
         rsCondicaoFaturado.CursorLocation = adUseClient
         rsCondicaoFaturado.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     
     Else
          
             If Not rsCondicaoFaturadoCentavos.EOF Then
                
                 Sql = "Select cp_tipo,CP_Codigo,CP_Condicao,CP_TipoCondicao,cp_parcelas," & vbNewLine _
                 & " Floor (SUM(pr_precovenda1 * qtde))," & vbNewLine _
                 & " Floor (sum(((pr_precovenda1 * qtde) * cp_coeficiente))) as PrecoID," & vbNewLine _
                 & " Floor (sum((((pr_precovenda1 * qtde) * cp_coeficiente)/cp_parcelas))) as ValorParcela " & vbNewLine _
                 & " From produtoloja, CondicaoPagamento, nfitens " & vbNewLine _
                 & " where PR_IndicePreco=CP_ID and CP_Tipo='" & CodigoCrediario _
                 & "' and PR_Referencia = REFERENCIA and NUMEROPED =" & wNroPedido _
                 & " group by cp_tipo,CP_Codigo,CP_TipoCondicao,cp_parcelas,CP_Condicao"
            
             Else
                 Sql = "Select cp_tipo,CP_Codigo,CP_Condicao,CP_TipoCondicao,cp_parcelas," & vbNewLine _
                     & " round (SUM(pr_precovenda1 * qtde), 1)," & vbNewLine _
                     & " round (sum(((pr_precovenda1 * qtde) * cp_coeficiente)), 1) as PrecoID," & vbNewLine _
                     & " round (sum((((pr_precovenda1 * qtde) * cp_coeficiente)/cp_parcelas)),1) as ValorParcela " & vbNewLine _
                     & " From produtoloja, CondicaoPagamento, nfitens " & vbNewLine _
                     & " where PR_IndicePreco=CP_ID and CP_Tipo='" & CodigoCrediario _
                     & "' and PR_Referencia = REFERENCIA and NUMEROPED =" & wNroPedido _
                     & " group by cp_tipo,CP_Codigo,CP_TipoCondicao,cp_parcelas,CP_Condicao"
                
                    
             End If
            rsCondicaoFaturado.CursorLocation = adUseClient
            rsCondicaoFaturado.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
          
        
     
     End If
     
     rsCondicaoFaturadoCentavos.Close
     
     
    
     Do While Not rsCondicaoFaturado.EOF
        If rsCondicaoFaturado("CP_Tipo") = "FA" Then
           grdPrecos.AddItem Format(rsCondicaoFaturado("PrecoID"), "###,###,##0.00") _
                  & Chr(9) & rsCondicaoFaturado("CP_TipoCondicao") _
                  & Chr(9) & rsCondicaoFaturado("CP_Condicao") _
                  & Chr(9) & rsCondicaoFaturado("CP_Tipo") _
                  & Chr(9) & rsCondicaoFaturado("CP_Codigo")
        Else
           grdPrecos.AddItem Format(rsCondicaoFaturado("CP_Parcelas"), "00") _
                  & Chr(9) & Format(rsCondicaoFaturado("PrecoID"), "###,###,##0.00") _
                  & Chr(9) & Format(rsCondicaoFaturado("ValorParcela"), "###,###,##0.00") _
                  & Chr(9) & rsCondicaoFaturado("CP_TipoCondicao") _
                  & Chr(9) & rsCondicaoFaturado("CP_Tipo") _
                  & Chr(9) & rsCondicaoFaturado("CP_Condicao") _
                  & Chr(9) & rsCondicaoFaturado("CP_Codigo")
       End If
                 
        rsCondicaoFaturado.MoveNext
     Loop
    
     rsCondicaoFaturado.Close

End Sub


Private Sub grdModalidade_EnterCell()
    grdPrecos.TextMatrix(0, 0) = "Parcelas"
    grdPrecos.ColWidth(0) = 1005
    grdPrecos.TextMatrix(0, 1) = "Preço"
    grdPrecos.ColWidth(1) = 1320
    grdPrecos.TextMatrix(0, 2) = "Valor Parcelas"
    grdPrecos.ColWidth(2) = 1425
    txtCondicaoFaturado.Text = ""
    If grdModalidade.TextMatrix(grdModalidade.Row, 0) = "A Vista" Then
       Call MontaPrecos("AV")
    ElseIf grdModalidade.TextMatrix(grdModalidade.Row, 0) = "Cartão" Then
       Call MontaPrecos("CC")
    ElseIf grdModalidade.TextMatrix(grdModalidade.Row, 0) = "Faturado" Then
       grdPrecos.TextMatrix(0, 0) = "Preço"
       grdPrecos.ColWidth(0) = 1065
       grdPrecos.TextMatrix(0, 1) = "Cod."
       grdPrecos.ColWidth(1) = 510
       grdPrecos.TextMatrix(0, 2) = "Descrição"
       grdPrecos.ColWidth(2) = 2355
       Call MontaPrecos("FA")
    ElseIf grdModalidade.TextMatrix(grdModalidade.Row, 0) = "Financiado" Then
       Call MontaPrecos("FI")
    ElseIf grdModalidade.TextMatrix(grdModalidade.Row, 0) = "Finan. / Cheque" Then
       Call MontaPrecos("CH")
    End If
End Sub

Private Sub grdModalidade_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Or KeyCode = vbKeyF2 Then
        cmbGravar_Click
    ElseIf KeyCode = 27 Then
        cmdRetorna_Click
        ElseIf KeyCode = 13 Then
        grdPrecos.SetFocus
        grdPrecos.Row = 1
    End If
End Sub

Private Sub grdPrecos_Click()
    If grdPrecos.Row > 0 Then
        If grdPrecos.TextMatrix(grdPrecos.Row, 3) <> "FA" Then
            wCodigo = grdPrecos.TextMatrix(grdPrecos.Row, 6)
            wTipo = grdPrecos.TextMatrix(grdPrecos.Row, 4)
            wTotalPedido = grdPrecos.TextMatrix(grdPrecos.Row, 1)
            wTipoCondicao = grdPrecos.TextMatrix(grdPrecos.Row, 3)
        Else
            wCodigo = grdPrecos.TextMatrix(grdPrecos.Row, 1)
            wTipo = grdPrecos.TextMatrix(grdPrecos.Row, 3)
            wTotalPedido = grdPrecos.TextMatrix(grdPrecos.Row, 0)
            wTipoCondicao = grdPrecos.TextMatrix(grdPrecos.Row, 1)
        End If
    End If
    If grdPrecos.TextMatrix(grdPrecos.Row, 1) <> "85" Then
mskDataInf.Visible = False
txtCondicaoFaturado.Text = ""
txtCondicaoFaturado.BackColor = &HC0C0C0
End If

If grdPrecos.TextMatrix(grdPrecos.Row, 1) = "85" And mskDataInf.Text = "__/__/____" Then
    mskDataInf.Visible = True
    mskDataInf.Text = "__/__/____"
    mskDataInf.SetFocus
    txtCondicaoFaturado.Text = "Informe á Data"
    txtCondicaoFaturado.BackColor = &HC0FFFF
End If

End Sub


Private Sub grdPrecos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Or KeyCode = vbKeyF2 Then
        cmbGravar_Click
    ElseIf KeyCode = 27 Then
        cmdRetorna_Click
        ElseIf KeyCode = 13 Then
        cmbGravar.SetFocus
        
    End If
End Sub


Private Sub mskDataInf_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If IsDate(mskDataInf.Text) = True Then
        wdata = mskDataInf.Text
        mskDataInf.Visible = True
        'txtCondicaoFaturado.Text = ""
        'txtCondicaoFaturado.BackColor = &HC0C0C0
        Else
            MsgBox "Data invalida"
            mskDataInf.SetFocus
        End If
    End If
End Sub

Private Sub mskDataInf_LostFocus()
    If IsDate(mskDataInf.Text) = True Then
        wdata = mskDataInf.Text
    End If
End Sub
