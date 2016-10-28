VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmGarantiaEstendida 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Garantia Estendida"
   ClientHeight    =   5580
   ClientLeft      =   2070
   ClientTop       =   4020
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   1
      Top             =   4875
      Width           =   6165
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdProdGarantia 
      Height          =   3270
      Left            =   150
      TabIndex        =   2
      Top             =   630
      Width           =   6165
      _cx             =   10874
      _cy             =   5768
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
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
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGarantiaEstendida.frx":0000
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
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5280
      TabIndex        =   8
      Top             =   5055
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmGarantiaEstendida.frx":00EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblValorUnitarioItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
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
      Left            =   1875
      TabIndex        =   7
      Top             =   4380
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Unitário:"
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
      Left            =   165
      TabIndex        =   6
      Top             =   4380
      Width           =   1545
   End
   Begin VB.Label lblDescricao 
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
      Left            =   165
      TabIndex        =   5
      Top             =   4005
      Width           =   3240
   End
   Begin VB.Label lblValorTotalGarantia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5250
      TabIndex        =   4
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label lblTotalVendas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total com Garantia:"
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
      Left            =   3135
      TabIndex        =   3
      Top             =   4380
      Width           =   2055
   End
   Begin VB.Label lblPagamento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Garantia Estendida"
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
      TabIndex        =   0
      Top             =   120
      Width           =   6165
   End
End
Attribute VB_Name = "frmGarantiaEstendida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wNumeroPedido As String
Dim I As Integer
Dim wColunaQuantidade As Integer
Dim wColunaGarantia24 As Integer
Dim wColunaGarantia36 As Integer
Dim wcorCelulaDesativada As ColorConstants
Dim contLinhasMarcada As Integer
Dim quantidadeOriginal As Integer
Public valorGarantiaEstendida As Double

Private Sub cmdGrava_Click()
    Screen.MousePointer = 11
    
    If atualizaItens Then
        'frmFechaPedido.wTotalGarantia = (lblValorTotalGarantia.Caption - frmPedido.cmdTotalPedido.Caption)
        'frmFechaPedido.lblTotalGarantia.Visible = True
        'frmFechaPedido.lblTotalGarantia.Caption = "GE: " & Format((txtTotal.Text - frmPedido.lblTotalPedido), "0.00")
        
        'frmPedido.cmdTotalPedidoGE.Caption = "+ G.E " & frmPedido.cmdTotalPedido.Caption + _
        valorGarantiaEstendida
        frmPedido.cmdTotalPedidoGE.Caption = "+ G.E " & valorGarantiaEstendida
        frmPedido.cmdTotalPedidoGE.Visible = True
        
        atualizaVendedor
        Unload Me
    Else
        MsgBox "Não há itens com garantia marcada" & vbNewLine, vbInformation, "Garantia Estendida"
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdRetorna_Click()
     Unload Me
End Sub

Private Sub limpaCamposGarantiaTabela(NumeroPedido)
    Dim SQL As String
    
    SQL = "update nfitens" & vbNewLine & _
    "set qtdeGarantia = default," & vbNewLine & _
    "coeficientePlano = default, garantiaEstendida = default, planoGarantia = default," & vbNewLine & _
    "ValorGarantia = default, ge_premioLiquido = default," & vbNewLine & _
    "ge_iof = default," & vbNewLine & _
    "ge_dataInicioVigencia = default," & vbNewLine & _
    "ge_dataFinalVigencia = default," & vbNewLine & _
    "ge_valorCustoSeguradora = default" & vbNewLine & _
    "Where numeroPed = " & NumeroPedido

    SQL = SQL & vbNewLine & vbNewLine
    SQL = SQL & "Update nfcapa" & vbNewLine & _
    "set garantiaEstendida = default, " & vbNewLine & _
    "totalGarantia = default, " & vbNewLine & _
    "vendedorGarantia = default" & vbNewLine & _
    "Where numeroPed = " & NumeroPedido

    adoCNLoja.Execute SQL

End Sub

Private Sub Form_Activate()
    
    valorGarantiaEstendida = 0
    wNumeroPedido = frmPedido.txtpedido.Text
    lblValorTotalGarantia.Caption = RTrim(frmPedido.cmdTotalPedido.Caption)
    montaCamposGrid
    
'    TimerIni.Enabled = True
    limpaCamposGarantiaTabela wNumeroPedido
    CarregaGrid wNumeroPedido
    grdProdGarantia.SetFocus
    grdProdGarantia.Row = 1
    
    End Sub

Private Sub montaCamposGrid()
    For I = 0 To grdProdGarantia.FixedRows - 1
        grdProdGarantia.MergeRow(I) = True
    Next I
   
    For I = 0 To grdProdGarantia.Cols - 1
        grdProdGarantia.MergeCol(I) = True
    Next I
    
    wColunaQuantidade = 1
    wColunaGarantia24 = 3
    wColunaGarantia36 = 5
    wcorCelulaDesativada = &HE0E0E0
    grdProdGarantia.Rows = grdProdGarantia.FixedRows
    'frmGarantiaEstendidaProdutos.Left = frmFechaPedido.Left
End Sub

Private Sub Form_Load()
    Call AjustaTela(frmGarantiaEstendida)
End Sub

Private Sub grdProdGarantia_Click()
    lblDescricao.Caption = grdProdGarantia.TextMatrix(grdProdGarantia.Row, 8)
End Sub

Private Sub grdProdGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
        cmdGrava_Click
    ElseIf KeyCode = 27 Then
           Unload Me
    ElseIf KeyCode = 13 Then
        cmdGrava.SetFocus
           
    End If
End Sub

Private Sub grdProdGarantia_LostFocus()
    lblDescricao.Caption = "Descrição do item selecionado"
End Sub



Private Sub CarregaGrid(numeroPed As String)
    Dim precoTotal As Double
    
    Dim SQL As String
    Dim valorGarantia24 As String
    Dim valorGarantia36 As String
    Dim rsProdutoGarantia As New ADODB.Recordset
    
    SQL = "select Referencia, Qtde, VlUnit, cast(pr_garantiaFabricante/30 as integer) as garantiaFabricante, " & vbNewLine & _
          "pr_descricao as descricaoProduto " & vbNewLine & _
          "from produtoLoja, nfitens " & vbNewLine & _
          "where numeroPed = " & numeroPed & " and pr_referencia = referencia" & vbNewLine & _
          "and pr_garantiaEstendida = 'S' " & vbNewLine & _
          "and cast(pr_garantiaFabricante/30 as integer) < 36 " & vbNewLine & _
          "order by item"
    
        rsProdutoGarantia.CursorLocation = adUseClient
        rsProdutoGarantia.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        Do While Not rsProdutoGarantia.EOF
        
            grdProdGarantia.AddItem rsProdutoGarantia("referencia") & Chr(9) & _
                                    rsProdutoGarantia("qtde") & Chr(9) & _
                                    rsProdutoGarantia("garantiaFabricante") & Chr(9) & _
                                    False & Chr(9) & " - - " & Chr(9) & _
                                    False & Chr(9) & " - - " & Chr(9) & _
                                    Format(rsProdutoGarantia("VLUNIT"), "##,#0.00") & Chr(9) & _
                                    rsProdutoGarantia("descricaoProduto")
                                                            
            If rsProdutoGarantia("garantiaFabricante") >= 24 Then
                desativaOpcaoGarantia wColunaGarantia24, grdProdGarantia.Rows - 1
            Else
                grdProdGarantia.TextMatrix(grdProdGarantia.Rows - 1, wColunaGarantia24 + 1) = _
                formatValorParaExibir(calculoCoeficientePedido(rsProdutoGarantia("VLUNIT"), _
                Val(rsProdutoGarantia("qtde")), 24, rsProdutoGarantia("VLUNIT")))
            End If
            grdProdGarantia.TextMatrix(grdProdGarantia.Rows - 1, wColunaGarantia36 + 1) = _
            formatValorParaExibir(calculoCoeficientePedido(rsProdutoGarantia("VLUNIT"), _
            Val(rsProdutoGarantia("qtde")), 36, rsProdutoGarantia("VLUNIT")))
            
            rsProdutoGarantia.MoveNext
        Loop
    rsProdutoGarantia.Close
End Sub

Private Sub desativaOpcaoGarantia(ByVal coluna As Integer, ByVal Linha As Integer)
    For coluna = coluna To coluna + 1
        grdProdGarantia.Col = coluna
            grdProdGarantia.Row = Linha
            grdProdGarantia.CellBackColor = wcorCelulaDesativada
    Next coluna
End Sub

Private Function calculoCoeficientePedido(precoTotal As Double, quantidade As Integer, meses As Integer, precoUnitario As String) As Double
    calculoCoeficientePedido = (precoTotal + obterCoeficiente(meses, CDbl(precoUnitario))) * quantidade
End Function

Private Function obterCoeficiente(garantia As Integer, valorUnitario As Double) As Double
    Dim rsProdutoGarantia As New ADODB.Recordset
    Dim SQL As String
    
On Error GoTo trataerro
    
    SQL = "select fpg_premio " & vbNewLine & _
          "from FIN_faixapremioge " & vbNewLine & _
          "where '" & Replace(CStr(valorUnitario), ",", ".") & "' between fpg_faixainicial " & vbNewLine & _
          "and fpg_faixaFinal and fpg_plano = " & garantia
    
    'Set rsProdutoGarantia = rdoCnLoja.OpenResultset(sql)
        rsProdutoGarantia.CursorLocation = adUseClient
        rsProdutoGarantia.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        obterCoeficiente = rsProdutoGarantia("fpg_premio")
    rsProdutoGarantia.Close
Exit Function
trataerro:
    Select Case Err.Number
        Case Else
            MsgBox "Não foi possível obter o Coeficiente", vbCritical, "Erro interno"
            'End
    End Select
    
End Function

Private Function formatValorParaExibir(valor) As String
    formatValorParaExibir = Format(valor, "##,#0.00")
End Function

Public Function Replace(Texto As String, caracter As String, caracterParaSubstituir As String) As String
    Do While Texto Like "*" & caracter & "*"
        Texto = left$(Texto, (InStr(Texto, caracter) - 1)) _
        & caracterParaSubstituir _
        & right$(Texto, (Len(Texto) - (InStr(Texto, caracter))))
    Loop
    Replace = Texto
End Function

Private Sub grdProdGarantia_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        quantidadeOriginal = quantidadeOriginal & Chr(KeyAscii)
    ElseIf KeyAscii = 8 Then
        If Len(quantidadeOriginal) <> 0 Then
            quantidadeOriginal = left(quantidadeOriginal, Len(quantidadeOriginal) - 1)
        End If
    ElseIf KeyAscii = 13 Then
    Else
        KeyAscii = 0
        'End If
    End If
End Sub

Private Sub grdProdGarantia_DblClick()
    If grdProdGarantia.Col = wColunaQuantidade Then
        quantidadeOriginal = 0
        grdProdGarantia.Editable = flexEDKbd
    Else
        grdProdGarantia.Editable = flexEDNone
    End If
End Sub

Private Sub grdProdGarantia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim valor As Double
    If grdProdGarantia.Row >= grdProdGarantia.FixedRows Then

        If grdProdGarantia.Col = wColunaGarantia24 And grdProdGarantia.CellBackColor <> wcorCelulaDesativada Then
            If grdProdGarantia = True Then
                grdProdGarantia = False
                contLinhasMarcada = contLinhasMarcada - 1
            Else
                grdProdGarantia = True
                contLinhasMarcada = contLinhasMarcada + 1
                grdProdGarantia.TextMatrix(grdProdGarantia.Row, wColunaGarantia36) = False
            End If
        ElseIf grdProdGarantia.Col = wColunaGarantia36 And grdProdGarantia.CellBackColor <> wcorCelulaDesativada Then
            If grdProdGarantia = True Then
                grdProdGarantia = False: contLinhasMarcada = contLinhasMarcada - 1
            Else
                grdProdGarantia = True
                contLinhasMarcada = contLinhasMarcada + 1
                grdProdGarantia.TextMatrix(grdProdGarantia.Row, wColunaGarantia24) = False
            End If
        End If
        
        exibitPrecoTotal
        exibirPrecoUnitario
        
    End If
End Sub

Private Sub exibirPrecoUnitario()
    lblValorUnitarioItem.Caption = grdProdGarantia.TextMatrix(grdProdGarantia.Row, 7)
End Sub

Private Sub exibitPrecoTotal()

    Dim coluna As Integer
    Dim Linha As Integer
    Dim valor As Double
    
    lblValorTotalGarantia.Caption = RTrim(frmPedido.cmdTotalPedido.Caption)
    'txtTotal.Text = frmPedido.lblTotalPedido

    For coluna = wColunaGarantia24 To wColunaGarantia36 Step 2
        For Linha = grdProdGarantia.FixedRows To grdProdGarantia.Rows - 1
            If grdProdGarantia.TextMatrix(Linha, coluna) = True Then
            
                valor = Format(obterCoeficiente(left(grdProdGarantia.TextMatrix(0, coluna + 1), 2) + 12, grdProdGarantia.TextMatrix(Linha, 7)), "##,#0.00")
                lblValorTotalGarantia.Caption = RTrim(Format(lblValorTotalGarantia.Caption + (valor * grdProdGarantia.TextMatrix(Linha, wColunaQuantidade)), "##,#0.00"))
        
            End If
        Next Linha
    Next coluna

End Sub


Private Sub grdProdGarantia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdProdGarantia
        If (.MouseCol = wColunaGarantia24 + 1 Or .MouseCol = wColunaGarantia36 + 1) Then
            If .MouseRow >= .FixedRows Then
                .ToolTipText = "Valor da Garantia: R$ " & _
                Format(obterCoeficiente(left(.TextMatrix(0, .MouseCol), 2) + 12, _
                .TextMatrix(.MouseRow, 7)), _
                "##,#0.00") & " por produto"
            End If
        ElseIf .MouseCol <> 8 And .MouseCol <> 10 Then
            .ToolTipText = Empty
        End If
    End With
End Sub

Private Sub grdProdGarantia_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    MousePointer = 11
    
    Cancel = Not (validaQuantidade(wNumeroPedido, grdProdGarantia.TextMatrix(Row, 0), quantidadeOriginal))
    quantidadeOriginal = 0
    grdProdGarantia = ""
    
    MousePointer = 0
End Sub

Private Sub atualizaVendedor()
    Dim SQL As String
    
    On Error GoTo trataerro
    
    SQL = "update nfcapa set vendedorGarantia = vendedor where numeroPed = " & wNumeroPedido
    adoCNLoja.Execute SQL

    Exit Sub
trataerro:
    Select Case Err.Number
        Case 40002
            MsgBox "Erro: Não foi possível gravar o vendedor no Banco de Dados da loja", _
            vbCritical, "Atualização vendedor"
        Case Else
            MsgBox "Ocorreu um erro desconhecido durante a execução" & vbNewLine & _
            "Código: " & Err.Number & vbNewLine & "Descrição: " & Err.description, vbCritical, "Atualização vendedor"
            End
    End Select
End Sub

Private Function atualizaItens() As Boolean
    Dim Linha As Integer
    Dim coluna As Integer
    Dim SQL As String
    Dim totalGarantia As Double
    
    atualizaItens = False
    
    For coluna = wColunaGarantia24 To wColunaGarantia36 Step 2
        For Linha = grdProdGarantia.FixedRows To grdProdGarantia.Rows - 1
            If grdProdGarantia.TextMatrix(Linha, coluna) = True Then
            
                SQL = SQL & montaSQLAtualizacaoItens(Linha, coluna)
                totalGarantia = totalGarantia + Replace(grdProdGarantia.TextMatrix(Linha, coluna + 1), ".", "") - _
                    (grdProdGarantia.TextMatrix(Linha, wColunaQuantidade) * Replace(grdProdGarantia.TextMatrix(Linha, 7), ".", ""))
                atualizaItens = True
        
                valorGarantiaEstendida = totalGarantia
        
            End If
        Next Linha
    Next coluna
    
    If atualizaItens Then
        SQL = SQL & "update nfcapa " & vbNewLine & _
              "set garantiaEstendida = 'S', " & _
              "totalGarantia = " & Replace(CStr(totalGarantia), ",", ".") & " " & vbNewLine & _
              "where numeroPed = " & wNumeroPedido
        adoCNLoja.Execute SQL
    End If
    
End Function

Private Function validaQuantidade(NumeroPedido As String, Referencia As String, quantidade As Integer) As Boolean
    Dim rsProdutoGarantia As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "select qtde " & vbNewLine & _
          "from nfitens " & vbNewLine & _
          "where numeroPed = " & NumeroPedido & " and referencia = " & Referencia
    
    rsProdutoGarantia.CursorLocation = adUseClient
    rsProdutoGarantia.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If quantidade <= 0 Or quantidade > rsProdutoGarantia("qtde") Then
            validaQuantidade = False
        Else
            validaQuantidade = True
            atualizaValoresGarantia quantidade, grdProdGarantia.Row
            grdProdGarantia.Editable = flexEDNone
        End If
    
    rsProdutoGarantia.Close
End Function

Private Sub grdProdGarantia_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        If grdProdGarantia.Col = wColunaQuantidade Then
            grdProdGarantia.Editable = flexEDKbd
        Else
            grdProdGarantia.Editable = flexEDNone
        End If
    Else
        grdProdGarantia.Editable = flexEDNone
    End If
End Sub

Private Function formataCampoCalculo(valor As String) As String
    formataCampoCalculo = Replace(Replace(valor, ".", ""), ",", ".")
End Function


Private Sub atualizaValoresGarantia(quantidade As Integer, Linha As Integer)
    Dim precoTotal As Double
    grdProdGarantia.TextMatrix(Linha, wColunaGarantia24 + 1) = formatValorParaExibir(calculoCoeficientePedido(Replace(grdProdGarantia.TextMatrix(Linha, 7), ".", ""), quantidade, 24, Replace(grdProdGarantia.TextMatrix(Linha, 7), ".", "")))
    grdProdGarantia.TextMatrix(Linha, wColunaGarantia36 + 1) = formatValorParaExibir(calculoCoeficientePedido(Replace(grdProdGarantia.TextMatrix(Linha, 7), ".", ""), quantidade, 36, Replace(grdProdGarantia.TextMatrix(Linha, 7), ".", "")))
End Sub


Private Function valGrava(valor As Double) As String
    valGrava = Replace(CStr(valor), ",", ".")
End Function

Private Function montaSQLAtualizacaoItens(Linha, coluna) As String

Dim planoGarantia As Integer
Dim valorGarantia As Double
Dim valorUnitario As Double
Dim coeficientePlano As Double
Dim qtdeGarantia As Double
Dim ValorIOF As Double
Dim PremioLiquido As Double
Dim ValorCustoSeguradora As Double
Dim SQL As String

Dim dataInicioVigencia As String
Dim dataFinalVigencia  As String

'Dim rsFaixaPremio As rdoResultset
    Dim rsFaixaPremio As New ADODB.Recordset

    valorUnitario = grdProdGarantia.TextMatrix(Linha, 7)
    planoGarantia = left(grdProdGarantia.TextMatrix(0, coluna + 1), 2) + 12
    qtdeGarantia = grdProdGarantia.TextMatrix(Linha, wColunaQuantidade)
    valorGarantia = Replace(grdProdGarantia.TextMatrix(Linha, coluna + 1), ".", "") - (qtdeGarantia * valorUnitario)
    'valorGarantia = (obterCoeficiente(planoGarantia, valorUnitario) * valorUnitario)
    coeficientePlano = obterCoeficiente(planoGarantia, valorUnitario)
    
    dataInicioVigencia = DateAdd("D", 1, Date)
    If grdProdGarantia.TextMatrix(Linha, 4) <> " - - " Then dataInicioVigencia = Format(DateAdd("M", grdProdGarantia.TextMatrix(Linha, 4), dataInicioVigencia), "YYYY-MM-DD")
    dataFinalVigencia = Format(DateAdd("M", planoGarantia, Date), "YYYY-MM-DD")
    
    SQL = "select fpg_premioLiquido, fpg_IOF, fpg_premio " & vbNewLine & _
          "from FIN_faixapremioge " & vbNewLine & _
          "where '" & valGrava(valorUnitario) & "' between fpg_faixainicial " & vbNewLine & _
          "and fpg_faixaFinal and fpg_plano = " & planoGarantia
    
rsFaixaPremio.CursorLocation = adUseClient
rsFaixaPremio.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        ValorIOF = rsFaixaPremio("fpg_IOF")
        ValorCustoSeguradora = rsFaixaPremio("fpg_premio")
        PremioLiquido = ValorCustoSeguradora - rsFaixaPremio("fpg_IOF")
    rsFaixaPremio.Close
    
    SQL = "update nfitens " & vbNewLine & _
          "set qtdeGarantia = " & qtdeGarantia & ", " & vbNewLine & _
          "coeficientePlano = " & valGrava(coeficientePlano) & ", " & _
          "garantiaEstendida = 'S', " & _
          "planoGarantia = " & planoGarantia & ", " & vbNewLine & _
          "ValorGarantia = " & valGrava(valorGarantia) & ", " & _
          "ge_premioLiquido = " & valGrava(PremioLiquido) & ", " & vbNewLine & _
          "ge_iof = " & valGrava(ValorIOF) & ", " & vbNewLine & _
          "ge_dataInicioVigencia = '" & dataInicioVigencia & "', " & vbNewLine & _
          "ge_dataFinalVigencia = '" & dataFinalVigencia & "', " & vbNewLine & _
          "ge_valorCustoSeguradora = " & valGrava(ValorCustoSeguradora) & " " & vbNewLine & _
          "where numeroPed = " & wNumeroPedido & " and referencia = " & grdProdGarantia.TextMatrix(Linha, 0)
          
    montaSQLAtualizacaoItens = SQL & vbNewLine & vbNewLine
    
End Function
