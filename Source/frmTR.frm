VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTR 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "T  R"
   ClientHeight    =   5610
   ClientLeft      =   6885
   ClientTop       =   2655
   ClientWidth     =   6660
   ControlBox      =   0   'False
   FillColor       =   &H000703BA&
   ForeColor       =   &H000703BA&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   6360
      TabIndex        =   6
      Top             =   4830
      Width           =   6360
   End
   Begin VB.TextBox txtPercentualTR 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   795
      TabIndex        =   5
      Top             =   195
      Width           =   1050
   End
   Begin VB.TextBox txtDescricaoProdutoTR 
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
      Height          =   405
      Left            =   90
      MaxLength       =   38
      TabIndex        =   3
      Top             =   5010
      Width           =   5235
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdValorTR 
      Height          =   600
      Left            =   1965
      TabIndex        =   1
      Top             =   90
      Width           =   4545
      _cx             =   8017
      _cy             =   1058
      _ConvInfo       =   1
      Appearance      =   1
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
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTR.frx":0000
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensProduto 
      Height          =   3930
      Left            =   135
      TabIndex        =   0
      Top             =   765
      Width           =   6360
      _cx             =   11218
      _cy             =   6932
      _ConvInfo       =   1
      Appearance      =   1
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTR.frx":0092
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
      ForeColorFrozen =   67372047
      WallPaperAlignment=   0
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
      Left            =   15
      TabIndex        =   2
      Text            =   "0"
      Top             =   2940
      Visible         =   0   'False
      Width           =   270
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   330
      OleObjectBlob   =   "frmTR.frx":013D
      Top             =   2940
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5385
      TabIndex        =   7
      Top             =   5025
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
      MICON           =   "frmTR.frx":0371
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPerTr 
      BackStyle       =   0  'Transparent
      Caption         =   "% TR"
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
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   285
      Width           =   600
   End
End
Attribute VB_Name = "frmTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCodigo As Integer
Dim wSequencia As Double
Dim wValorCampo As String
Dim wDesconto As Double
Dim wValorTR As Double
Dim wTotalPedido As Double
Dim wValorSN As Double
Dim wReferenciaAlternativa As String
Dim Index As Integer
Dim SQL As String
Dim DescricaoAlternativaEmBranco As Boolean

Private Sub cmdGrava_Click()
  Call VerificarDescricaoAlternativa
  If DescricaoAlternativaEmBranco = True Then
     txtDescricaoProdutoTR.SetFocus
     txtDescricaoProdutoTR.Text = "Falta preencher descrição Alternativa"
     txtDescricaoProdutoTR.SelStart = 0
     txtDescricaoProdutoTR.SelLength = Len(txtDescricaoProdutoTR.Text)
     txtDescricaoProdutoTR.SetFocus
     Exit Sub
  End If

On Error GoTo FinalizaTR
  
  wCodigo = 1
  wSequencia = 19
  wValorCampo = "TR"
  
  wDesconto = grdValorTR.TextMatrix(1, 1)
  wValorTR = grdValorTR.TextMatrix(1, 2)
  wValorSN = grdValorTR.TextMatrix(1, 3)
  
  SQL = ""
  SQL = "ValorMercadoriaAlternativa = " & ConverteVirgula(Format((wDesconto + wValorTR + wValorSN), "###,###,###,##0.00")) & ", " & _
        "ValorTotalCodigoZero = " & ConverteVirgula(Format(wValorTR, "##0.00")) & ", TotalNotaAlternativa = " & ConverteVirgula(Format(wValorSN, "##0.00"))
  
  adoCNLoja.BeginTrans
  adoCNLoja.Execute "Exec SP_GravaComplementoVenda " & txtpedido.Text & ",1,1,'" & SQL & "'"
  adoCNLoja.CommitTrans
  
 ' frmPedido.cmdBotoes(6).Enabled = False
 ' frmPedido.cmdBotoes(8).Visible = False
 ' frmPedido.cmdTR.visible = False
  Unload Me
  Exit Sub
  
FinalizaTR:
  MsgBox "Não foi possível finalizar o TR." & vbLf & _
         Err.Number & " - " & Err.description, vbInformation, "Atenção"
  adoCNLoja.RollbackTrans
  Exit Sub

End Sub

Private Sub cmdGrava_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  Call VerificarDescricaoAlternativa
  
  If DescricaoAlternativaEmBranco = True Then
     txtDescricaoProdutoTR.SetFocus
     txtDescricaoProdutoTR.Text = "Falta preencher descrição Alternativa"
     txtDescricaoProdutoTR.SelStart = 0
     txtDescricaoProdutoTR.SelLength = Len(txtDescricaoProdutoTR.Text)
     txtDescricaoProdutoTR.SetFocus
     Exit Sub
  End If
End If

If KeyAscii = 27 Then
   txtPercentualTR.SetFocus
   txtPercentualTR.SelStart = 0
   txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
   txtPercentualTR.SetFocus
   
End If
End Sub

Private Sub cmdRetorna_Click()

 Unload Me
 frmPedido.txtPesquisar.SetFocus

 
End Sub

Private Sub cmdRetorna_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
 
 Call AjustaTela(frmTR)
 
  grdValorTR.Rows = 2
  txtpedido.Text = frmPedido.txtpedido.Text
  
  txtDescricaoProdutoTR.Enabled = False
  
  Call LimpaTR
  Call SomaItensVenda
  Call CarregaDesconto
  Call CarregaItens
  

  grdItensProduto.WallPaper = frmPedido.grdItensProduto.WallPaper
  
  grdItensProduto.Row = 1
  txtDescricaoProdutoTR.Text = grdItensProduto.TextMatrix(grdItensProduto.Row, 1)
  
  txtPercentualTR.SelStart = 0
  txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
  

End Sub

Private Sub SomaItensVenda()

'******************* NFItens
  SQL = "Select  TipoNota, sum(VLUNIT * QTDE) as TotalVenda," _
        & "count(*) as TotalItens From NFItens Where NumeroPed = " & txtpedido.Text & " and " _
        & "TipoNota = 'PD' Group By TipoNota"
  
  rsSomaItens.CursorLocation = adUseClient
  rsSomaItens.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  grdValorTR.TextMatrix(1, 0) = Format(rsSomaItens("TotalVenda"), "###,###,##0.00")
  
End Sub
Private Sub CarregaDesconto()
   
'******************* NFCapa
   SQL = ""
   SQL = "Select * From NFCapa Where NumeroPed = " & txtpedido.Text & " and " _
         & "TipoNota = 'PD'"
   
   rsComplementoVenda.CursorLocation = adUseClient
   rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   
   If rsComplementoVenda.EOF Then
      wDesconto = 0
      grdValorTR.TextMatrix(1, 1) = Format(wDesconto, "###,###,##0.00")
      grdValorTR.TextMatrix(1, 3) = Format((rsSomaItens("TotalVenda") - wDesconto), "###,###,##0.00")

   Else
      wDesconto = Format(rsComplementoVenda("Desconto"), "###,###,##0.00")
      grdValorTR.TextMatrix(1, 1) = Format(wDesconto, "###,###,##0.00")
      grdValorTR.TextMatrix(1, 3) = Format((rsSomaItens("TotalVenda") - wDesconto), "###,###,##0.00")
   End If
   rsComplementoVenda.Close

End Sub

Private Sub CarregaItens()

  grdItensProduto.Rows = 1
  
  SQL = "Select NFItens.*, PR_Descricao From NFItens,ProdutoLoja " _
        & "Where Referencia = PR_Referencia and NumeroPed = " & txtpedido.Text _
        & " Order By Item"
   
   rsComplementoVenda.CursorLocation = adUseClient
   rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   If Not rsComplementoVenda.EOF Then
      Do While Not rsComplementoVenda.EOF
          grdItensProduto.AddItem rsComplementoVenda("Referencia") & Chr(9) _
          & rsComplementoVenda("PR_Descricao")
            rsComplementoVenda.MoveNext
      Loop
   End If
   rsComplementoVenda.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
        rsSomaItens.Close
End Sub

Private Sub grdItensProduto_Click()
'  txtDescricaoProdutoTR.Text = grdItensProduto.TextMatrix(grdItensProduto.Row, 1)
'  txtDescricaoProdutoTR.SetFocus
'  txtDescricaoProdutoTR.SelStart = 0
'  txtDescricaoProdutoTR.SelLength = Len(txtDescricaoProdutoTR.Text)
'  txtDescricaoProdutoTR.SetFocus
    
End Sub

Private Sub grdItensProduto_EnterCell()
  If txtPercentualTR.Text <> "" Then
    txtDescricaoProdutoTR.Text = grdItensProduto.TextMatrix(grdItensProduto.Row, 1)
    wReferenciaAlternativa = RTrim(LTrim(grdItensProduto.TextMatrix(grdItensProduto.Row, 0)))
  End If
End Sub


Private Sub grdItensProduto_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = vbKeyF2 Then
  cmdGrava_Click
 ElseIf KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub

Private Sub grdValorTR_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = vbKeyF2 Then
  cmdGrava_Click
 ElseIf KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub

Private Sub txtDescricaoProdutoTR_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = vbKeyF2 Then
  cmdGrava_Click
 ElseIf KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub

Private Sub txtPercentualTR_Change()
    If IsNumeric(txtPercentualTR.Text) = False Then
        txtPercentualTR.Text = ""
        txtPercentualTR.SelStart = 0
        txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
        txtPercentualTR.SetFocus
    ElseIf txtPercentualTR.Text <= 0 Then
        txtPercentualTR.Text = ""
        txtPercentualTR.SelStart = 0
        txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
        txtPercentualTR.SetFocus
    End If
End Sub

Private Sub txtPercentualTR_GotFocus()
txtPercentualTR.SetFocus
txtPercentualTR.SelStart = 0
txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
End Sub

Private Sub txtPercentualTR_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   If Not IsNumeric(txtPercentualTR) Then
      txtPercentualTR.SetFocus
      txtPercentualTR.SelStart = 0
      txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
      Exit Sub
   End If
 
   If Numeros(txtPercentualTR) = "" Then
      txtPercentualTR.SetFocus
      txtPercentualTR.SelStart = 0
      txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
      Exit Sub
   End If
    
   If txtPercentualTR.Text > 100 Then
      txtPercentualTR.SetFocus
      txtPercentualTR.SelStart = 0
      txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
      Exit Sub
   End If
   
End If

If KeyCode = 13 Then
   If IsNumeric(txtPercentualTR.Text) Then
      Call CalculaTR
      grdItensProduto.Enabled = True
      cmdGrava.Enabled = True
      cmdGrava.SetFocus
      grdItensProduto.SetFocus
   Else
      txtPercentualTR.Text = ""
      txtPercentualTR.SetFocus
      Exit Sub
   End If
Else
      txtPercentualTR.SetFocus
End If
End Sub
Private Sub txtPercentualTR_KeyPress(KeyAscii As Integer)

If KeyAscii = 46 Then
   txtPercentualTR.Text = ""
   txtPercentualTR.SelStart = 0
   txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
   txtPercentualTR.SetFocus
   Exit Sub
End If

If KeyAscii = 44 Then
   txtPercentualTR.Text = ""
   txtPercentualTR.SelStart = 0
   txtPercentualTR.SelLength = Len(txtPercentualTR.Text)
   txtPercentualTR.SetFocus
   Exit Sub
End If

 
 If KeyAscii = vbKeyF2 Then
  cmdGrava_Click
 ElseIf KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub

Private Sub txtPercentualTR_LostFocus()
    txtDescricaoProdutoTR.Enabled = True
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPercentualTR.Enabled = True
   txtPercentualTR.SetFocus
   Exit Sub
End If


End Sub

Private Sub txtDescricaoProdutoTR_GotFocus()
  txtDescricaoProdutoTR.SelStart = 0
  txtDescricaoProdutoTR.SelLength = Len(txtDescricaoProdutoTR.Text)
  txtDescricaoProdutoTR.SetFocus
End Sub

Private Sub txtDescricaoProdutoTR_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo erronaUpdate

If KeyCode = 13 Then
   ChecaCaracterDigitado (txtDescricaoProdutoTR)
   If CaracterDigitado = True Then
      txtDescricaoProdutoTR.SelStart = 0
      txtDescricaoProdutoTR.SelLength = Len(txtDescricaoProdutoTR.Text)
      txtDescricaoProdutoTR.SetFocus
      Exit Sub
   End If
End If

If KeyCode = 13 Then
   If Mid(txtDescricaoProdutoTR, 1, 15) = "Falta preencher" Then
      txtDescricaoProdutoTR.Text = ""
      txtDescricaoProdutoTR.SetFocus
      Exit Sub
   End If
End If

If KeyCode = 13 Then
   If grdItensProduto.Row = 0 Then
      Exit Sub
   End If
End If

If KeyCode = 13 And Trim(txtDescricaoProdutoTR.Text) = Trim(grdItensProduto.TextMatrix(grdItensProduto.Row, 1)) Then
   MsgBox "Verifique a descrição alternativa do produto." & vbLf & _
          "A descrição alternativa tem que ser diferente da descrição.", vbInformation, "Atenção"
      
   txtDescricaoProdutoTR.SelStart = 0
   txtDescricaoProdutoTR.SelLength = Len(txtDescricaoProdutoTR.Text)
   txtDescricaoProdutoTR.SetFocus
   Exit Sub
   
End If

If KeyCode = 13 Then
  
    If txtDescricaoProdutoTR.Text <> "" Then
   
    adoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    
    grdItensProduto.TextMatrix(grdItensProduto.Row, 2) = txtDescricaoProdutoTR.Text
    
    wReferenciaAlternativa = Mid(wReferenciaAlternativa, 4, Len(wReferenciaAlternativa)) & _
                            Mid(wReferenciaAlternativa, 1, 3)
    
    SQL = "UPDATE NFItens Set ReferenciaAlternativa = '" & wReferenciaAlternativa & "', " & _
          "DescricaoAlternativa = '" & txtDescricaoProdutoTR.Text & "' Where NumeroPed = " & txtpedido.Text & " and " & _
          "Referencia = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "'"
        
    adoCNLoja.Execute (SQL)
    Screen.MousePointer = vbNormal
    adoCNLoja.CommitTrans
    Exit Sub
  Else
     Exit Sub
  End If
Else
     txtDescricaoProdutoTR.SetFocus
     Exit Sub
End If


erronaUpdate:
MsgBox "Erro na Atualização " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal

End Sub
Private Sub CalculaTR()
On Error GoTo erronoUpdate1
  wValorTR = (grdValorTR.TextMatrix(1, 0) * txtPercentualTR / 100)
  wTotalPedido = grdValorTR.TextMatrix(1, 0)
  grdValorTR.TextMatrix(1, 2) = Format(wValorTR, "###,###,##0.00")
  wValorSN = (wTotalPedido - (wValorTR + wDesconto))
  grdValorTR.TextMatrix(1, 3) = Format(wValorSN, "###,###,##0.00")
  
  adoCNLoja.BeginTrans
  Screen.MousePointer = vbHourglass
      
  SQL = "UPDATE NFItens Set PrecoUnitAlternativa = (VLUnit - ((VLUnit * " & ConverteVirgula(txtPercentualTR.Text) & ") / 100)) " & _
        "WHERE TipoNota = 'PD' and NumeroPed = " & txtpedido.Text
      
      adoCNLoja.Execute (SQL)
      Screen.MousePointer = vbNormal
      adoCNLoja.CommitTrans
      Exit Sub
      
erronoUpdate1:
    MsgBox "Erro na Atualização " & Err.description, vbCritical, "Aviso"
    adoCNLoja.RollbackTrans
    Screen.MousePointer = vbNormal
End Sub
Sub VerificarDescricaoAlternativa()
    Dim Linha As Integer
    
    DescricaoAlternativaEmBranco = False
    For Linha = 1 To grdItensProduto.Rows - 1
        If grdItensProduto.TextMatrix(Linha, 2) = "" Then
           DescricaoAlternativaEmBranco = True
        End If
    Next
    Exit Sub
End Sub


Private Sub DesfazProcessoTR()
SQL = "Select * from ComplementoVenda where COV_NumeroPedido = " & txtpedido.Text _
      & " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " _
      & wSequencia
      rsComplementoVenda.CursorLocation = adUseClient
      rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic


  On Error GoTo ErronaDelecao
  
  If rsComplementoVenda.EOF Then
     rsComplementoVenda.Close
     Exit Sub
  Else
     adoCNLoja.BeginTrans
     Screen.MousePointer = vbHourglass
     SQL = ""
     SQL = "Delete ComplementoVenda " _
           & " Where COV_Numeropedido = " & txtpedido.Text & _
           " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " & wSequencia
            adoCNLoja.Execute SQL
            Screen.MousePointer = vbNormal
            adoCNLoja.CommitTrans
            rsComplementoVenda.Close
            
            Exit Sub
  End If
  
ErronaDelecao:
MsgBox "Erro ao Deletar Complemento de Venda Sequencia ==> " & wSequencia & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
rsComplementoVenda.Close
End Sub


