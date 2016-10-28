VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7D.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmPagamento 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Pagamento"
   ClientHeight    =   5055
   ClientLeft      =   2475
   ClientTop       =   5925
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5475
      OleObjectBlob   =   "frmPagamento.frx":0000
      Top             =   135
   End
   Begin VB.TextBox txtTotalPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
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
      Left            =   2565
      TabIndex        =   6
      Text            =   "0"
      Top             =   975
      Visible         =   0   'False
      Width           =   1875
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
      Left            =   5055
      TabIndex        =   5
      Text            =   "0"
      Top             =   435
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtValorEntrada 
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
      Left            =   1965
      TabIndex        =   1
      Text            =   "0"
      Top             =   1500
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdCondicaoFaturado 
      Height          =   2910
      Left            =   60
      TabIndex        =   0
      Top             =   2025
      Visible         =   0   'False
      Width           =   4575
      _cx             =   8070
      _cy             =   5133
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16448250
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483639
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483639
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483647
      GridColorFixed  =   -2147483647
      TreeColor       =   -2147483639
      FloodColor      =   192
      SheetBorder     =   -2147483639
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
      FormatString    =   $"frmPagamento.frx":0234
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Project1.chameleonButton cmdFaturado 
      Height          =   690
      Left            =   1305
      TabIndex        =   2
      Top             =   735
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1217
      BTYPE           =   14
      TX              =   "FATURADO"
      ENAB            =   0   'False
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
      BCOL            =   8454143
      BCOLO           =   8454143
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagamento.frx":029D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdFinanciado 
      Height          =   690
      Left            =   60
      TabIndex        =   3
      Top             =   735
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1217
      BTYPE           =   14
      TX              =   "FINANCIADO"
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
      BCOL            =   8438015
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagamento.frx":02B9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdAvista 
      Height          =   675
      Left            =   60
      TabIndex        =   4
      Top             =   45
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   1191
      BTYPE           =   14
      TX              =   "A VISTA"
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
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagamento.frx":02D5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblEntrada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E n t r a d a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   7
      Top             =   1560
      Width           =   1635
   End
End
Attribute VB_Name = "frmPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wPagamento As Integer
Dim wCodigo As Integer
Dim wSequencia As Integer
Dim wValorCampo As String
Dim SQL As String
Dim wValorentrada As Double
Dim wTotalpedido As Double

Private Sub cmdAvista_Click()
  wValorCampo = 1
'  rdoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & "," & wCodigo & "," & wSequencia & ",'" & wValorCampo & "'" ', rdExecDirect
  SQL = ""
  SQL = "CondPag = " & wValorCampo
  rdoCNLoja.Execute "Exec SP_GravaComplementoVenda '" & txtPedido.Text & "',1," & wSequencia & ",'" & SQL & "'"
  
  frmPedido.cmdTransferencia.Enabled = False
  frmPedido.cmdPagamento.Enabled = False
  Unload Me
 End Sub

Private Sub cmdAvista_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub cmdFaturado_Click()
'   frmPagamento.Width = 4900
   cmdFaturado.Width = 4560
   cmdFaturado.Top = 135
   cmdFaturado.Left = 105
   cmdAvista.Visible = False
   cmdFinanciado.Visible = False
   LblEntrada.Visible = True
   txtValorEntrada.Visible = True
   txtValorEntrada.SetFocus
'   frmPagamento.Height = 2100
'   Left = (Screen.Width - Width) / 2
'   Top = (Screen.Height - Height) / 2
   wValorCampo = 4
End Sub

Private Sub cmdFaturado_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub cmdFinanciado_Click()
'   frmPagamento.Width = 4900
   cmdFinanciado.Width = 4560
   
   cmdFinanciado.Top = 135
   cmdFinanciado.Left = 105
   cmdAvista.Visible = False
   cmdFaturado.Visible = False
   LblEntrada.Visible = True
   txtValorEntrada.Visible = True
   txtValorEntrada.SetFocus
'   frmPagamento.Height = 2100
  
'   Left = (Screen.Width - Width) / 2
'   Top = (Screen.Height - Height) / 2
   wValorCampo = 3
End Sub

Private Sub cmdFinanciado_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
  txtTotalPedido.Text = frmPedido.txtTotalPedido.Text
  txtPedido.Text = frmPedido.txtPedido.Text
  cmdFaturado.Enabled = True

'  Skin1.LoadSkin "c:\WINDOWS\system\skin.skn"
'  Skin1.ApplySkin Me.hwnd
'  frmPagamento.Height = 2100
  LblEntrada.Top = 1000
  LblEntrada.Left = 105
  txtValorEntrada.Top = 930
  grdCondicaoFaturado.Top = 1620
  txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
  
 wSequencia = 6
 wCodigo = 1
 
 
 SQL = "Select * from ComplementoVenda where COV_NumeroPedido = " & txtPedido.Text _
      & " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " _
      & wSequencia
      rsComplementoVenda.CursorLocation = adUseClient
      rsComplementoVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 
      If rsComplementoVenda.EOF = False Then
         If Trim(rsComplementoVenda("COV_ValorComplemento")) > 899999 Then
            cmdFaturado.Enabled = False
         Else
            cmdFaturado.Enabled = True
         End If
      Else
          cmdFaturado.Enabled = True
      End If
   
 
  rsComplementoVenda.Close
  
  
  wCodigo = 1
  wSequencia = 4
  
  Call AjustaTela(frmPagamento)
  
'  Left = (Screen.Width - Width) / 2
'  Top = (Screen.Height - Height) / 2
  
End Sub

Private Sub grdCondicaoFaturado_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     wValorCampo = grdCondicaoFaturado.TextMatrix(grdCondicaoFaturado.Row, 0)
'     rdoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & "," & wCodigo & "," & wSequencia & ",'" & wValorCampo & "'" ', rdExecDirect
     If txtValorEntrada.Text <> 0 Then
        wCodigo = 1
        wSequencia = 17
        wValorCampo = txtValorEntrada.Text
        SQL = ""
        SQL = "CondPag = " & grdCondicaoFaturado.TextMatrix(grdCondicaoFaturado.Row, 0) & ",PGEntra = " & ConverteVirgula(wValorCampo)
        
        rdoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & "," & wCodigo & "," & wSequencia & ",'" & SQL & "'" ', rdExecDirect
     frmPedido.cmdTransferencia.Enabled = False
     frmPedido.cmdPagamento.Enabled = False
     End If
     Unload Me
   End If
End Sub

Private Sub grdCondicaoFaturado_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   txtValorEntrada.SelStart = 0
   txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
   txtValorEntrada.SetFocus
End If
End Sub

Private Sub txtValorEntrada_Change()
    If IsNumeric(txtValorEntrada.Text) = False Then
        txtValorEntrada.Text = 0
        txtValorEntrada.SelStart = 0
        txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
        txtValorEntrada.SetFocus
    ElseIf txtValorEntrada.Text <= 0 Then
           txtValorEntrada.SelStart = 0
           txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
           txtValorEntrada.SetFocus
           txtValorEntrada.Text = 0
    End If
End Sub

Private Sub txtValorEntrada_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 27 Then
      Unload Me
   End If
   
   If KeyAscii = 46 Then
      txtValorEntrada.Text = 0
      txtValorEntrada.SelStart = 0
      txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
      txtValorEntrada.SetFocus
      Exit Sub
   End If
   
   If KeyAscii = 13 Then
  
   wValorentrada = Format(txtValorEntrada.Text, "###,###,###,##0.00")
   wTotalpedido = Format(txtTotalPedido.Text, "###,###,###,##0.00")
   
    If wValorentrada >= wTotalpedido Then
       MsgBox "Valor da Entrada não pode ser Igual ou Maior que o Pedido", vbCritical, "Atenção"
       txtValorEntrada.SelStart = 0
       txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
       txtValorEntrada.SetFocus
       Exit Sub
   End If
  
   If IsNumeric(txtValorEntrada.Text) = False Then
        txtValorEntrada.Text = 0
        txtValorEntrada.SelStart = 0
        txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
        txtValorEntrada.SetFocus
        
    ElseIf txtValorEntrada.Text <= 0 Then
           txtValorEntrada.SelStart = 0
           txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
           txtValorEntrada.SetFocus
           txtValorEntrada.Text = 0
           Exit Sub
    End If
     
     If IsNumeric(Trim(txtValorEntrada.Text)) Then
        If wValorCampo = 3 Then
           txtValorEntrada.Text = Format(txtValorEntrada.Text, "###,###,###,##0.00")
'           rdoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & "," & wCodigo & "," & wSequencia & ",'" & wValorCampo & "'" ', rdExecDirect
           If txtValorEntrada.Text <> 0 Then
              wCodigo = 1
              wSequencia = 17
              wValorCampo = txtValorEntrada.Text
              SQL = ""
              SQL = "CondPag = 3, PGEntra = " & ConverteVirgula(wValorCampo)
              rdoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & "," & wCodigo & "," & wSequencia & ",'" & SQL & "'" ', rdExecDirect
              frmPedido.cmdTransferencia.Enabled = False
              frmPedido.cmdPagamento.Enabled = False
           End If
           Unload Me
        ElseIf wValorCampo = 4 Then
               txtValorEntrada.Text = Format(txtValorEntrada.Text, "###,###,###,##0.00")
'               frmPagamento.Height = 5365
               grdCondicaoFaturado.Visible = True
               grdCondicaoFaturado.Enabled = True
               grdCondicaoFaturado.Redraw = True
               grdCondicaoFaturado.SetFocus
               Call CarregaCondicaoFaturado
        End If
     Else
        txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
        txtValorEntrada.SetFocus
     End If
   End If
End Sub

Private Sub txtValorEntrada_LostFocus()
  
  If IsNumeric(txtValorEntrada.Text) = False Then
        txtValorEntrada.Text = 0
        txtValorEntrada.SelStart = 0
        txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
        txtValorEntrada.SetFocus
    ElseIf txtValorEntrada.Text <= 0 Then
           txtValorEntrada.SelStart = 0
           txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
           txtValorEntrada.SetFocus
           txtValorEntrada.Text = 0
    End If
   
    
    wValorentrada = Format(txtValorEntrada.Text, "###,###,###,##0.00")
    wTotalpedido = Format(txtTotalPedido.Text, "###,###,###,##0.00")
   
    If wValorentrada >= wTotalpedido Then
       MsgBox "Valor da Entrada não pode ser Igual ou Maior que o Pedido", vbCritical, "Atenção"
       txtValorEntrada.SelStart = 0
       txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
       txtValorEntrada.SetFocus
   Else
     
      txtValorEntrada.Text = Format(txtValorEntrada.Text, "###,###,###,##0.00")
   End If
End Sub
Private Sub CarregaCondicaoFaturado()
  grdCondicaoFaturado.Rows = 1
  SQL = "Select * from CondicaoPagamento where CP_Tipo = 'FA' Order By CP_Condicao"
  rsCondicaoFaturado.CursorLocation = adUseClient
  rsCondicaoFaturado.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsCondicaoFaturado.EOF Then
     Do While Not rsCondicaoFaturado.EOF
        grdCondicaoFaturado.AddItem rsCondicaoFaturado("CP_Codigo") & Chr(9) _
        & rsCondicaoFaturado("CP_Condicao") & Chr(9) _
        & rsCondicaoFaturado("CP_Parcelas")
        rsCondicaoFaturado.MoveNext
     Loop
     grdCondicaoFaturado.Redraw = True
     grdCondicaoFaturado.SetFocus
     grdCondicaoFaturado.Row = 1
  End If
  rsCondicaoFaturado.Close
End Sub

