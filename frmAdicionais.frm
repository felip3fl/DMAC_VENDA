VERSION 5.00
Begin VB.Form frmAdicionais 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Finaliza pedido"
   ClientHeight    =   6465
   ClientLeft      =   5520
   ClientTop       =   2145
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00505050&
      Caption         =   "Frete"
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   3480
      TabIndex        =   17
      Top             =   3090
      Width           =   2865
      Begin VB.OptionButton optFreteDestinatario 
         BackColor       =   &H00505050&
         Caption         =   "Destinatário"
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
         Height          =   375
         Left            =   195
         TabIndex        =   5
         Top             =   345
         Width           =   1845
      End
      Begin VB.OptionButton optFreteEmitente 
         BackColor       =   &H00505050&
         Caption         =   "Emitente"
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
         Height          =   375
         Left            =   195
         TabIndex        =   6
         Top             =   750
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   15
      Top             =   4875
      Width           =   6165
   End
   Begin VB.Frame fraPagamento 
      BackColor       =   &H00505050&
      ForeColor       =   &H00FFFFFF&
      Height          =   2325
      Left            =   150
      TabIndex        =   11
      Top             =   570
      Width           =   6195
      Begin VB.Frame Frame1 
         BackColor       =   &H00505050&
         ForeColor       =   &H00FFFFFF&
         Height          =   1050
         Left            =   120
         TabIndex        =   16
         Top             =   1095
         Width           =   5940
         Begin VB.TextBox txtEntrada 
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
            Left            =   150
            MaxLength       =   10
            TabIndex        =   2
            Top             =   555
            Width           =   1260
         End
         Begin VB.CheckBox chkEntrada 
            BackColor       =   &H00505050&
            Caption         =   "Entrada"
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
            Height          =   315
            Left            =   150
            MaskColor       =   &H00343434&
            TabIndex        =   7
            Top             =   195
            Width           =   1365
         End
      End
      Begin VB.TextBox txtCodigoCliente 
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
         Left            =   135
         TabIndex        =   0
         ToolTipText     =   " "
         Top             =   615
         Width           =   1125
      End
      Begin VB.TextBox txtNomeCliente 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   1350
         TabIndex        =   8
         Top             =   615
         Width           =   4710
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Height          =   390
         Left            =   135
         TabIndex        =   12
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.Frame fraVolume 
      BackColor       =   &H00505050&
      Caption         =   "Volume"
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   150
      TabIndex        =   10
      Top             =   3090
      Width           =   3180
      Begin VB.TextBox txtPesoVolume 
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
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   4
         Top             =   675
         Width           =   1320
      End
      Begin VB.TextBox txtQtdeVolume 
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
         Left            =   165
         MaxLength       =   10
         TabIndex        =   3
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Peso"
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
         Height          =   390
         Left            =   1635
         TabIndex        =   14
         Top             =   345
         Width           =   930
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Height          =   390
         Left            =   180
         TabIndex        =   13
         Top             =   345
         Width           =   1470
      End
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5280
      TabIndex        =   1
      Top             =   5040
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
      MICON           =   "frmAdicionais.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPagamento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pagamento"
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
      Left            =   180
      TabIndex        =   9
      Top             =   120
      Width           =   6300
   End
End
Attribute VB_Name = "frmAdicionais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim wTipoFrete As Integer
Dim wCodigoCliente As String

Private Sub chkEntrada_Click()
 If chkEntrada.Value = 0 Then
    txtEntrada.Visible = False

 ElseIf chkEntrada.Value = 1 Then
    txtEntrada.Visible = True

 End If
End Sub

Private Sub cmdGrava_Click()
    
    If txtCodigoCliente.Text = "" Then
        MsgBox "Codigo Invalido", vbCritical, "Atenção"
        Exit Sub
    End If
    If wValor > 10000 And txtCodigoCliente = "999999" Then
        MsgBox "Não é permitido cliente consumidor para vendas maiores que R$10.000,00"
        txtCodigoCliente.Text = ""
        txtNomeCliente.Text = ""
        txtCodigoCliente.SetFocus
        Exit Sub
    End If
     If txtQtdeVolume.Text = "" Or IsNumeric(txtQtdeVolume.Text) = False Or txtQtdeVolume.Text = "0" Then
            txtQtdeVolume.SelStart = 0
            txtQtdeVolume.SelLength = Len(txtQtdeVolume.Text)
            MsgBox "Informar quantidade."
            txtQtdeVolume.SetFocus
            Exit Sub
     End If
         
     If txtPesoVolume.Text = "" Or IsNumeric(txtPesoVolume.Text) = False Or txtPesoVolume.Text = "0" Then
            txtPesoVolume.SelStart = 0
            txtPesoVolume.SelLength = Len(txtPesoVolume.Text)
            MsgBox "Informar peso."
            txtPesoVolume.SetFocus
            Exit Sub
     End If
     
     
     If optFreteDestinatario.Value = True Then
        wTipoFrete = 1
     Else
        wTipoFrete = 0
     End If


      SQL = ""
      SQL = "Update NFCapa set pgentra = " & ConverteVirgula(Format(txtEntrada.Text, "###,##0.00")) & _
            ", cliente = " & txtCodigoCliente.Text & ", tipofrete = " & wTipoFrete & _
            ", pesoLq = " & ConverteVirgula(txtPesoVolume.Text) & ", pesoBr = " & ConverteVirgula(txtPesoVolume.Text) & _
            ", volume = " & ConverteVirgula(txtQtdeVolume.Text) & _
            " Where Numeroped = " & frmPedido.txtPedido.Text
      adoCNLoja.Execute (SQL)

Unload Me

            'frmPedido.cmdBotoes(0).Visible = True
            frmPedido.txtPesquisar.SelStart = 0
            frmPedido.txtPesquisar.SelLength = Len(frmPedido.txtPesquisar.Text)
            frmPedido.cmbPedido.Visible = False
            frmPedido.cmdBotoes(1).Visible = False
            frmPedido.cmdBotoes(4).Visible = False
            frmPedido.cmdBotoes(11).Visible = False
            frmPedido.cmdFechaPedido.Visible = True
            frmPedido.cmdBotoes(0).Visible = False
            'frmPedido.cmdBotoes(13).Visible = False
            frmPedido.cmdFechaPedido.left = frmPedido.cmdBotoes(0).left
            frmPedido.cmdBotoes(2).Visible = True
            frmPedido.cmdBotoes(12).Visible = True
            frmPedido.cmdBotoes(9).Visible = True
            frmPedido.cmdBotoes(8).Visible = True
            frmPedido.cmdBotoes(6).Visible = True
            
            frmPedido.cmdBotoes(10).Visible = True
            frmPedido.cmdBotoes(7).Visible = True

  End Sub

Private Sub cmdRetorna_Click()
  Unload Me
End Sub

Private Sub Form_Activate()

  txtEntrada.Visible = False
  optFreteDestinatario.Value = False
  optFreteEmitente.Value = True
  
  SQL = ""
  SQL = "select nfcapa.CONDPAG as condpag from nfcapa " & _
        "where nfcapa.numeroped = " & frmPedido.txtPedido.Text
  rsCliente.CursorLocation = adUseClient
  rsCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

  
  If Val(rsCliente("condpag")) > 1 Then
         'lblPagamento.Caption = frmPedido.grdPrecos.TextMatrix(0, 0)
         chkEntrada.Visible = True
         chkEntrada.Value = 0
  Else
     chkEntrada.Visible = False
  End If
  
  rsCliente.Close
    
  SQL = ""
  SQL = "select nfcapa.cliente as Codigo, fin_cliente.ce_razao as Nome, nfcapa.pesolq, nfcapa.volume, nfcapa.pgentra, " & _
        "nfcapa.garantiaEstendida as GE, nfcapa.CONDPAG " & _
        "from fin_cliente, nfcapa " & _
        "where nfcapa.cliente = fin_cliente.ce_codigocliente and nfcapa.numeroped = " & frmPedido.txtPedido.Text
        '''AQUI ERRO
            rsCliente.CursorLocation = adUseClient
            rsCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
  If Not rsCliente.EOF Then
    txtQtdeVolume.Text = rsCliente("volume")
    txtPesoVolume.Text = rsCliente("pesoLq")
    txtEntrada.Text = rsCliente("pgentra")
    
    ''AQUI'''''

    If rsCliente("Codigo") > 0 And rsCliente("Codigo") <= 999999 And rsCliente("Codigo") <> 888888 And rsCliente("GE") <> "S" And rsCliente("CONDPAG") <= 2 Then
         txtCodigoCliente.Text = rsCliente("Codigo")
         txtNomeCliente.Text = rsCliente("Nome")
         txtCodigoCliente.SelStart = 0
         txtCodigoCliente.SelLength = Len(txtCodigoCliente.Text)
    End If
    Else
       MsgBox "Cliente 999999 não encontrado, favor cadastrar."
    End If
   rsCliente.Close
   'txtCodigoCliente.SetFocus
End Sub

Private Sub Form_Load()
    Call AjustaTela(Me)
End Sub

Private Sub txtCodigoCliente_GotFocus()
   txtCodigoCliente.SelStart = 0
   txtCodigoCliente.SelLength = Len(txtCodigoCliente.Text)
End Sub

Private Sub txtCodigoCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If txtCodigoCliente.Text = "999999" Then
        
          SQL = "select nfcapa.CONDPAG as condpag, " & _
          "nfcapa.garantiaEstendida as garantiaEstendida from nfcapa " & _
          "where nfcapa.numeroped = " & frmPedido.txtPedido.Text
          rsCliente.CursorLocation = adUseClient
          rsCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
          If rsCliente("condPag") = "01" And rsCliente("garantiaEstendida") <> "S" Then
              cmdGrava_Click
              Call frmPedido.FechaPedido
              Unload Me
          End If
          rsCliente.Close
          
          ElseIf KeyCode = 27 Then
    frmPedido.cmdTotalPedidoGE.Visible = False
    Unload Me
 
End If
  End If
  If KeyCode = vbKeyF2 Then
    cmdGrava_Click
    
        'Call frmPedido.FechaPedido
  
  End If
End Sub

Private Sub txtCodigoCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmPedido.cmdTotalPedidoGE.Visible = False
    Unload Me
End If

If KeyAscii = 13 Then
    If txtCodigoCliente = wCodigoCliente Then
        wCodigoCliente = ""
        cmdGrava_Click
    Else
        Call VerificaCliente
    End If
End If

End Sub


Private Sub VerificaCliente()

 If rsCliente.State = 1 Then
    rsCliente.Close
 End If
 
     SQL = ""
  SQL = "select nfcapa.cliente as Codigo, nfcapa.CONDPAG as condpag, fin_cliente.ce_razao as Nome, nfcapa.pgentra as entrada from fin_cliente, nfcapa " & _
        "where nfcapa.cliente = fin_cliente.ce_codigocliente and nfcapa.numeroped = " & frmPedido.txtPedido.Text
  rsCliente.CursorLocation = adUseClient
  rsCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  

  If IsNumeric(txtCodigoCliente) = False And txtCodigoCliente <> "" Then
         MsgBox "Informe somente números."
         txtCodigoCliente.Text = rsCliente("Codigo")
         txtNomeCliente.Text = rsCliente("Nome")
         txtCodigoCliente.SelStart = 0
         txtCodigoCliente.SelLength = Len(txtCodigoCliente.Text)
         txtCodigoCliente.SetFocus
         rsCliente.Close
         Exit Sub
  
  
  ElseIf txtCodigoCliente = "0" Then
         wClienteTelaAdicionais = True
         rsCliente.Close
         Unload Me
         Unload frmConsCliente
         frmConsCliente.Show 1
         
         Exit Sub

  ElseIf (txtCodigoCliente = "999999" Or txtCodigoCliente = "") And Val(rsCliente("condpag")) > 1 Then
         MsgBox "Não é permitido cliente consumidor para Nota Fiscal Faturada / Financiado"
         txtCodigoCliente.Text = ""
         txtNomeCliente.Text = ""
         txtCodigoCliente.SetFocus
         rsCliente.Close
         Exit Sub
         
  ElseIf txtCodigoCliente = "999999" Or txtCodigoCliente = "0" Or txtCodigoCliente = "900000" Then
        If pedidoComGarantia(frmPedido.txtPedido) Then
            MsgBox "Não é permitido cliente consumidor para Garantia Estendida"
            txtCodigoCliente.Text = ""
            txtNomeCliente.Text = ""
            txtCodigoCliente.SetFocus
            'rsCliente.Close
            Exit Sub
        End If
        rsCliente.Close
         
  ElseIf txtCodigoCliente >= "900000" And Val(rsCliente("condpag")) > 3 Then
         MsgBox "Faturamento não permitido para esse cliente"
         txtCodigoCliente.Text = ""
         txtNomeCliente.Text = ""
         txtCodigoCliente.SetFocus
         rsCliente.Close
         Exit Sub
  ElseIf txtCodigoCliente = "" Then
         If rsCliente("Codigo") > 0 And rsCliente("Codigo") <= 999999 Then
                 txtCodigoCliente.Text = rsCliente("Codigo")
                 txtNomeCliente.Text = rsCliente("Nome")
                 txtCodigoCliente.SelStart = 0
                 txtCodigoCliente.SelLength = Len(txtCodigoCliente.Text)
                 If rsCliente("Entrada") > 0 Then
                    chkEntrada.Value = 1
                    txtEntrada.Text = rsCliente("Entrada")
                 End If
                 
         Else
                 MsgBox "Código do cliente inválido"
                 txtCodigoCliente.Text = ""
                 txtNomeCliente.Text = ""

         End If
         rsCliente.Close
         Exit Sub
  Else
  
         rsCliente.Close
         
         SQL = ""
         SQL = "select ce_CodigoCliente,ce_razao from FIN_Cliente " & _
               "where ce_CodigoCliente = " & txtCodigoCliente
         
         rsCliente.CursorLocation = adUseClient
         rsCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

         If Not rsCliente.EOF Then
                 txtCodigoCliente.Text = rsCliente("ce_CodigoCliente")
                 txtNomeCliente.Text = rsCliente("ce_razao")
                 If chkEntrada.Visible = True Then
                     chkEntrada.SetFocus
                 End If
                 
                 wCodigoCliente = txtCodigoCliente.Text
         Else
                 MsgBox "Código do cliente inválido"
                 txtCodigoCliente.Text = ""
                 txtNomeCliente.Text = ""
                 txtCodigoCliente.SetFocus
         End If
         
         rsCliente.Close

         Exit Sub
  End If
End Sub

Private Sub txtCodigoCliente_LostFocus()
        Call VerificaCliente
End Sub

Private Sub txtTotalGeral_Change()

End Sub

Private Sub txtEntrada_Change()
If IsNumeric(txtEntrada.Text) = False Then
   txtEntrada.Text = ""
   txtEntrada.SelStart = 0
   txtEntrada.SelLength = Len(txtEntrada.Text)

'ElseIf txtEntrada.Text = 0 Then
   'txtQtdeVolume.SetFocus
End If


End Sub

Private Sub txtEntrada_GotFocus()
   txtEntrada.SelStart = 0
   txtEntrada.SelLength = Len(txtEntrada.Text)
End Sub

Private Sub txtEntrada_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      txtQtdeVolume.SetFocus
  End If
End Sub

Private Sub txtEntrada_LostFocus()
txtEntrada.Text = Format(txtEntrada.Text, "###,###,###,##0.00")
 If Val(ConverteVirgula(txtEntrada.Text)) > Val(ConverteVirgula(frmPedido.cmdTotalPedido.Caption)) Then
       txtEntrada.Text = ""
       txtEntrada.SelStart = 0
       txtEntrada.SelLength = Len(txtEntrada.Text)
       txtEntrada.SetFocus
       MsgBox "Valor da entrada maior que o valor da Nota Fiscal"
   Else
      txtQtdeVolume.SetFocus
  End If
End Sub




Private Sub txtPesoVolume_GotFocus()
   txtPesoVolume.SelStart = 0
   txtPesoVolume.SelLength = Len(txtPesoVolume.Text)

End Sub

Private Sub txtQtdeVolume_GotFocus()
   txtQtdeVolume.SelStart = 0
   txtQtdeVolume.SelLength = Len(txtQtdeVolume.Text)
End Sub

Private Sub txtQtdeVolume_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      txtPesoVolume.SetFocus
  End If
End Sub

Public Function pedidoComGarantia(NumeroPedido As String) As Boolean
    Dim rsProdutoGarantiaEstendida As New ADODB.Recordset
    
    pedidoComGarantia = False
    SQL = "select count(*) garantiaEstendida " & _
          "from nfcapa where numeroPed = " & NumeroPedido & " and garantiaEstendida = 'S'"
    
    rsProdutoGarantiaEstendida.CursorLocation = adUseClient
    rsProdutoGarantiaEstendida.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rsProdutoGarantiaEstendida("garantiaEstendida") > 0 Then
        pedidoComGarantia = True
    End If
    rsProdutoGarantiaEstendida.Close
End Function
