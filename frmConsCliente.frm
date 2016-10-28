VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form frmConsCliente 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   ClientHeight    =   5580
   ClientLeft      =   1350
   ClientTop       =   2265
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
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
      TabIndex        =   5
      Top             =   4875
      Width           =   6165
   End
   Begin VB.TextBox txtPesquisaCliente 
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
      TabIndex        =   0
      ToolTipText     =   "[Insert] para inserir um novo cliente  "
      Top             =   390
      Width           =   6165
   End
   Begin VSFlex7Ctl.VSFlexGrid grdCliente 
      Height          =   3870
      Left            =   150
      TabIndex        =   1
      Top             =   870
      Width           =   6165
      _cx             =   10874
      _cy             =   6826
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsCliente.frx":0000
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
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.TextBox txtpedido 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   645
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   915
      OleObjectBlob   =   "frmConsCliente.frx":004A
      Top             =   5265
   End
   Begin Project1.chameleonButton cmdImportarContato 
      Height          =   405
      Left            =   4350
      TabIndex        =   4
      Top             =   5055
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Importar"
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
      MICON           =   "frmConsCliente.frx":027E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome / Codigo / CPF / CNPJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   150
      TabIndex        =   3
      Top             =   150
      Width           =   6165
   End
End
Attribute VB_Name = "frmConsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim wCodigo As Integer
Dim wSequencia As Double
Dim wValorDados As String

Private Sub cmdImportar_Click()
    
'    SQL = ""
'    SQL = "Select * from FIN_Cliente Where CE_CodigoCliente = " & Val(txtPesquisaCliente) & ""
'    RSPegaCliente.CursorLocation = adUseClient
'    RSPegaCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
'    'Set RSPegaCliente = Conexao.OpenResultset(SQL)
'
'    On Error Resume Next
'
'    If Not RSPegaCliente.EOF Then
'
'        SQL = ""
'        SQL = "Select * from FIN_ClienteWhere CE_CodigoCliente =  " & Val(txtPesquisaCliente) & ""
'          VerificaCliente.CursorLocation = adUseClient
'          VerificaCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'       ' Set VerificaCliente = adoCNLoja.OpenResultset(SQL)
'        adoCNLoja.BeginTrans
'
'
'        If Not VerificaCliente.EOF Then
'            SQL = "Delete FIN_Cliente where CE_CodigoCliente=" & Val(txtPesquisaCliente) & ""
'                   adoCNLoja.Execute (SQL)
'
'            If Err.Number = 0 Then
'                adoCNLoja.CommitTrans
'            Else
'                adoCNLoja.RollbackTrans
'            End If
'
'            adoCNLoja.BeginTrans
'        End If
'
'         SQL = ""
'         SQL = "Insert into FIN_Cliente([CE_CodigoCliente], [CE_CGC], [CE_InscricaoEstadual], [CE_Razao], " _
'                & "[CE_Endereco], [CE_Bairro], [CE_Municipio], [CE_Estado], [CE_CEP], [CE_Telefone], [CE_Fax], " _
'                & "[CE_EMail], [CE_TipoPessoa], [CE_Praca], [CE_PagamentoCarteira], [CE_EnderecoCobranca], " _
'                & "[CE_BairroCobranca], [CE_MunicipioCobranca], [CE_EstadoCobranca], [CE_CEPCobranca], [CE_LimiteCredito], " _
'                & "[CE_DataLimiteCredito], [CE_MaiorCompra], [CE_DataMaiorCompra], [CE_UltimaCompra], [CE_DataUltimaCompra], " _
'                & "[CE_UltimoPagamento], [CE_DataUltimoPagamento], [CE_MaiorAtraso], [CE_QuantidadeCompras], [CE_JurosCartorio], [CE_DataCadastro], [CE_DataCancelamento], [CE_Alteracao], [CE_Situacao], [CE_HoraManutencao])" _
'                & "Values(" & RSPegaCliente("CE_CodigoCliente") & ", '" & RSPegaCliente("CE_CGC") & "', '" & RSPegaCliente("CE_InscricaoEstadual") & "', '" & RSPegaCliente("CE_Razao") & "', " _
'                & "'" & RSPegaCliente("CE_Endereco") & "', '" & RSPegaCliente("CE_Bairro") & "', '" & RSPegaCliente("CE_Municipio") & "', '" & RSPegaCliente("CE_Estado") & "', '" & RSPegaCliente("CE_CEP") & "', '" & RSPegaCliente("CE_Telefone") & "', '" & RSPegaCliente("CE_Fax") & "', " _
'                & "'" & RSPegaCliente("CE_EMail") & "', '" & RSPegaCliente("CE_TipoPessoa") & "', " & RSPegaCliente("CE_Praca") & ", '" & RSPegaCliente("CE_PagamentoCarteira") & "', '" & RSPegaCliente("CE_EnderecoCobranca") & "', " _
'                & "'" & RSPegaCliente("CE_BairroCobranca") & "', '" & RSPegaCliente("CE_MunicipioCobranca") & "', '" & RSPegaCliente("CE_EstadoCobranca") & "', '" & RSPegaCliente("CE_CEPCobranca") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_LimiteCredito"), "0.00")) & ", " _
'                & "'" & Format(RSPegaCliente("CE_DataLimiteCredito"), "yyyy/mm/dd") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_MaiorCompra"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataMaiorCompra"), "yyyy/mm/dd") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_UltimaCompra"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataUltimaCompra"), "yyyy/mm/dd") & "', " _
'                & "" & ConverteVirgula(Format(RSPegaCliente("CE_UltimoPagamento"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataUltimoPagamento"), "yyyy/mm/dd") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_MaiorAtraso"), "0.00")) & ", " & RSPegaCliente("CE_QuantidadeCompras") & ", " & ConverteVirgula(Format(RSPegaCliente("CE_JurosCartorio"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataCadastro"), "yyyy/mm/dd") & "', '" & Format(RSPegaCliente("CE_DataCancelamento"), "yyyy/mm/dd") & "', '" & RSPegaCliente("CE_Alteracao") & "', '" & RSPegaCliente("CE_Situacao") & "', '" & Format(RSPegaCliente("CE_HoraManutencao"), "yyyy/mm/dd hh:mm") & "')"
'         adoCNLoja.Execute (SQL)
'
'        If Err.Number = 0 Then
'            adoCNLoja.CommitTrans
'            MsgBox "Cliente gravado com sucesso.", vbInformation, "Balcao 2000"
'            txtPesquisaCliente.SelStart = 0
'            txtPesquisaCliente.SelLength = Len(txtPesquisaCliente)
'            txtPesquisaCliente.SetFocus
'        Else
'            adoCNLoja.RollbackTrans
'            MsgBox "Problemas gravação do cliente. Contate o CPD.", vbInformation, "Balcao 2000"
'            txtPesquisaCliente.SelStart = 0
'            txtPesquisaCliente.SelLength = Len(txtPesquisaCliente)
'            txtPesquisaCliente.SetFocus
'        End If
'    Else
'        MsgBox "Código do Cliente não existe.", vbInformation, "Balcao 2000"
'        txtPesquisaCliente.SelStart = 0
'        txtPesquisaCliente.SelLength = Len(txtPesquisaCliente)
'        txtPesquisaCliente.SetFocus
'    End If
'
'    RSPegaCliente.Close
'    VerificaCliente.Close
    
    
End Sub

Private Sub cmdImportar2_Click()

End Sub

Private Sub cmdAtualizar_Click()
    If (grdCliente.Rows - grdCliente.FixedRows) > 1 Then
    
    End If
End Sub

Private Sub chameleonButton1_Click()

End Sub

Private Sub cmdImportarContato_Click()
 
    ConectaODBCMatriz
 
    SQL = "Select CE_CodigoCliente as codigoCliente from FIN_Cliente Where CE_CodigoCliente = " & Val(txtPesquisaCliente) & ""
    RSPegaCliente.CursorLocation = adUseClient
    RSPegaCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic

    If Not RSPegaCliente.EOF Then
        
        adoCNLoja.BeginTrans
        SQL = "exec SP_GLB_Importa_Cliente '" & Val(txtPesquisaCliente) & "'"
        adoCNLoja.Execute (SQL)
        
        If Err.Number = 0 Then
            adoCNLoja.CommitTrans
            If cmdImportarContato.Caption = "Atualizar" Then MsgBox "Cliente atualizado com sucesso!", vbInformation, "DMAC Venda"
            txtPesquisaCliente.SelStart = 0
            txtPesquisaCliente.SelLength = Len(txtPesquisaCliente)
            txtPesquisaCliente_KeyDown 13, 1
        Else
            adoCNLoja.RollbackTrans
            MsgBox "Ocorreu um erro ao tentar realizar a gravação do cliente.", vbCritical, "DMAC Venda"
            txtPesquisaCliente.SelStart = 0
            txtPesquisaCliente.SelLength = Len(txtPesquisaCliente)
            txtPesquisaCliente.SetFocus
        End If
    Else
        MsgBox "Código do Cliente não existe ou nenhuma atualização disponível.", vbInformation, "DMAC Venda"
        txtPesquisaCliente.SelStart = 0
        txtPesquisaCliente.SelLength = Len(txtPesquisaCliente)
        txtPesquisaCliente.SetFocus
    End If
    
    RSPegaCliente.Close
    'VerificaCliente.Close
End Sub

Private Sub cmdRetornar_Click()
 On Error Resume Next
 Unload Me
' frmPedido.picQuadroGeral.Width = 9975
 frmPedido.txtPesquisar.SetFocus

End Sub

Private Sub cmdRetorna_Click()

End Sub

Private Sub Form_Activate()
    'cmdImportarContato.Enabled = False
End Sub

Private Sub Form_Load()
Call AjustaTela(frmConsCliente)

  'Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
 ' Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
  
grdCliente.Rows = 1
'grdCliente.Rows = 2
txtPesquisaCliente.Text = ""
txtpedido.Text = frmPedido.txtpedido
'cmdImportar.Enabled = False
'cmdImportarContato.Enabled = False

End Sub

Private Sub frPesquisa_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grdCliente_Click()
    If grdCliente.Row > 0 And grdCliente.Row <= grdCliente.Rows Then
        cmdImportarContato.Caption = "Atualizar"
    End If
End Sub

Private Sub grdCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wTipoPessoa As Integer

    If KeyCode = 13 Then
        On Error Resume Next
        
        If frmPedido.txtVendedor.Text = "888" Then
            Unload Me
            Unload Me
            Esperar 1
        Else
            wCodigo = 1
            wSequencia = 6
            wValorDados = grdCliente.TextMatrix(grdCliente.Row, 0)

            
            SQL = ""
            SQL = "Select * from FIN_Cliente Where CE_CodigoCliente = " & wValorDados & " "
            
            rsClientePedido.CursorLocation = adUseClient
            rsClientePedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            With rsClientePedido
                wTipoPessoa = 2
                If UCase(.Fields("CE_TipoPessoa")) = "J" Then
                    wTipoPessoa = 1
                End If
            
                SQL = ""
                SQL = "Cliente = " & .Fields("CE_CodigoCliente")
                
                wValorDados = "Cliente = " & .Fields("CE_CodigoCliente")
                rsClientePedido.Close
                
                adoCNLoja.Execute "Exec SP_GravaComplementoVenda " & txtpedido.Text & ",1,1,'" & SQL & "'"

                
                
            End With
            
            inibebotoes (frmPedido.txtpedido)

            If wClienteTelaAdicionais = True Then
                
                frmAdicionais.Show 1
                frmAdicionais.ZOrder
                Unload Me
            Else
               frmPedido.txtPesquisar.SetFocus
               frmPedido.txtPesquisar.SelStart = 0
               frmPedido.txtPesquisar.SelLength = Len(frmPedido.txtPesquisar.Text)
            End If
        End If
    ElseIf KeyCode = 27 Then
        txtPesquisaCliente.SetFocus
    ElseIf KeyCode = 45 Then
        '***************Cadastrar Novo Cliente*********************
        wPreencherCliente = False
        frmCliente.Show 1
        frmCliente.ZOrder
        
    ElseIf KeyCode = vbKeyF1 Then
        '***************Dados Adicionais do cliente***************
           If grdCliente.TextMatrix(grdCliente.Row, 0) <> "" Then
              wNumeroClientePedido = grdCliente.TextMatrix(grdCliente.Row, 0)
              wPreencherCliente = True
              'frmCliente.Top = 960
              'frmCliente.Left = 1470
              frmCliente.ZOrder
              frmCliente.Show 1
           End If
    End If
    
    
End Sub

Private Sub txtPesquisaCliente_GotFocus()
    
    grdCliente.Rows = 1
    'grdCliente.Rows = 2
    txtPesquisaCliente.SetFocus
    txtPesquisaCliente.SelStart = 0
    txtPesquisaCliente.SelLength = Len(txtPesquisaCliente.Text)
    cmdImportarContato.Caption = "Importar"
    
End Sub

Private Sub txtPesquisaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    
   ' Dim rsClientePedido As rdoResultset
    
    If KeyCode = 13 Then
        If Trim(txtPesquisaCliente.Text) <> "" Then
            txtPesquisaCliente.Text = UCase(txtPesquisaCliente.Text)
            grdCliente.Redraw = False
            Screen.MousePointer = 11
            If IsNumeric(txtPesquisaCliente.Text) = True Then
                If Len(txtPesquisaCliente.Text) < 11 Then
                    '************Pesquisa Pelo Numero do Cliente (1)**********************
                    If PesquisaCliente(1, txtPesquisaCliente.Text) = True Then
                        grdCliente.Rows = 1
                        Do While Not rsClientePedido.EOF
                            grdCliente.AddItem rsClientePedido("CE_CodigoCliente") & Chr(9) & rsClientePedido("CE_Razao")
                            rsClientePedido.MoveNext
                        Loop
                        rsClientePedido.Close
                        grdCliente.Row = 1
                        grdCliente.RowSel = 1
                        grdCliente.SetFocus
                    Else
                        'cmdImportar.Enabled = True
                        cmdImportarContato.Caption = "Importar"
                        cmdImportarContato.Enabled = True
                    End If
                ElseIf Len(txtPesquisaCliente.Text) >= 11 Then
                    '***************Pesquisa por CGC ou Cpf (2)***************************
 '                   txtPesquisaCliente.Text = Right(String(15, "0") & txtPesquisaCliente.Text, 15)
                    If PesquisaCliente(2, txtPesquisaCliente.Text) = True Then
                        grdCliente.Rows = 1
                        Do While Not rsClientePedido.EOF
                            grdCliente.AddItem rsClientePedido("CE_CodigoCliente") & Chr(9) & rsClientePedido("CE_Razao")
                            rsClientePedido.MoveNext
                        Loop
                        rsClientePedido.Close
                        grdCliente.Row = 1
                        grdCliente.RowSel = 1
                        grdCliente.SetFocus
                    End If
                Else
                End If
            ElseIf IsNumeric(txtPesquisaCliente.Text) = False Then
                '************Pesquisa Pelo Nome Cliente (3)******************************
                If PesquisaCliente(3, txtPesquisaCliente.Text) = True Then
                    grdCliente.Rows = 1
                    Do While Not rsClientePedido.EOF
                            grdCliente.AddItem rsClientePedido("CE_CodigoCliente") & Chr(9) & rsClientePedido("CE_Razao")
                            rsClientePedido.MoveNext
                    Loop
                    rsClientePedido.Close
                    grdCliente.Row = 1
                    grdCliente.RowSel = 1
                    grdCliente.SetFocus
                    cmdImportarContato.Enabled = True
                End If
            End If
            grdCliente.Redraw = True
            Screen.MousePointer = 0
        Else
            wPreencherCliente = False
            frmCliente.cmbTipoCliente.Locked = False
            frmCliente.Show 1
            frmCliente.ZOrder
            
        End If
    ElseIf KeyCode = 27 Then
        On Error Resume Next
        wClienteTelaAdicionais = False
        Unload Me
    ElseIf KeyCode = 45 Then
        wPreencherCliente = False
        frmCliente.Show 1
        frmCliente.ZOrder
    End If
      
    
End Sub

Function PesquisaCliente(ByVal tipoPesquisa As Integer, ByVal Cliente As String) As Boolean

'
'--------------------------------Pesquisa Pelo Codigo do Cliente (1)-------------------------
'
'    DescricaoOperacao "Pesquisando Cliente"
    If tipoPesquisa = 1 Then
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from FIN_Cliente " _
            & "where CE_CodigoCliente = " & Cliente & " "
            rsClientePedido.CursorLocation = adUseClient
            rsClientePedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = adoCNLoja.OpenResultset(SQL)

'
'-------------------------------Pesquisa por cgc ou cpf (2) ---------------------------------
'
    ElseIf tipoPesquisa = 2 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from FIN_Cliente" _
            & " where CE_Cgc = '" & Cliente & "' "
            rsClientePedido.CursorLocation = adUseClient
            rsClientePedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = adoCNLoja.OpenResultset(SQL)
    
'
'-------------------------------Pesquisa Pelo Nome Cliente (3) ---------------------------------
'
    ElseIf tipoPesquisa = 3 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_razao,CE_CodigoCliente from FIN_Cliente " _
            & "where CE_Razao like '" & UCase(Cliente) & "%' order by CE_Razao"
            rsClientePedido.CursorLocation = adUseClient
            rsClientePedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = adoCNLoja.OpenResultset(SQL)
    
'
'-------------------------------Pesquisa Cliente Tela frmCadCliente(4) --------------------------
'
    ElseIf tipoPesquisa = 4 Then
        SQL = ""
        SQL = ""
        SQL = "Select * from FIN_Cliente " _
            & "where CE_CodigoCliente = " & Cliente & " order by CE_CodigoCliente"
            rsClientePedido.CursorLocation = adUseClient
            rsClientePedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = adoCNLoja.OpenResultset(SQL)
    
    Else
        Exit Function
    End If
    If Not rsClientePedido.EOF Then
        PesquisaCliente = True
    Else
        PesquisaCliente = False
        rsClientePedido.Close
    End If
    DescricaoOperacao "Pronto"
   
End Function

Private Sub txtPesquisaCliente_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
End If
End Sub
