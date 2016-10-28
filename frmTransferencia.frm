VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTransferencia 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transferência de Mercadoria"
   ClientHeight    =   5865
   ClientLeft      =   6375
   ClientTop       =   3480
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3015
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   60
      TabIndex        =   0
      Top             =   5985
      Visible         =   0   'False
      Width           =   300
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   390
      OleObjectBlob   =   "frmTransferencia.frx":0000
      Top             =   5970
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdLojas 
      Height          =   5040
      Left            =   150
      TabIndex        =   1
      Top             =   405
      Width           =   6165
      _cx             =   10874
      _cy             =   8890
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTransferencia.frx":0234
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
      Begin VB.Frame frmTransferenciaTxt 
         BackColor       =   &H00505050&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtSolicitadoPor 
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
            Left            =   2400
            TabIndex        =   8
            Top             =   720
            Width           =   2025
         End
         Begin VB.TextBox txtRetiradoPor 
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
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   2025
         End
         Begin VB.Label lblLojadestino 
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Index           =   2
            Left            =   1320
            TabIndex        =   10
            Top             =   120
            Width           =   3600
         End
         Begin VB.Label lblLojadestino 
            BackStyle       =   0  'Transparent
            Caption         =   "Loja Destino:"
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
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado Por:"
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
            Height          =   225
            Left            =   2400
            TabIndex        =   6
            Top             =   480
            Width           =   1560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Retirado Por:"
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
            Height          =   225
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1200
         End
      End
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5280
      TabIndex        =   3
      Top             =   5055
      Visible         =   0   'False
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
      MICON           =   "frmTransferencia.frx":029A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblLojadestino 
      BackStyle       =   0  'Transparent
      Caption         =   "Loja Destino"
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
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCodigo As Integer
Dim wSequencia As Double
Dim wQtdeItens As Integer
Dim wNumProtocolo As Integer
Dim wValorCampo As String
Dim SQL As String * 500

Dim wTotalNota As Double
Dim wNroCaixa As Integer

Dim wST20 As String
Dim wST60 As String
Dim wST00 As String
Dim wST As String
Dim wSequenciaS As Integer

Dim wCarimboAux As String * 500
Dim wTransporte As String

Dim wQuantItensCapaNF As Integer
Dim wQuantItensNF As Integer
Dim wQuantdadeTotalItem As Integer
Dim wquant As Integer
Dim wDetalheImpressao As String

Dim wRestoItens As Integer
Dim wTotalLinha As Integer
Dim wTotalCarimbo As Integer
Dim wContCarimbo As Integer
Dim wContLinha As Integer
Dim wcfopCapa As Double
Dim wTipoNota As String
Dim wCNPJLoja As String


Private Sub cmdGrava_Click()
  
'  On Error GoTo erronoUpdate
    Dim GLB_AliquotaAplicadaICMS As Double
    Dim GLB_ValorCalculadoICMS As Double
    Dim GLB_TotalIcmsCalculado As Double
    Dim GLB_BasedeCalculoICMS As Double
    Dim GLB_BaseTotalICMS As Double
    
    Dim rsCarimbo As New ADODB.Recordset
    
    Dim NumeroDocumento As String
    Dim wChaveICMSItem As Double
    Dim RsCapaNF As New ADODB.Recordset
    Dim wPessoa As String
    Dim wChaveICMS As Double
    Dim wSubstituicaoTributaria As String
    Dim EncerraVenda As Boolean

    Dim wAnexoIten As Double
  
    SQL = "update nfitens set vlunit = Round(pr_customedio1,2), vltotitem = (Round(pr_customedio1,2) * qtde) from nfitens, produtoloja where pr_referencia = referencia and numeroped = " & frmPedido.txtpedido.Text
    adoCNLoja.Execute SQL
  
    SQL = "Select LO_TipoTransferencia from Loja where lo_loja = '" & Trim(grdLojas.TextMatrix(grdLojas.Row, 0)) & "'"
    rdoControle.CursorLocation = adUseClient
    rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rdoControle.EOF Then
       MsgBox "Problemas com o Sistema (Controle Sistema) entrar em contato com o CPD", vbCritical, "Atenção"
       rdoControle.Close
       Exit Sub
   End If
   
   wSequencia = 0
   If Trim(rdoControle("LO_TipoTransferencia")) = "T" Then
       'wValorCampo = PegaSerieNota
       wTipoNota = "TA"
        If Trim(grdLojas.TextMatrix(grdLojas.Row, 2)) <> wCNPJLoja Then
         wValorCampo = PegaSerieNota
        Else
          wValorCampo = "CT"
        End If
       
   Else
      wValorCampo = "CT"
      wTipoNota = "TA"
      wSequencia = Trim(frmPedido.txtpedido.Text)
   End If
   rdoControle.Close
    
    wQtdeItens = 0
    wNumProtocolo = 0
    wST20 = "N"
    wST60 = "N"
    wST00 = "N"
    wST = "00"
    wSequenciaS = 0
    wTransporte = Trim(txtRetiradoPor.Text) & "/" & Trim(txtSolicitadoPor.Text)
    
    
    'SQL = ""
    'SQL = "Select Count(*) as Itens From NFItens Where NumeroPed = " & frmPedido.txtPedido.Text
    'rdoControle.CursorLocation = adUseClient
    'rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    'If rdoControle.EOF = True Then
       'MsgBox "Não foram encontrados itens para este pedido.", vbCritical, "Atenção"
       'Exit Sub
    'Else
       'wQtdeItens = rdoControle("Itens")
    'End If
    'rdoControle.Close
 
    'SQL = ""
    'SQL = "Select CTR_Protocolo,CTR_NumeroCaixa From ControleCaixa Where CTR_SituacaoCaixa = 'A'"
    'rdoControle.CursorLocation = adUseClient
    'rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    'If rdoControle.EOF = True Then
       'MsgBox "Problemas com o Sistema (Controle Caixa) entrar em contato com o TI", vbCritical, "Atenção"
       'Exit Sub
    'Else
       
       'wNroCaixa = rdoControle("CTR_NumeroCaixa")
       'wNumProtocolo = rdoControle("CTR_Protocolo")
    'End If
    'rdoControle.Close
 

        SQL = "Select nfcapa.*, fin_Estado.*,fin_Cliente.* from nfcapa, fin_Estado, fin_cliente where nfcapa.numeroped = " & _
            frmPedido.txtpedido.Text & " and nfcapa.cliente = fin_cliente.ce_codigocliente " & _
            "And fin_cliente.ce_estado = fin_Estado.UF_Estado"
             
             RsCapaNF.CursorLocation = adUseClient
             RsCapaNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
             
            If Not RsCapaNF.EOF Then
            'If RsCapaNF("ce_Tipopessoa") = "F" Or RsCapaNF("ce_Tipopessoa") = "U" Then
              'wPessoa = 2
            'ElseIf RsCapaNF("ce_Tipopessoa") = "O" Then
              'wPessoa = 3
            'Else
              wPessoa = 1
            'End If
                wChaveICMS = RsCapaNF("UF_Regiao") & wPessoa
            End If

    If wValorCampo = "0" Or wValorCampo = "" Then
       MsgBox "Não foi possível criar a nota fiscal.", vbCritical, "Aviso"
       Exit Sub
    Else
     
    Screen.MousePointer = vbHourglass
    
     SQL = "Select produtoloja.*, nfitens.* from produtoloja,nfitens " _
              & "where nfitens.numeroped = " & frmPedido.txtpedido.Text & "" _
              & " and pr_referencia = nfitens.referencia order by NfItens.Item"
          
              rsItensNF.CursorLocation = adUseClient
              rsItensNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
              
          If Not rsItensNF.EOF Then
             Do While Not rsItensNF.EOF
                    wChaveICMSItem = wChaveICMS
                    If rsItensNF("PR_substituicaotributaria") = "N" _
                       And rsItensNF("PR_codigoreducaoicms") > 0 Then
                        wST20 = "S"
                        wST = "20"
                    End If
                                 
                    If rsItensNF("PR_substituicaotributaria") = "S" Then
                        wSubstituicaoTributaria = 1
                        wST60 = "S"
                        wChaveICMSItem = wChaveICMSItem & "000" & wSubstituicaoTributaria
                    Else
                        wSubstituicaoTributaria = 0
                        wChaveICMSItem = wChaveICMSItem & Format(rsItensNF("pr_icmssaida"), "####00") & rsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
                    End If
                    
                    If AcharICMSInterEstadual(rsItensNF("PR_Referencia"), wChaveICMSItem) = False Then
                          
                          If AcharICMSInterEstadual(rsItensNF("PR_Referencia"), Mid(Trim(wChaveICMSItem), 1, 2) & "1200") = False Then
                                EncerraVenda = False
                                rsItensNF.Close
                                Exit Sub
                          End If
                    End If
                    
                    'If rsItensNF("PR_substituicaotributaria") = "S" Then
                        'wSubstituicaoTributaria = 1
                        'wST60 = "S"
                        'wChaveICMSItem = wChaveICMSItem & "000" & wSubstituicaoTributaria
                    'Else
                        'wSubstituicaoTributaria = 0
                        'wChaveICMSItem = wChaveICMSItem & Format(rsItensNF("pr_icmssaida"), "####00") & rsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
                    'End If
                                     
                            wCFOItem = wIE_Cfo
                            GLB_AliquotaAplicadaICMS = wIE_icmsAplicado
                            GLB_Tributacao = wIE_Tributacao
                            GLB_CFOP = wIE_Cfo
                            wAnexoIten = rsItensNF("PR_CodigoReducaoICMS")
                                     
                            GLB_ValorCalculadoICMS = Format((((rsItensNF("vltotitem") - rsItensNF("desconto")) * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                            GLB_TotalIcmsCalculado = (GLB_TotalIcmsCalculado + GLB_ValorCalculadoICMS)
                            If GLB_TotalIcmsCalculado > 0 Then
                                If wIE_BasedeReducao = 0 Then
                                    If GLB_AliquotaAplicadaICMS = 0 Then
                                        GLB_BasedeCalculoICMS = 0
                                    Else
                                        GLB_BasedeCalculoICMS = (rsItensNF("vltotitem") - rsItensNF("desconto"))
                                    End If
                                Else
                                    GLB_BasedeCalculoICMS = Format((rsItensNF("vltotitem") - rsItensNF("desconto")) - _
                                    (((rsItensNF("vltotitem") - rsItensNF("desconto")) * wIE_BasedeReducao) / 100), "0.00")
                                End If
                                GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
                            End If
                                     
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
'aqui
                                     
                    SQL = "UPDATE nfitens set baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
                    & "Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & " ,TipoNota = 'T'," _
                    & "DetalheImpressao = '" & "T" & "', CSTICMS = " & GLB_Tributacao & ", " _
                    & "CFOP = " & GLB_CFOP & ", ICMSAplicado = " & ConverteVirgula(wIE_icmsdestino) & " " _
                    & "" _
                    & " where nfitens.numeroped = " & frmPedido.txtpedido.Text _
                    & " and Referencia = '" & rsItensNF("PR_Referencia") & "' and Item=" & rsItensNF("Item") & ""
                    adoCNLoja.Execute (SQL)
                                     
                    'If Err.Number = 0 Then
                        'adoCNLoja.CommitTrans
                    'Else
                        'adoCNLoja.RollbackTrans
                    'End If
                                     
                                     
                    'SQL = "Update Nfitens set CSTICMS = '" & wST & "' Where Referencia = '" & rsItensNF("Referencia") _
                        & "' and NumeroPed = " & frmPedido.txtPedido.Text
                    'adoCNLoja.Execute SQL
                    
                    rsItensNF.MoveNext
                Loop
            End If
            rsItensNF.Close
                
                
            If wST60 = "S" And wST20 = "N" And wST00 = "N" Then
                wcfopCapa = 5409
            Else
                wcfopCapa = 5152
            End If


'******************* ATUALIZA NFCAPA

        SQL = "update nfitens set BaseICMS = 0 where BaseICMS is null and numeroped = " & frmPedido.txtpedido.Text
        adoCNLoja.Execute (SQL)
             
       SQL = ""
       SQL = "Select * From Loja Where LO_Loja = '" & Trim(grdLojas.TextMatrix(grdLojas.Row, 0)) & "'"
       rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
            SQL = "Select sum(BASEICMS) as BaseICMS from nfitens where numeroped = " & frmPedido.txtpedido.Text
             rsItensNF.CursorLocation = adUseClient
             rsItensNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
       
       ''SQL = ""
'       SQL = "update nfcapa set TipoNota = '" & Trim(wTipoNota) & "', VendedorLojaVenda = 999,LojaVenda = LojaOrigem, OutraLoja = LojaOrigem," _
'             & "OutroVend = 999, Vendedor = 999, NF = " & wSequencia & ",desconto = 0,fretecobr = 0, Serie = '" _
'             & wValorCampo & "', Dataemi = '" & Format(Date, "yyyy/mm/dd") & "', nroCaixa = " & wNroCaixa _
'             & ",CodOper = " & wcfopCapa & ", CondPag = '01'" _
'             & ", cliente = 0 " _
'             & ", Protocolo = " & wNumProtocolo & ",QtdItem = " & wQtdeItens & ",TipoTransporte = '" & wTransporte & "'," _
'             & " baseicms = " & ConverteVirgula(rsItensNF("BaseICMS")) _
'             & ",garantiaEstendida = DEFAULT, totalGarantia = DEFAULT" _
'             & " Where NumeroPed = " & frmPedido.txtPedido.Text
             
       SQL = "update nfcapa set TipoNota = '" & Trim(wTipoNota) & "', VendedorLojaVenda = 999,LojaVenda = LojaOrigem, OutraLoja = LojaOrigem," _
             & "OutroVend = 999, Vendedor = 999, NF = " & wSequencia & ",desconto = 0,fretecobr = 0, Serie = '" _
             & wValorCampo & "', Dataemi = '" & Format(Date, "yyyy/mm/dd") & "', nroCaixa = " & wNroCaixa _
             & ",CodOper = " & wcfopCapa & ", CFOAUX = " & wcfopCapa & ", CondPag = '01'" _
             & ", lojat = '" & Trim(grdLojas.TextMatrix(grdLojas.Row, 0)) & "', cliente = 0 " _
             & ", Protocolo = " & wNumProtocolo & ",QtdItem = " & wQtdeItens & ",TipoTransporte = '" & wTransporte & "'," _
             & " baseicms = " & ConverteVirgula(rsItensNF("BaseICMS")) _
             & ",garantiaEstendida = DEFAULT, totalGarantia = DEFAULT" _
             & " Where NumeroPed = " & frmPedido.txtpedido.Text
'''   FOI RETIRADO LOJAT E CFOAUX
''' FELIPE 2014
             
       adoCNLoja.Execute (SQL)
       rdoControle.Close
       Screen.MousePointer = vbHourglass
       
       rsItensNF.Close
       
       
'------Nova--------------------------------------------------------------------------------------


       wQuantItensCapaNF = 0
       wQuantItensNF = 0
       wQuantdadeTotalItem = 0
       wquant = 0
       
       SQL = "select qtditem from nfcapa Where NumeroPed = " & frmPedido.txtpedido.Text
       rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
           
       SQL = "select numeroped,item from nfitens Where NumeroPed = " & frmPedido.txtpedido.Text _
           & " order by item"
       rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

       If Not rdoControle.EOF Then
          Do While Not rdoControle.EOF
                
              wQuantItensCapaNF = rsComplementoVenda("QtdItem")
              wQuantItensNF = wQuantItensNF + 1
              wQuantdadeTotalItem = wQuantdadeTotalItem + 1
              wquant = (wQuantItensNF Mod 12)
                      
               If wquant <> 0 Then

                   If wQuantItensCapaNF = wQuantItensNF Then
                        If wquant = 11 Then
                             wDetalheImpressao = "C"
                        Else
                             wDetalheImpressao = "T"
                        End If
                   ElseIf wQuantItensCapaNF = wQuantdadeTotalItem Then
                        If wquant = 11 Then
                             wDetalheImpressao = "C"
                        Else
                             wDetalheImpressao = "T"
                        End If
                  Else
                        wDetalheImpressao = "D"
                  End If

               Else

                  wDetalheImpressao = "C"
                  wUltimoItem = wUltimoItem + 1
               End If
               
               SQL = "UPDATE NFItens Set VLUNIT = Round(PR_CustoMedio1,2), VLTOTITEM = (Round(PR_CustoMedio1,2) * QTDE)," _
                   & "TipoNota = '" & wTipoNota & "',DataEmi = '" & Format(Date, "yyyy/mm/dd") & "',NF = " & wSequencia & ", Serie = '" & wValorCampo & "'," _
                   & "DescricaoAlternativa = '', ReferenciaAlternativa = ''," _
                   & "detalheimpressao = '" & Trim(wDetalheImpressao) & "', desconto = 0, " _
                   & "" _
                   & "garantiaEstendida = DEFAULT, planogarantia = DEFAULT, certificadoInicio = DEFAULT, certificadoFim = DEFAULT " _
                   & "From NFItens, ProdutoLoja " _
                   & "Where PR_Referencia = Referencia And NumeroPed = " & frmPedido.txtpedido.Text _
                   & " and item = " & rdoControle("Item")
                   
''               SQL = "UPDATE NFItens Set VLUNIT = Round(PR_CustoMedio1,2), VLTOTITEM = (Round(PR_CustoMedio1,2) * QTDE)," _
''                   & "TipoNota = '" & wTipoNota & "',DataEmi = '" & Format(Date, "yyyy/mm/dd") & "',NF = " & wSequencia & ", Serie = '" & wValorCampo & "'," _
''                   & "DescricaoAlternativa = '', ReferenciaAlternativa = ''," _
''                   & "detalheimpressao = '" & Trim(wDetalheImpressao) & "', desconto = 0, " _
''                   & "" _
''                   & "garantiaEstendida = DEFAULT, planogarantia = DEFAULT, certificadoInicio = DEFAULT, certificadoFim = DEFAULT " _
''                   & "From NFItens, ProdutoLoja " _
''                   & "Where PR_Referencia = Referencia And NumeroPed = " & frmPedido.txtPedido.Text _
''                   & " and item = " & rdoControle("Item")
                   
               adoCNLoja.Execute (SQL)
               
               rdoControle.MoveNext
        Loop
     End If
     rdoControle.Close
     rsComplementoVenda.Close
   '--------------------------------------------------------------------------------------------------------
   
'****************** CARIMBO

   SQL = "Select max(cnf_Sequencia) as Sequencia From CarimboNotaFiscal Where CNF_NumeroPed = " & frmPedido.txtpedido.Text
   rsCarimbo.CursorLocation = adUseClient
   rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   If IsNull(rsCarimbo("Sequencia")) Then
      auxCarimbo = 2
   Else
      auxCarimbo = rsCarimbo("Sequencia") + 1
   End If
   rsCarimbo.Close
   
   
   'SQL = ""
   'SQL = "Select LojaOrigem,Serie,nf,numeroped From NFCapa Where NumeroPed = " & txtPedido.Text & " and " _
         '& "TipoNota = 'PD'"
   'rsCarimbo.CursorLocation = adUseClient
   'rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    SQL = ""
    SQL = "Insert into CarimboNotaFiscal(CNF_NumeroPed,CNF_Loja,CNF_serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo)" _
        & "Values ( " & frmPedido.txtpedido.Text & ",'" & wLoja & _
        "','',0," & auxCarimbo & ",'" & "TIPO TRANSPORTE: " & wTransporte & "','I')"
        
    adoCNLoja.Execute SQL
        
     'rsCarimbo.CursorLocation = adUseClient
     'rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    'rsCarimbo.Close

'****************** DETALHE IMPRESSAO
       Call encerraTransferencia(wSequencia, frmPedido.txtpedido.Text)
       SQL = "Exec SP_Totaliza_Capa_Nota_Fiscal_Loja " & frmPedido.txtpedido.Text
       adoCNLoja.Execute SQL
       
     End If
       
       ''Call encerraTransferencia2(Val(txtpedido.Text), "NE")
     
   Unload Me
   Call LimpaForm

'  Exit Sub

'erronoUpdate:
'MsgBox "Erro na atualização da situação do pedido " & Err.description, vbCritical, "Aviso"
Screen.MousePointer = vbNormal
  
End Sub

Private Sub cmdRetorna_Click()
 Unload Me

 frmPedido.picQuadroGeral.Width = 11550
 frmPedido.txtPesquisar.SetFocus

End Sub

Private Sub cmdGrava_LostFocus()
grdLojas.Enabled = False

End Sub

Private Sub Form_Activate()
    cmdGrava.Enabled = False
End Sub

Private Sub Form_Load()
  
  Call AjustaTela(frmTransferencia)
  
    
'  Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
  'txtTransporte.Enabled = False
  wTotalLinha = 0
  wTotalCarimbo = 0
  wRestoItens = 0
  wContLinha = 0
  wUltimoItem = 0
  wContCarimbo = 0
  txtpedido.Text = frmPedido.txtpedido.Text
  
  wCodigo = 1
  SQL = "Select LO_Loja, LO_Endereco, lo_cgc from Loja where " _
      & "LO_OrdemLoja <> 888 and  lo_situacao = 'A' and  LO_Loja not in('CONSO','CMCS','CMCE') and lo_loja <> '" & wLoja & "' Order By lo_regiao"
  
  rsCarregaLoja.CursorLocation = adUseClient
  rsCarregaLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  If Not rsCarregaLoja.EOF Then
     Do While Not rsCarregaLoja.EOF
        grdLojas.AddItem rsCarregaLoja("LO_Loja") & Chr(9) _
        & rsCarregaLoja("LO_Endereco") & Chr(9) & rsCarregaLoja("lo_cgc")
        rsCarregaLoja.MoveNext
     Loop
  End If
  
  rsCarregaLoja.Close
  
  SQL = "Select top 1 lo_cgc from Loja where " _
      & "LO_Loja = '" & wLoja & "'"
  
  rsCarregaLoja.CursorLocation = adUseClient
  rsCarregaLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  If Not rsCarregaLoja.EOF Then
    wCNPJLoja = rsCarregaLoja("lo_cgc")
  End If
  
  rsCarregaLoja.Close
  
End Sub

Private Sub grdLojas_Click()
'    fraTipoTransp.Enabled = True
    frmTransferenciaTxt.Visible = True
    txtRetiradoPor.SetFocus
    lblLojadestino(2).Caption = Trim(grdLojas.TextMatrix(grdLojas.Row, 0)) & " - " & Trim(grdLojas.TextMatrix(grdLojas.Row, 1))
'    cmdGrava.Enabled = True
End Sub

Private Sub grdLojas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
  cmdGrava_Click
 End If
End Sub

Private Sub grdLojas_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If

End Sub

Private Sub txtTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
  cmdGrava_Click
 End If
End Sub

'    Private Sub grdLojas_LostFocus()
'    If txtTransporte <> "" Then
'        cmdGrava.Enabled = True
'        cmdGrava.SetFocus
'    End If
'
'End Sub

Private Sub txtTransporte_KeyPress(KeyAscii As Integer)
 
If KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If

If KeyAscii = 13 Then
    'If txtTransporte <> "" Then
        'cmdGrava_Click
    'Else
        'txtTransporte.SetFocus
    'Exit Sub
    'End If
End If
    

End Sub

Private Sub encerraTransferencia(NumeroDocumento As Double, NumeroPedido As Double)
Dim rsItensNF As New ADODB.Recordset
Dim RsCapaNF As New ADODB.Recordset
Dim rsCarimbo As New ADODB.Recordset

Dim wTM As Integer
Dim wPessoa As Integer
Dim wChaveICMS As Integer
Dim wSubstituicaoTributaria As Integer
Dim wQuantidadeItensCapaNF As Integer
Dim wQuantidadeItensNF As Integer
Dim wQuantidadeTotalItem As Integer
Dim wQuantidade As Integer
Dim wUltimoItem As Integer
'Dim wQuantdadeTotalItem As Integer
Dim wQuantItensNF As Integer
Dim wAnexoIten As Integer
Dim wUFRegiao As Integer

Dim GLB_AliquotaAplicadaICMS As Double
Dim GLB_AliquotaICMS As Double
Dim GLB_Tributacao As Double
Dim GLB_CFOP As Double
Dim GLB_ValorCalculadoICMS As Double
Dim GLB_TotalIcmsCalculado As Double
Dim GLB_BasedeCalculoICMS As Double
Dim GLB_BaseTotalICMS As Double
Dim wAnexoItem As Double

Dim wAnexo1 As String
Dim wAnexo2 As String
Dim wAnexoAux As String
Dim wDetalheImpressao As String
Dim SerieProd As String
Dim wCarimbo2 As String
Dim wCarimbo5 As String
Dim wChaveICMSItem As String
Dim wPegaCarimboNF As String
Dim wRecebeCarimboAnexo As String
Dim wCFO1 As String
Dim wCFO2 As String
Dim wCFOItem As String


' On Error GoTo ErroEncerraTransferencia
wCFO1 = ""
wCFO2 = ""

'
'  --------------------------------- CALCULO DO ICMS ------------------------------------------------------------------------
'
             
''        SQL = "Select top 1 nfcapa.*, fin_Estado.*, Loja.* from nfcapa, fin_Estado, Loja " _
''              & "Where nfcapa.numeroped = " & NumeroPedido & "" _
''              & " and nfcapa.nf = " & NumeroDocumento & " And convert(Char(7), nfcapa.Cliente) = convert(char(7),Loja.lo_loja)" _
''              & " And LOJA.lo_UF = fin_Estado.UF_Estado"
''
''        RsCapaNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''
''        If RsCapaNF.EOF Then
''           MsgBox "Nota não encontrada.", vbInformation, "AVISO"
''           Exit Sub
''        End If
''
''       wTM = RsCapaNF("TM")
''       wPessoa = "1"
''       wChaveICMS = RsCapaNF("UF_Regiao") & wPessoa
''       'wTipoNota = UCase(Trim(rsCapaNF("Tiponota")))
''       wQuantidadeItensCapaNF = RsCapaNF("QtdItem")
''       wUFRegiao = RsCapaNF("UF_Regiao")
''
''
''
''        SQL = "Select ProdutoLoja.*, NFItens.* from ProdutoLoja, NFItens " _
''              & "Where NFItens.NumeroPed = " & NumeroPedido & " " _
''              & "and NFItens.nf = " & NumeroDocumento & " and PR_Referencia = NFItens.Referencia Order By NfItens.Item"
''
''        rsItensNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''        If Not rsItensNF.EOF Then
''           Do While Not rsItensNF.EOF
''             If wTM <> 1 Then
''                wChaveICMSItem = wChaveICMS
''                If rsItensNF("PR_IcmsSaida") = 0 And rsItensNF("PR_SubstituicaoTributaria") = "N" Then
''                   wST20 = "S"
''                End If
''
''                If rsItensNF("PR_SubstituicaoTributaria") = "S" Then
''                   wSubstituicaoTributaria = 1
''                   wST60 = "S"
''                   wChaveICMSItem = wChaveICMSItem & "000" & wSubstituicaoTributaria
''                Else
''                   wSubstituicaoTributaria = 0
''                   wChaveICMSItem = wChaveICMSItem & Format(rsItensNF("pr_icmssaida"), "####00") & rsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
''                End If
''
''                If AcharICMSInterEstadual(rsItensNF("PR_Referencia"), wChaveICMSItem) = False Then
''                   Exit Sub
''                Else
''
''                   GLB_AliquotaAplicadaICMS = RsICMSInter("IE_icmsAplicado")
''                   GLB_AliquotaICMS = RsICMSInter("IE_IcmsDestino")
''                   GLB_CFOP = RsICMSInter("IE_CFOP")
''                   wCFOItem = RsICMSInter("IE_CfOP")
''                End If
''
''                   GLB_ValorCalculadoICMS = Format((((rsItensNF("vltotitem") - rsItensNF("Desconto")) _
''                                            * GLB_AliquotaAplicadaICMS) / 100), "0.00")
''                   GLB_TotalIcmsCalculado = (GLB_TotalIcmsCalculado + GLB_ValorCalculadoICMS)
''
''                   If GLB_TotalIcmsCalculado > 0 Then
''                     If RsICMSInter("IE_BasedeReducao") = 0 Then
''                       If GLB_AliquotaAplicadaICMS = 0 Then
''                          GLB_BasedeCalculoICMS = 0
''                       Else
''                          GLB_BasedeCalculoICMS = (rsItensNF("vltotitem") - rsItensNF("Desconto"))
''                       End If
''                     Else
''                       GLB_BasedeCalculoICMS = Format((rsItensNF("vltotitem") - rsItensNF("Desconto")) - _
''                                               (((rsItensNF("vltotitem") - rsItensNF("Desconto")) * _
''                                               RsICMSInter("IE_BasedeReducao")) / 100), "0.00")
''                     End If
''                       GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
''                   End If
''
''
''                   If wTipoNota = "TA" Then
''                      If rsItensNF("PR_substituicaotributaria") = "S" Then
''                         wCFO2 = 5409
''                      Else
''                         wCFO1 = 5152 & " "
''                      End If
''                   End If
''             End If
''
''             RsICMSInter.Close
''             rsItensNF.MoveNext
''           Loop
''           rsItensNF.Close
''        Else
''           MsgBox "Nota não encontrada.", vbInformation, "AVISO"
''           Exit Sub
''
''        End If
        
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
        
        
        
'
' -------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
'
     If wTM <> 1 Then
'
'             RsCapaNF.Close
'
'             SQL = "UPDATE NFCapa Set BaseIcms = " & ConverteVirgula(Format(GLB_BaseTotalICMS, "###,###,##0.00")) & ", " _
'                   & "VlrIcms = " & ConverteVirgula(GLB_TotalIcmsCalculado) & ", " _
'                   & "Paginanf = " & ConverteVirgula(wUltimoItem) & ", TipoTransporte = '" & wTransporte & "' " _
'                   & "Where NumeroPed = " & NumeroPedido & " and NF = " & NumeroDocumento & ""
'              adoCNLoja.Execute (SQL)
'
' -------------------------------------- INSERIR CARIMBOS --------------------------------------------------


            SQL = ""
            SQL = "update CarimboNotafiscal set CNF_Serie = '" & PegaSerieNota & "', CNF_NF = " & NumeroDocumento & _
                  " where CNF_NumeroPed = " & NumeroPedido
            adoCNLoja.Execute (SQL)

            If wST20 = "S" Then
            
              wSequenciaS = wSequenciaS + 1
              
               
            SQL = ""
            SQL = "Select CE_linha12 from CarimbosEspeciais where ce_Referencia = '9999991'"
            rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
                  
              SQL = ""
              SQL = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo) " & _
                    " Values(" & NumeroPedido & ",'" & Trim(wLoja) & "','" & PegaSerieNota & "'," & NumeroDocumento & _
                    "," & wSequenciaS & " ,'" & rsCarimbo("CE_linha12") & "' , 'S')"
              adoCNLoja.Execute (SQL)
              rsCarimbo.Close
                
            End If
            
            If wST60 = "S" Then
            
              wSequenciaS = wSequenciaS + 1
              
               
            SQL = ""
            SQL = "Select CE_linha12 from CarimbosEspeciais where ce_Referencia = '9999992' "
            rsCarimbo.CursorLocation = adUseClient
            rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
              SQL = ""
              SQL = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo) " & _
                    " Values(" & NumeroPedido & ",'" & Trim(wLoja) & "','" & PegaSerieNota & "'," & NumeroDocumento & _
                    "," & wSequenciaS & " ,'" & rsCarimbo("CE_linha12") & "' , 'S')"
              adoCNLoja.Execute (SQL)
              rsCarimbo.Close
                
            End If

            SQL = ""
            SQL = "select count(*) as somacarimbo from carimbonotafiscal where " & _
                         " CNF_NumeroPed = " & NumeroPedido & " "
            rsCarimbo.CursorLocation = adUseClient
            rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            If Not rsCarimbo.EOF Then
               wTotalCarimbo = rsCarimbo("somacarimbo")
            End If
            rsCarimbo.Close
            
            SQL = ""
            SQL = "select count(*) as somaitens from nfitens where " & _
                         " NumeroPed = " & NumeroPedido & " "
            rsCarimbo.CursorLocation = adUseClient
            rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            If Not rsCarimbo.EOF Then
               wRestoItens = ((rsCarimbo("somaitens")) Mod 12)
            End If
            rsCarimbo.Close
            
            SQL = ""
            SQL = "select * from carimbonotafiscal where CNF_NumeroPed = " & NumeroPedido & _
                  " order by cnf_tipocarimbo desc, cnf_sequencia asc"
            rsCarimbo.CursorLocation = adUseClient
            rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

            If Not rsCarimbo.EOF Then
                
                wTotalLinha = ((wTotalCarimbo + wRestoItens) + 1)
                wContCarimbo = 0
                  
                If wRestoItens <> 0 Then
                   wContLinha = (wRestoItens + 1)
                End If
                                            
                Do While Not rsCarimbo.EOF
                   wContLinha = wContLinha + 1
                   wContCarimbo = wContCarimbo + 1
                   If (wContLinha Mod 12) <> 0 Then
                       wDetalheImpressao = "D"
                   Else
                       wDetalheImpressao = "C"
                       wUltimoItem = wUltimoItem + 1
                   End If
                   
                   If wTotalLinha = wContLinha Then
                       wDetalheImpressao = "T"
                   End If
                   
                   If wTotalCarimbo = wContCarimbo Then
                      wDetalheImpressao = "T"
                   End If

                   adoCNLoja.BeginTrans
                   
                   SQL = ""
                   SQL = "update CarimboNotafiscal set CNF_DetalheImpressao = '" & wDetalheImpressao & "', cnf_data = '" & Format(Date, "yyyy/mm/dd") & "'" & _
                         " where cnf_Loja = '" & Trim(rsCarimbo("cnf_Loja")) & _
                         "' and cnf_tipocarimbo = '" & rsCarimbo("cnf_tipocarimbo") & "' " & _
                         " and cnf_sequencia = '" & rsCarimbo("cnf_sequencia") & "' " & _
                         " and CNF_NumeroPed = " & NumeroPedido & " "
                         adoCNLoja.Execute (SQL)
                         
                   If Err.Number = 0 Then
                        adoCNLoja.CommitTrans
                   Else
                        adoCNLoja.RollbackTrans
                   End If
                   
                   rsCarimbo.MoveNext
                Loop
             End If
             rsCarimbo.Close
        '------------------------------------------
          Else
             MsgBox "Não foi possível acessar os carimbos fiscais", vbCritical, "AVISO"
             Exit Sub

          End If
    
   Exit Sub
End Sub

Function CriaMovimentoCaixa(ByVal Nf As Double, ByVal Serie As String, ByVal TotalNota As Double, ByVal loja As String, ByVal Grupo As Double, ByVal NroProtocolo As Integer, ByVal NroCaixa As Integer, ByVal NroPedido As Double)
    
    SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
        & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo, MC_NroCaixa, MC_DataProcesso, MC_Pedido) values(" & GLB_ECF & ",'0','" & Trim(loja) & "', " _
        & " '" & Format(Date, "yyyy/mm/dd") & "'," & Grupo & ", " & Nf & ",'" & Serie & "', " _
        & "" & ConverteVirgula(Format(TotalNota, "##,###0.00")) & ", " _
        & "0,0,0,0,0,9,'A'," & NroProtocolo & "," & NroCaixa & ",'" & Format(Date, "yyyy/mm/dd") & "'," & NroPedido & ")"
        adoCNLoja.Execute (SQL)

End Function




Private Sub txtRetiradoPor_KeyPress(KeyAscii As Integer)
If (KeyAscii > 33) And (KeyAscii < 65) Then
        KeyAscii = 0
    ElseIf KeyAscii = 27 Then
        frmTransferenciaTxt.Visible = False
        grdLojas.Enabled = True
          ElseIf KeyAscii = 13 Then
            If txtRetiradoPor.Text <> "" Then
                txtSolicitadoPor.SetFocus
            End If

    Else
    KeyAscii = KeyAscii
End If


End Sub

Private Sub txtSolicitadoPor_KeyPress(KeyAscii As Integer)
If (KeyAscii > 33) And (KeyAscii < 65) Then
        KeyAscii = 0
        
    ElseIf KeyAscii = 27 Then
        frmTransferenciaTxt.Visible = False
        grdLojas.Enabled = True
        
        
    ElseIf KeyAscii = 13 Then
            If txtSolicitadoPor.Text <> "" Then
                cmdGrava_Click
            End If

End If


End Sub


