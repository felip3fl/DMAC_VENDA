Attribute VB_Name = "Variaveis"
Option Explicit

Global ConexaoDLLAdo As New DMACD.conexaoADO

Global adoCNLoja As New ADODB.Connection
Global rdoCNMatriz As New ADODB.Connection
Global rdoCNAccess As New ADODB.Connection
Global RsDados As New ADODB.Recordset
Global RsDadosItens As New ADODB.Recordset
Global rdoProduto As New ADODB.Recordset
Global RsDadosCapa As New ADODB.Recordset
Global rdoModalidade As New ADODB.Recordset
Global rdoCondPag As New ADODB.Recordset
Global rdoControle As New ADODB.Recordset
Global rsItensNF As New ADODB.Recordset
Global rsControleCaixa As New ADODB.Recordset
Global rsComplementoVenda As New ADODB.Recordset
Global rsCrediario As New ADODB.Recordset
Global rsPesquisaPed As New ADODB.Recordset
Global rsItensVenda As New ADODB.Recordset
Global rsSomaItens As New ADODB.Recordset
Global rdoConPag As New ADODB.Recordset
Global rsPegaNumeroPedido As New ADODB.Recordset
Global rsVendedor As New ADODB.Recordset
Global rsCondicaoFaturado As New ADODB.Recordset
Global rsDadosGerais As New ADODB.Recordset
Global rsCarregaLoja As New ADODB.Recordset
Global rsHabilitaOpcoes As New ADODB.Recordset
Global rsInfLoja As New ADODB.Recordset
Global rsPedido As New ADODB.Recordset
Global rdoCliente As New ADODB.Recordset
Global rdoValidadeCotacao As New ADODB.Recordset
Global rsCondicaoPagamento As New ADODB.Recordset
Global rsControleLoja As New ADODB.Recordset
Global rsPedidosAbertos As New ADODB.Recordset
Global rsCliente As New ADODB.Recordset
Global rsComplemento As New ADODB.Recordset
Global rdoPedido As New ADODB.Recordset
Global rsNumeroCliente As New ADODB.Recordset
Global rdoUf As New ADODB.Recordset
Global rdoCep As New ADODB.Recordset
Global rdoRamo As New ADODB.Recordset
Global rdoDescricao As New ADODB.Recordset
Global rdoSegmento As New ADODB.Recordset
Global rsClientePedido As New ADODB.Recordset
Global RSPegaCliente As New ADODB.Recordset
Global VerificaCliente As New ADODB.Recordset
Global rdoCepEOF As New ADODB.Recordset
Global RsPegaItensEspeciais As New ADODB.Recordset
Global RsICMSInter As New ADODB.Recordset
Global rsLembreMe As New ADODB.Recordset
Global rsLembrete As New ADODB.Recordset
Global rsLembreSeq As New ADODB.Recordset
Global rdoSerie As New ADODB.Recordset
Global rdoConexaoINI As New ADODB.Recordset
Global rdoParametroINI As New ADODB.Recordset
Global adoCNAccess As New ADODB.Connection

Global rmNewUsuario As New ADODB.Recordset
Global rmNew As New ADODB.Recordset

Global wPreencherCliente As Boolean
Global NomeImpressora As Printer
Global wNumeroClientePedido As Double
'Global wPreencherCliente As Boolean

Global GLB_ImpressoraNota As String
Global GLB_ImpCotacao As String
Global sql1 As String
Global GLB_ConectouOK As Boolean
Global GLB_Cotacao As String
Global GLB_VerificaImpressoraFiscal As String
Global GLB_Servidor As String
Global GLB_Banco As String
Global wPesquisaCodigo As Integer
Global CaracterDigitado As Integer
Global GLB_Servidorlocal As String
Global Glb_BancoLocal As String
Global Glb_AlteraResolucao As Boolean

Global wPagamento As String
Global wPagamentototal As String
Global VlTotItem As String
Global wRazao As String
Global Wendereco As String
Global wbairro As String
Global wCGC As String
Global wIest As String
Global WMunicipio As String
Global westado As String
Global WCep As String
Global WFone As String
Global wDDDLoja As String
Global WFax As String
Global wLoja As String
Global GLB_Loja As String
Global wNovaRazao As String
Global wGuardaPagamento As String
Global GLB_ECF As Integer
Global GLB_ImpressoraResumo As String
Global GLB_Impr00 As String

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Ficha financeira
Global GLB_Usuario As String
Global GLB_Senha As String
Global wConexao As String
Global GLB_BancosOnLine As String
Global wCodigoCliFinan As String

Global auxCarimbo As Integer
Global NroBanner As Integer

Global wClienteTelaAdicionais As Boolean
Global GLB_logoPedido As String

Global GBL_Frete As Double
Global wUltimoItem As Integer

Global wConta As Long
Global wContItem As Integer

Global wPagina As Integer
Global wGravaModalidade As Boolean

    Global wIE_Cfo As Double
    Global wCFOItem As Double
    Global GLB_Tributacao As String * 3
    Global wIE_Tributacao  As String * 3
    Global GLB_CFOP As Double
    
    Global wIE_icmsAplicado As Double
    Global wIE_icmsdestino As Double
    Global wIE_BasedeReducao As Double
Global codigoCotacaohtml As String
Global loja() As String
Global Item() As String

Global tabelaHtml As String
Global vendedor As String
Global imagemLogo As String
Global pedidoCotacao As String

Global wValor As Double

Global wLiberaBloqueioPreco As Boolean

Global wCodigoCliente As String
