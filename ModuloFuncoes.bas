Attribute VB_Name = "ModuloFuncoes"
                                                                                                                                                                                            
                                                                                                                                                                                            '********************************************************************************************
'
'       Inicío das Declarações das Funções para o uso da Impressora Fiscal
'
'********************************************************************************************

' Funções de Inicialização
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BEMAFI32.DLL" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BEMAFI32.DLL" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BEMAFI32.DLL" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_ForcaImpactoAgulhas Lib "BEMAFI32.DLL" (ByVal ForcaImpacto As Integer) As Integer

' Funções do Cupom Fiscal
Public Declare Function Bematech_FI_AbreCupom Lib "BEMAFI32.DLL" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal quantidade As String, ByVal CasasDecimais As Integer, ByVal valorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal valorUnitario As String, ByVal quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BEMAFI32.DLL" (ByVal Descricao As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BEMAFI32.DLL" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_EstornoFormasPagamento Lib "BEMAFI32.DLL" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal valor As String) As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoDescricaoForma Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal Mensagem As String) As Integer

' Funções dos Relatórios Fiscais
Public Declare Function Bematech_FI_LeituraX Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BEMAFI32.DLL" (ByVal DATA As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_RelatorioGerencialTEF Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer

' Funções das Operações Não Fiscais
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal IndiceTotalizador As String, ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculadoTEF Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_Sangria Lib "BEMAFI32.DLL" (ByVal valor As String) As Integer
Public Declare Function Bematech_FI_Suprimento Lib "BEMAFI32.DLL" (ByVal valor As String, ByVal FormaPagamento As String) As Integer

' Funções de Informação da Impressora
Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BEMAFI32.DLL" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BEMAFI32.DLL" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmware Lib "BEMAFI32.DLL" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CGC_IE Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Bematech_FI_GrandeTotal Lib "BEMAFI32.DLL" (ByVal GrandeTotal As String) As Integer
Public Declare Function Bematech_FI_Cancelamentos Lib "BEMAFI32.DLL" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BEMAFI32.DLL" (ByVal ValorDescontos As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BEMAFI32.DLL" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Bematech_FI_NumeroCuponsCancelados Lib "BEMAFI32.DLL" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Bematech_FI_NumeroIntervencoes Lib "BEMAFI32.DLL" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Bematech_FI_NumeroReducoes Lib "BEMAFI32.DLL" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Bematech_FI_NumeroSubstituicoesProprietario Lib "BEMAFI32.DLL" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_ClicheProprietario Lib "BEMAFI32.DLL" (ByVal Cliche As String) As Integer
Public Declare Function Bematech_FI_NumeroCaixa Lib "BEMAFI32.DLL" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Bematech_FI_NumeroLoja Lib "BEMAFI32.DLL" (ByVal NumeroLoja As String) As Integer
Public Declare Function Bematech_FI_SimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_MinutosLigada Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_MinutosImprimindo Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_VerificaModoOperacao Lib "BEMAFI32.DLL" (ByVal Modo As String) As Integer
Public Declare Function Bematech_FI_VerificaEpromConectada Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BEMAFI32.DLL" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BEMAFI32.DLL" (ByVal ValorCupom As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal DATA As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BEMAFI32.DLL" (ByVal DATA As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BEMAFI32.DLL" (ByVal DATA As String) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_Acrescimos Lib "BEMAFI32.DLL" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Bematech_FI_ContadorBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Bematech_FI_VerificaAliquotasIss Lib "BEMAFI32.DLL" (ByVal AliquotasIss As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamento Lib "BEMAFI32.DLL" (ByVal Formas As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaDepartamentos Lib "BEMAFI32.DLL" (ByVal Departamentos As String) As Integer
Public Declare Function Bematech_FI_VerificaTipoImpressora Lib "BEMAFI32.DLL" (ByRef TipoImpressora As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciais Lib "BEMAFI32.DLL" (ByVal cTotalizadores As String) As Integer
Public Declare Function Bematech_FI_RetornoAliquotas Lib "BEMAFI32.DLL" (ByVal cAliquotas As String) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BEMAFI32.DLL" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_MonitoramentoPapel Lib "BEMAFI32.DLL" (ByRef Linhas As String) As Integer
Public Declare Function Bematech_FI_ValorFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal valor As String) As Integer
Public Declare Function Bematech_FI_ValorTotalizadorNaoFiscal Lib "BEMAFI32.DLL" (ByVal Totalizador As String, ByVal valor As String) As Integer

' Funções de Autenticação
Public Declare Function Bematech_FI_Autenticacao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BEMAFI32.DLL" (ByVal Parametros As String) As Integer

' Funções de Gaveta de Dinheiro
Public Declare Function Bematech_FI_AcionaGaveta Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoGaveta Lib "BEMAFI32.DLL" (ByRef EstadoGaveta As Integer) As Integer

' Funções de Impressão de Cheques
Public Declare Function Bematech_FI_ProgramaMoedaSingular Lib "BEMAFI32.DLL" (ByVal MoedaSingular As String) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaPlural Lib "BEMAFI32.DLL" (ByVal MoedaPlural As String) As Integer
Public Declare Function Bematech_FI_CancelaImpressaoCheque Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaStatusCheque Lib "BEMAFI32.DLL" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Bematech_FI_ImprimeCheque Lib "BEMAFI32.DLL" (ByVal Banco As String, ByVal valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal DATA As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_IncluiCidadeFavorecido Lib "BEMAFI32.DLL" (ByVal Cidade As String, ByVal Favorecido As String) As Integer

' Outras Funções
Public Declare Function Bematech_FI_ResetaImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_MapaResumo Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Analitico Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Mestre Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeConfiguracoesImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AberturaDoDia Lib "BEMAFI32.DLL" (ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_FechamentoDoDia Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImpressaoCarne Lib "BEMAFI32.DLL" (ByVal Titulo As String, ByVal Percelas As String, ByVal Datas As Integer, ByVal quantidade As Integer, ByVal Texto As String, ByVal Cliente As String, ByVal RG_CPF As String, ByVal Cupom As String, ByVal Vias As Integer, ByVal Assina As Integer) As Integer
Public Declare Function Bematech_FI_InfoBalanca Lib "BEMAFI32.DLL" (ByVal Porta As String, ByVal Modelo As Integer, ByVal Peso As String, ByVal PrecoKilo As String, ByVal total As String) As Integer
Public Declare Function Bematech_FI_DadosSintegra Lib "BEMAFI32.DLL" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Public Declare Function Bematech_FI_IniciaModoTEF Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FinalizaModoTEF Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VersaoDll Lib "BEMAFI32.DLL" (ByVal Versao As String) As Integer
Public Declare Function Bematech_FI_RegistrosTipo60 Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeArquivoRetorno Lib "BEMAFI32.DLL" (ByVal retorno As String) As Integer

'********************************************************************************************
'
'       Fim das Declarações das Funções da Impressora Fiscal
'
'********************************************************************************************
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim Fim
Dim wValorComplemento As String
Dim wDesconto As Double
Dim TotalComDesconto As Double
Dim CondPag, total As String
Dim i, cont, n, qtotal, limite As Integer

Public Function VerificaRetornoImpressora(Label As String, RetornoFuncao As String, TituloJanela As String)
    
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim RetornaMensagem As Integer
    Dim StringRetorno As String
    Dim ValorRetorno As String
    Dim RetornoStatus As Integer
    Dim Mensagem As String
    
    wVerificaImpressoraFiscal = False
    
    If retorno = 0 Then
        MsgBox "Erro de comunicação com a impressora.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    
    ElseIf retorno = 1 Then
        RetornoStatus = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ValorRetorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
        
        If Label <> "" And RetornoFuncao <> "" Then
            RetornaMensagem = 1
        End If
        
        If ACK = 21 Then
            MsgBox "Status da Impressora: 21" & vbCr & vbLf & "Comando não executado", vbOKOnly + vbInformation, TituloJanela
            Exit Function
        End If
        
        If (ST1 <> 0 Or ST2 <> 0) Then
                If (ST1 >= 128) Then
                    StringRetorno = "Fim de Papel" & vbCr
                    ST1 = ST1 - 128
                End If
                
                If (ST1 >= 64) Then
                    StringRetorno = StringRetorno & "Pouco Papel" & vbCr
                    ST1 = ST1 - 64
                End If
                
                If (ST1 >= 32) Then
                    StringRetorno = StringRetorno & "Erro no relógio" & vbCr
                    ST1 = ST1 - 32
                End If
                
                If (ST1 >= 16) Then
                    StringRetorno = StringRetorno & "Impressora em erro" & vbCr
                    ST1 = ST1 - 16
                End If
                    
                If (ST1 >= 8) Then
                    StringRetorno = StringRetorno & "Primeiro dado do comando não foi Esc" & vbCr
                    ST1 = ST1 - 8
                End If
                
                If (ST1 >= 4) Then
                    StringRetorno = StringRetorno & "Comando inexistente" & vbCr
                    ST1 = ST1 - 4
                End If
                    
                If (ST1 >= 2) Then
                    StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
                    ST1 = ST1 - 2
                End If
                
                If (ST1 >= 1) Then
                    StringRetorno = StringRetorno & "Número de parâmetros inválido no comando" & vbCr
                    ST1 = ST1 - 1
                End If
                    
                If (ST2 >= 128) Then
                    StringRetorno = "Tipo de Parâmetro de comando inválido" & vbCr
                    ST2 = ST2 - 128
                End If
                
                If (ST2 >= 64) Then
                    StringRetorno = StringRetorno & "Memória fiscal lotada" & vbCr
                    ST2 = ST2 - 64
                End If
                
                If (ST2 >= 32) Then
                    StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
                    ST2 = ST2 - 32
                End If
                
                If (ST2 >= 16) Then
                    StringRetorno = StringRetorno & "Alíquota não programada" & vbCr
                    ST2 = ST2 - 16
                End If
                    
                If (ST2 >= 8) Then
                    StringRetorno = StringRetorno & "Capacidade de alíquota programáveis lotada" & vbCr
                    ST2 = ST2 - 8
                End If
                
                If (ST2 >= 4) Then
                    StringRetorno = StringRetorno & "Cancelamento não permitido" & vbCr
                    ST2 = ST2 - 4
                End If
                    
                If (ST2 >= 2) Then
                    StringRetorno = StringRetorno & "CGC/IE do proprietário não programados" & vbCr
                    ST2 = ST2 - 2
                End If
                
                If (ST2 >= 1) Then
                    StringRetorno = StringRetorno & "Comando não executado" & vbCr
                    ST2 = ST2 - 1
                End If
                
                If RetornaMensagem Then
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                           vbCr & vbLf & StringRetorno & vbCr & vbLf & _
                           Label & RetornoFuncao
                Else
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                       vbCr & vbLf & StringRetorno
                End If
        
                MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
                Exit Function
        End If 'fim do ST1 <> 0 and ST2 <> 0
        
        If RetornaMensagem Then
            Mensagem = Label & RetornoFuncao
        End If
        
        If Mensagem <> "" Then
            MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
        End If
        Exit Function
    ElseIf retorno = -1 Then
        MsgBox "Erro de execução da função.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    ElseIf retorno = -2 Then
        MsgBox "Parâmetro inválido na função.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf retorno = -3 Then
        MsgBox "Alíquota não programada.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf retorno = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default. " + vbCr + "Por favor, copie esse arquivo para o diretório de sistema do Windows." + vbCr + "Se for o Windows 95 ou 98 é o diretório 'System' se for o Windows NT é o diretório 'System32'.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf retorno = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação.", vbOKOnly + vbExclamation, TituloJanela
        retorno = Bematech_FI_ResetaImpressora()
        Exit Function
    ElseIf retorno = -6 Then
        MsgBox "Impressora desligada ou cabo de comunicação desconectado.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf retorno = -7 Then
        MsgBox "Banco não encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf retorno = -8 Then
        MsgBox "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    End If
    wVerificaImpressoraFiscal = True
   
End Function


Function ConverteVirgula(ByVal Numero As String) As String

    Dim Ret As String
    Dim CharLido As String
    Dim Maximo As Long
    Dim i As Long
    
    Ret = "0"
    Numero = IIf(IsNull(Numero), 0, Numero)
    Maximo = Len(Numero)
    
    For i = 1 To Maximo
        CharLido = Mid(Numero, i, 1)
        
        If IsNumeric(CharLido) Then
            Ret = Ret & CharLido
        ElseIf CharLido = "," And InStr(Ret, ".") = 0 Then
            Ret = Ret & "."
        End If
    Next
    
    ConverteVirgula = Ret

End Function
Function ConverteVirgula1(ByVal Expressao) As String
    Dim ContPad As String
    Dim flgpad As Integer
    
    If Len(Expressao) <> 0 Then
        ContPad = CStr(Expressao)
        flgpad = InStr(ContPad, ".")
        Do While flgpad <> 0
            Mid(ContPad, flgpad, 1) = ","
            flgpad = InStr(ContPad, ".")
        Loop
    Else
        ContPad = 0
    End If
    ConverteVirgula1 = ContPad
End Function


Public Function inibebotoes(ByVal PedidoTXT As String)

    SQL = ""
    SQL = "Select * From NFCapa Where NumeroPed = " & PedidoTXT & " and TipoNota = 'PD'"
    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   
    If Not rsComplementoVenda.EOF Then
    
       Select Case Not rsComplementoVenda.EOF
            Case Format(rsComplementoVenda("Desconto"), "##0.00") <> "0,00"
                 frmPedido.cmdBotoes(8).Visible = False
                 frmPedido.cmdBotoes(9).Visible = False
    
            Case rsComplementoVenda("Cliente") <> 999999
                 frmPedido.cmdTR.Visible = False
                 frmPedido.cmdBotoes(8).Visible = False
                 frmPedido.cmdBotoes(10).Visible = False
                 
       End Select
    End If
    rsComplementoVenda.Close
End Function


Public Function Numeros(ByVal Texto As String) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim CharLido As String * 1
    Dim retorno As String
    
    Maximo = Len(Texto)
    
    retorno = ""
    For Char = 1 To Maximo Step 1
        CharLido = Mid(Texto, Char, 1)
        If IsNumeric(CharLido) Then
            retorno = retorno & CharLido
        End If
    Next Char
    
    Texto = retorno
    
    Numeros = Texto

End Function


Public Function ChecaCaracterDigitado(ByVal Texto As String)
' Se o campo conter um dos caracteres abaixo retorna True

CaracterDigitado = Texto Like "*'*" _
                Or Texto Like "*,*"
End Function

Function AchaLojaControle() As String
       
    Dim SQL As String
    
    SQL = "Select CTS_Loja from ControleSistema"
    rsControleLoja.CursorLocation = adUseClient
    rsControleLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    AchaLojaControle = rsControleLoja("CTS_Loja")
       
    rsControleLoja.Close
   
End Function

Public Sub CriaCotacaoHtml(ByVal Pedido As Double)
    
Dim vendedor As String
Dim NomeVend As String
Dim Razao As String
Dim CepLoja As String
Dim UfLoja As String
Dim TelefoneLoja As String
Dim EndLoja As String
Dim BairroLoja As String
Dim InfLoja As String
Dim IntFile1
Dim NomeArquivo As String
Dim Logo  As String
Dim SomaPedido As Double

    
    
    CepLoja = ""
    UfLoja = ""
    TelefoneLoja = ""
    EndLoja = ""
    BairroLoja = ""
    Razao = ""
    
    SQL = ""
    SQL = "Select CTS_ValidadeCotacao From ControleSistema"
    rdoValidadeCotacao.CursorLocation = adUseClient
    rdoValidadeCotacao.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
   
    
   
    NomeArquivo = "Cot" & Pedido
    

    IntFile1 = FreeFile()
'    GLB_Cotacao = "C:\Cotacao\"
    Open GLB_Cotacao & NomeArquivo & ".html" For Output Access Write As #IntFile1
     
    Print #IntFile1, "<html>"
    Print #IntFile1, "<head>"
    Print #IntFile1, "<title>" & Razao & "</title>"
    Print #IntFile1, "</head>"
    
    SQL = ""
'    SQL = "Select  COV_ValorComplemento from  ComplementoVenda" _
'          & " where COV_NumeroPedido =" & Pedido & " and COV_SequenciaComplemento = 6"
          
    SQL = "SELECT Cliente From NFCapa Where Numeroped = " & Pedido
    
          rsComplemento.CursorLocation = adUseClient
          rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
    If Not rsComplemento.EOF Then
'       wValorComplemento = Trim(rsComplemento("COV_ValorComplemento"))
        wValorComplemento = Trim(rsComplemento.Fields("Cliente"))
    Else
       wValorComplemento = 999999
    End If

    rsComplemento.Close
        sql1 = ""
   
        sql1 = "Select * from FIN_Cliente Where CE_CodigoCliente= '" & wValorComplemento & "'"
            rsCliente.CursorLocation = adUseClient
            rsCliente.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsCliente.EOF Then
       
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=39>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=" & GLB_logoPedido & " width=245 height=80 align=left>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "<td width=333 height=1>"
        Print #IntFile1, "<font face=System color=#000080>" & UCase(rsCliente("CE_Razao")) & " </font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>" & UCase(rsCliente("CE_Endereco")) & ", " & UCase(rsCliente("CE_Numero")) & " - " & UCase(rsCliente("CE_Bairro")) & "</font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>" & UCase(right(String(8, "0") & rsCliente("CE_CEP"), 8)) & " - " & UCase(rsCliente("CE_Municipio")) & " - " & UCase(rsCliente("CE_estado")) & " - Fone: " & UCase(rsCliente("CE_Telefone")) & " </font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "<div align=justify style=width: 750; height: 56>"
        Print #IntFile1, "<div align=justify>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=336 height=41>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=21><font size=1 color=#000080 face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & EndLoja & " - " & BairroLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=20><font face=Arial size=1 color=#000080>SAO PAULO - " & UfLoja & " - CEP : " & UCase(right(String(8, "0") & CepLoja, 8)) & " - FONE : " & TelefoneLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=48>"
        Print #IntFile1, "<tr>"
'        Print #IntFile1, "<td width=700 rowspan=3><img border=0 src=http://www.demeonews.com.br/biblioteca/fotos/11726_PedidodeCompras_04.jpg width=750 height=40 align=right></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<div align=left>"
        Print #IntFile1, "<table border=1 cellpadding=1 cellspacing=0 width=759 height=36 solid; border-width: 0 bordercolor=#000080>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=640 height=20><font face=Arial color=#000080><b>Referencia / Descrição</b></font></td>"
        Print #IntFile1, "<td width=41 height=20><font face=Arial color=#000080><b>Qtde</b></font></td>"
        Print #IntFile1, "<td width=76 height=20><font face=Arial color=#000080><b>Valor</b></font></td>"
        
        Print #IntFile1, "<td width=79 height=20><font face=Arial color=#000080><b>Total</b></font></td>"
        Print #IntFile1, "</tr>"
    Else
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=39>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src= " & GLB_logoPedido & "; Width = 245; Height = 80; Align = left > """
        Print #IntFile1, "</td>"
        Print #IntFile1, "<td width=333 height=1>"
        Print #IntFile1, "<font face=System color=#000080>CONSUMIDOR</font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>CONSUMIDOR</font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>00000000 - CONSUMIDOR </font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "<div align=justify style=width: 750; height: 56>"
        Print #IntFile1, "<div align=justify>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=336 height=41>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=21><font size=1 color=#000080 face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & EndLoja & " - " & BairroLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=20><font face=Arial size=1 color=#000080>SAO PAULO - " & UfLoja & " - CEP : " & UCase(right(String(8, "0") & CepLoja, 8)) & " - FONE : " & TelefoneLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=48>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=700 rowspan=3><img border=0 src=http://www.demeo.com.br/images/truck.jpg width=750 height=40 align=right></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<div align=left>"
        Print #IntFile1, "<table border=1 cellpadding=1 cellspacing=0 width=761 height=36 solid; border-width: 0 bordercolor=#000080>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=640 height=20><font face=Arial color=#000080><b>Referencia / Descrição</b></font></td>"
        Print #IntFile1, "<td width=41 height=20><font face=Arial color=#000080><b>Qtde</b></font></td>"
        Print #IntFile1, "<td width=76 height=20><font face=Arial color=#000080><b>Valor</b></font></td>"
        Print #IntFile1, "<td width=66 height=20><font face=Arial color=#000080><b>Desconto</b></font></td>"
        Print #IntFile1, "<td width=79 height=20><font face=Arial color=#000080><b>Total</b></font></td>"
        Print #IntFile1, "</tr>"
    End If
   
    
    SQL = ""
    
'    SQL = "Select PR_Descricao,ITV_Quantidade,ITV_PrecoUnitario,(ITV_PrecoUnitario * ITV_Quantidade) as VlUnit2, " _
'        & "PR_Referencia,VEN_NomeVendedor " _
'        & "From Produto, NFItens as i,nfcapa as c, Vende   " _
'        & "Where PR_Referencia = ITV_CodigoProduto " _
'        & "and ITV_NumeroPedido=" & Pedido & " " _
'        & "and VEN_CodigoVendedor=ITV_Vendedor "
        
    SQL = "Select PR_Descricao, Qtde, VLUnit, (VLUnit * Qtde) as VLUnit2, PR_Referencia, VE_Nome,VE_Codigo " _
          & "From ProdutoLoja, NFItens as i,nfcapa as c, Vende  " _
          & "Where PR_Referencia = Referencia and VE_Codigo = c.Vendedor and c.NumeroPed = i.NumeroPed and " _
          & "i.NumeroPed = " & Pedido
        
    rsPedido.CursorLocation = adUseClient
    rsPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
   
    SomaPedido = 0
    NomeVend = ""
    vendedor = ""
    CondPag = ""
    EndLoja = ""
    CepLoja = ""
    TelefoneLoja = ""
    BairroLoja = ""
    UfLoja = ""
    If Not rsPedido.EOF Then
        NomeVend = Trim(rsPedido("VE_Nome"))
        vendedor = Trim(rsPedido("VE_Codigo"))
        SQL = ""
'        SQL = "Select  COV_ValorComplemento from  ComplementoVenda" _
'            & " where COV_NumeroPedido =" & Pedido & " and COV_SequenciaComplemento = 4"
 ''       SQL = "Select CP_Condicao, Desconto From CondicaoPagamento, NFCapa " _
''            & "Where CondPag = CP_Codigo And NumeroPed = " & Pedido
          SQL = "select modalidadevenda, desconto,condpag,VendedorLojaVenda from nfcapa where NumeroPed = " & Pedido
             
             rsComplemento.CursorLocation = adUseClient
             rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
        If Not rsComplemento.EOF Then

            CondPag = IIf(IsNull(rsComplemento.Fields("modalidadevenda")), "A VISTA", rsComplemento.Fields("modalidadevenda"))
            wDesconto = IIf(IsNull(rsComplemento.Fields("Desconto")), "0", rsComplemento.Fields("Desconto"))
 '            Vendedor = rsComplemento("VendedorLojaVenda")
            
        Else
        CondPag = "A VISTA"
        End If
        rsComplemento.Close
        
        
        Do While Not rsPedido.EOF
            Print #IntFile1, "<tr>"
            Print #IntFile1, "<td width=640 height=16><font face=Verdana size=2 color=#000080>" & rsPedido("PR_Referencia") & " " & rsPedido("PR_Descricao") & "</font></td>"
'            Print #IntFile1, "<td width=41 height=16 align=right><font face=Verdana size=2 color=#000080>" & rsPedido("ITV_Quantidade") & "</font></td>"
            Print #IntFile1, "<td width=41 height=16 align=right><font face=Verdana size=2 color=#000080>" & rsPedido("Qtde") & "</font></td>"
'            Print #IntFile1, "<td width=76 height=16 align=right><font face=Verdana size=2 color=#000080>" & Format(rsPedido("ITV_precounitario"), "##,###,###0.00") & "</font></td>"
            Print #IntFile1, "<td width=76 height=16 align=right><font face=Verdana size=2 color=#000080>" & Format(rsPedido("VLUnit"), "##,###,###0.00") & "</font></td>"
            
            Print #IntFile1, "<td width=79 height=16 align=right><font face=Verdana size=2 color=#000080>" & Format(rsPedido("VlUnit2"), "##,###,###0.00") & "</font></td>"
            
            Print #IntFile1, "</tr>"
            SomaPedido = SomaPedido + Format(rsPedido("VlUnit2"), "##,###,###0.00")
           
            rsPedido.MoveNext
        Loop
    End If
'    SQL = ""
'         SQL = "Select  COV_ValorComplemento from  ComplementoVenda" _
'            & " where COV_NumeroPedido =" & Pedido & " and COV_SequenciaComplemento = 15"
'                rsComplemento.CursorLocation = adUseClient
'                rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
                
'                If Not rsComplemento.EOF Then
'                   wDesconto = ConverteVirgula1(Trim(rsComplemento("COV_Valorcomplemento")))
'                End If
        
         TotalComDesconto = SomaPedido - wDesconto
'         rsComplemento.Close
    
    Print #IntFile1, "</tr>"
    Print #IntFile1, "</table>"
    Print #IntFile1, "</div>"
    Print #IntFile1, "</p>"
    Print #IntFile1, "<TR><TD><PRE> &nbsp;&nbsp;<font face=Arial size=3 color=#000080>                                                                                                                                        Sub Total                   " & Format(SomaPedido, "###,###,###,##0.00") & "</font></pre></TD></TR>"
   
    If wDesconto > 0 Then
    Print #IntFile1, "<TR><TD><PRE> &nbsp;&nbsp;<font face=Arial size=3 color=#000080>                                                                                                                                        Desconto                     " & Format(wDesconto, "###,###,###,##0.00") & "</font></pre></TD></TR>"
    End If
   
    Print #IntFile1, "<TR><TD><PRE> &nbsp;&nbsp;<font face=Arial size=4 color=#000080>                                                                                                             Total                   " & Format(TotalComDesconto, "###,###,###,##0.00") & "</font></pre></TD></TR>"
  
     
    Print #IntFile1, "</table>"
    Print #IntFile1, "<p>&nbsp;</p>"
    Print #IntFile1, "<p>&nbsp;</p>"
    Print #IntFile1, "<div align=justify>"
    Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=62>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=259 height=21><font color=#000080 face=System>COND PAGTO&nbsp;&nbsp; : " & UCase(CondPag) & UCase(wGuardaPagamento) & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=259 height=21><font color=#000080 face=System>VALIDADE&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : " & Format(DateAdd("D", rdoValidadeCotacao("CTs_ValidadeCotacao"), Date), "dd/mm/yyyy") & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=259 height=20><font color=#000080 face=System>VENDEDOR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: " & UCase(vendedor) & " - " & UCase(NomeVend) & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "</table>"
    Print #IntFile1, "</div>"
    
    Print #IntFile1, "<tr>"
    Print #IntFile1, "</td>"
    Print #IntFile1, "<td width=333 height=1>"
        
    
    Print #IntFile1, "</p>"
  
      
    Print #IntFile1, "</table>"
    

    Close #IntFile1
    
    rdoValidadeCotacao.Close
    rsInfLoja.Close
    rsPedido.Close
    rsCliente.Close
    

End Sub


Public Function ImprimirCotacao(ByVal Pedido As Double)
    
    Dim VarImp As String
    Dim pagina As Integer
    Dim SubTotal As Double
    Dim total As Double
    Dim Desconto As Double
    Dim Linhas As Integer
    Dim Descricao As String
    Dim Referencia As String
    Dim VlUnit As Double
    Dim VlUnit2 As Double
    Dim VlTotItem As Double
    Dim DescontoItem As Double
    Dim vendedor As Integer
    Dim NomeVend As String
    Dim Espacos As String
    
         SQL = ""
        SQL = "select modalidadevenda, desconto,condpag from nfcapa where NumeroPed = " & Pedido
         
                rsComplemento.CursorLocation = adUseClient
                rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
         If Not rsComplemento.EOF Then
        CondPag = RTrim(LTrim(IIf(IsNull(rsComplemento.Fields("modalidadevenda")), "A VISTA", rsComplemento.Fields("modalidadevenda"))))
        End If
        rsComplemento.Close
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = GLB_ImpCotacao Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    Screen.MousePointer = 11
    pagina = 1
    SubTotal = 0
    total = 0
    Desconto = 0
    Linhas = 7
    CabecalhoCotacao Pedido, pagina
            
    SQL = ""
    
        
    SQL = "Select VE_Codigo, VE_Nome, PR_Descricao, I.Qtde, I.Vlunit,(I.VLUnit * I.Qtde) as VLUnit2,PR_Referencia, C.Desconto as Desconto " _
        & "From ProdutoLoja, NFItens as I, NFCapa as C, Vende " _
        & "Where PR_Referencia = I.Referencia and VE_Codigo = C.Vendedor and I.NumeroPed = C.NumeroPed and " _
        & " I.DataEmi = C.DataEmi and C.NumeroPed = " & Pedido
        
    rdoPedido.CursorLocation = adUseClient
    rdoPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    vendedor = rdoPedido("VE_Codigo")
    NomeVend = rdoPedido("VE_Nome")
    Desconto = rdoPedido("Desconto")
    
    If Not rdoPedido.EOF Then
        Do While Not rdoPedido.EOF
           Descricao = rdoPedido("PR_Descricao")
           Referencia = rdoPedido("PR_Referencia")
           Linhas = Linhas + 1
           VlUnit = rdoPedido("VLUnit") 'rdoPedido("ITV_PrecoUnitario")
           VlUnit2 = rdoPedido("VlUnit2")
           SubTotal = SubTotal + (rdoPedido("VLUnit") * rdoPedido("Qtde")) '(rdoPedido("ITV_PrecoUnitario") * rdoPedido("ITV_Quantidade"))
           total = total + rdoPedido("VlUnit2")
           VarImp = left(Referencia & Space(11), 11) _
                & left(Descricao & Space(42), 42) _
                & right(Space(6) & rdoPedido("QTDE"), 6) _
                & right(Space(12) & Format(VlUnit, "###,###,###,##0.00"), 12) _
                & right(Space(10) & Espacos, 10) _
                & right(Space(12) & Format(VlUnit2, "###,###,###,##0.00"), 12)
            Printer.Print VarImp
            If Linhas = 62 Then
                Printer.Print "_________________________________________________________________________________________________________________________"
                Printer.NewPage
                CabecalhoCotacao Pedido, pagina + 1
                Linhas = 8
            End If
            rdoPedido.MoveNext
        Loop
        If Desconto = 0 Then
            SubTotal = total
        End If
        FinalizaCotacao Linhas, SubTotal, Desconto, total, vendedor, NomeVend
    Else
        Screen.MousePointer = 0
        MsgBox "Impossivel imprimir cotação", vbExclamation, "Aviso"
        Exit Function
    End If
    Screen.MousePointer = 0
    rdoPedido.Close
End Function



Public Function CabecalhoCotacao(ByVal Pedido As Double, ByVal pagina As Integer)
    'Dim rdoPedido As rdoResultset
    Dim VarImp As String

    SQL = ""
'    SQL = "Select Itensvenda.*,Lojas.* from Itensvenda,Lojas " _
'        & "where ITV_NumeroPedido=" & Pedido & " and Lo_Loja=ITV_Loja "
        
''    SQL = "Select NFItens.*,Lojas.*,Cliente.* From NFItens, Lojas, Cliente " _
''          & " Where Cliente = CE_CodigoCliente and LojaOrigem = LO_Loja and NumeroPed = " & Pedido
''        rdoPedido.CursorLocation = adUseClient
''        rdoPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''

    SQL = "Select NFCapa.*,Loja.*,fin_Cliente.* From NFCapa, Loja, fin_Cliente " _
          & " Where Cliente = CE_CodigoCliente and LojaOrigem = LO_Loja and NumeroPed = " & Pedido
        rdoPedido.CursorLocation = adUseClient
        rdoPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        
'   SQL = ""
'   SQL = "Select  COV_ValorComplemento from  ComplementoVenda" _
'         & " where COV_NumeroPedido =" & Pedido & " and COV_SequenciaComplemento = 6"
'         rsComplemento.CursorLocation = adUseClient
'         rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
'    If Not rsComplemento.EOF Then
'       wValorComplemento = Trim(rsComplemento("COV_ValorComplemento"))
'    Else
'       wValorComplemento = 999999
'    End If

'    rsComplemento.Close
'        SQL1 = ""
   
'        SQL1 = "Select * from FIN_ClienteWhere CE_CodigoCliente= '" & wValorComplemento & "'"
'            rsCliente.CursorLocation = adUseClient
'            rsCliente.Open SQL1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
                         
          
    If Not rdoPedido.EOF Then
        Printer.ScaleMode = vbMillimeters
        Printer.ForeColor = "0"
        Printer.FontSize = 8
        Printer.FontName = "draft 20cpi"
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.DrawWidth = 3
        Printer.FontName = "COURIER NEW"
        Printer.FontSize = 10#
        Printer.Print "COTACAO DE VENDA" & Space(35) & "NUMERO: " & right(String(6, "0") & Pedido, 6) & Space(3) & "Data: " & Format(Date, "yyyy/mm/dd") & Space(5) & "PAG: " & pagina
        Printer.Print "_________________________________________________________________________________________________________________________"
'        VarImp = Left(rdoPedido("LO_Razao") & Space(45), 45) & Left(rsCliente("CE_Razao") & Space(60), 60)
        VarImp = left(rdoPedido("LO_Razao") & Space(45), 45) & left(rdoPedido("CE_Razao") & Space(60), 60)
        Printer.Print VarImp
        VarImp = left(rdoPedido("LO_Endereco") & ", " & rdoPedido("LO_Numero") & Space(45), 45) & left(rdoPedido("CE_Endereco") & ", " & rdoPedido("CE_Numero") & Space(60), 60)
        Printer.Print VarImp
        VarImp = left(right(String(7, "0") & rdoPedido("LO_CEP"), 7) & " - " & rdoPedido("LO_Bairro") & "  -  " & rdoPedido("LO_Municipio") & " - " & rdoPedido("LO_UF") & Space(45), 45)
        VarImp = VarImp & left(rdoPedido("CE_CEP") & " - " & rdoPedido("CE_Bairro") & "  -  " & rdoPedido("CE_Municipio") & " - " & rdoPedido("CE_Estado") & Space(60), 60)
        Printer.Print VarImp
        VarImp = left("Telefone " & rdoPedido("LO_Telefone") & Space(45), 45) & left("Telefone " & rdoPedido("CE_Telefone") & Space(60), 60)
        Printer.Print VarImp
        Printer.Print "_________________________________________________________________________________________________________________________"
        Printer.Print ""
        Printer.Print "REFERENCIA DESCRICAO                                   QTDE  PRECO UNIT            PRECO TOTAL"
        Printer.Print ""
        
    End If
       rdoPedido.Close
  '     rsCliente.Close

End Function
Function FinalizaCotacao(ByVal Linhas As Integer, ByVal SubTotal As Double, ByVal Desconto As Double, ByVal total As Double, ByVal vendedor As String, ByVal NomeVend As String)
    'Dim rdoDescPag As rdoResultset
   ' Dim rdoValidadeCotacao As rdoResultset
    Dim desc As String
        
    SQL = ""
    SQL = "Select CTS_ValidadeCotacao From ControleSistema"
    rdoValidadeCotacao.CursorLocation = adUseClient
    rdoValidadeCotacao.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    'rdoValidadeCotacao ("CTS_ValidadeCotacao")
    rdoValidadeCotacao.Close
    
   ' SQL = ""
   ' SQL = "Select CP_Condicao from CondicaoPagamento " _
   '     & "where CP_Codigo=" & CondPag & ""
   ' Set rdoDescPag = adoCNLoja.OpenResultset(SQL)
   ' If Not rdoDescPag.EOF Then
    ' Desc = rdoDescPag("CP_Condicao")
    '   rdoDescPag.Close
    'End If
    For Linhas = Linhas To 56
        Printer.Print ""
    Next
    
    Printer.Print "_________________________________________________________________________________________________________________________"
    Printer.Print "COND PAGTO : " & left(CondPag & wGuardaPagamento & Space(42), 42) & "SUB-TOTAL" & Space(9) & "DESCONTO" & Space(9) & "TOTAL"
    Printer.Print "VALIDADE   : " & left(Format(DateAdd("D", rdoValidadeCotacao("CTS_ValidadeCotacao"), Date), "yyyy/mm/dd") & Space(12), 12) & Space(24) & right(Space(15) & Format(SubTotal, "###,###,###,##0.00"), 15) _
        & Space(2) & right(Space(14) & Format(Desconto, "###,###,###,##0.00"), 14) & right(Space(15) & Format((SubTotal - Desconto), "###,###,###,##0.00"), 15)
    Printer.Print "VENDEDOR   : " & vendedor & " - " & NomeVend
    Printer.Print "_________________________________________________________________________________________________________________________"
    Printer.EndDoc
    
    rdoValidadeCotacao.Close
    
End Function

Function DescricaoOperacao(ByVal Descricao As String)

'    mdiBalcao.stbBarra.Panels.Item(1).Text = Descricao

End Function
Sub Esperar(ByVal Tempo As Integer)
    
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
        DoEvents
    Loop

End Sub

Public Function ExtraiSeqNotaControle() As Double
     Dim WnovaSeqNota As Long
     

'On Error GoTo ErroSeqNotaControle


     
     
     SQL = ""
     SQL = "Select (CTS_NumeroNF + 1) as NumNota from ControleSistema"
        
     adoCNLoja.BeginTrans
  If RsDados.State = 1 Then
    RsDados.Close
 End If
     
     RsDados.CursorLocation = adUseClient
     RsDados.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     If Not RsDados.EOF Then
        ExtraiSeqNotaControle = RsDados("NumNota")
        SQL = "update ControleSistema set CTS_NumeroNF= " & RsDados("NumNota") & ""
        adoCNLoja.Execute (SQL)
     End If
     adoCNLoja.CommitTrans
     RsDados.Close
     Exit Function

End Function
Public Function EmiteNotafiscal(ByVal Nota As Double, ByVal Serie As String)
Dim wControlaQuebraDaPagina As Integer
wControlaQuebraDaPagina = 0

    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(Glb_ImpNotaFiscal) Then
           ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
      
    wSerie = Serie
    wNotaTransferencia = False
    wPagina = 0

    Call DadosLoja
            
    SQL = "select qtditem from nfcapa Where NumeroPed = " & frmPedido.txtPedido.Text
    rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            
    SQL = ""
   
    SQL = "Select NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV," _
        & "NFCAPA.Nf,NFCAPA.BASEICMS, NFCAPA.Serie, NFCAPA.PAGINANF, " _
        & "NFCAPA.volume,NFCAPA.PESOBR, NFCAPA.PESOLQ,  " _
        & "NFCAPA.CLIENTE,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.VLRMERCADORIA,Nfcapa.nf,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.lojaOrigem,NFCapa.PgEntra," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,LOJA.*,NFCAPA.CONDPAG, " _
        & "NfCapa.DataPag,NFCAPA.TOTALNOTAALTERNATIVA,NFCAPA.VALORTOTALCODIGOZERO," _
        & "NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS,NfItens.TipoNota,NfCapa.EmiteDataSaida " _
        & "From NFCAPA,NFITENS,LOJA " _
        & "Where NfCapa.nf= " & Nota & " and NfCapa.Serie in ('" & Serie & "') " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "' " _
        & "and NfItens.LojaOrigem=NfCapa.LojaOrigem " _
        & "and NfItens.Serie=NfCapa.Serie " _
        & "and NfItens.Nf=NfCapa.NF " _
        & "and ltrim(rtrim(convert(char(5),NFCAPA.CLIENTE))) = LOJA.LO_LOJA"


    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    If Not RsDados.EOF Then
      Cabecalho "T"
      
      SQL = "Select produtoloja.pr_referencia,produtoloja.pr_descricao, " _
          & "produtoloja.pr_classefiscal,produtoloja.pr_unidade,produtoloja.pr_st, " _
          & "produtoloja.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota," _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao,nfitens.CSTICMS," _
          & "nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produtoloja,nfitens " _
          & "where produtoloja.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & Nota & " and Serie='" & Serie & "' order by nfitens.item"
     
      rsItensVenda.CursorLocation = adUseClient
      rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

      If Not rsItensVenda.EOF Then
         wConta = 0
         wContItem = 0
         Printer.Print ""
         Do While Not rsItensVenda.EOF
            wContItem = wContItem + 1
            wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = rsItensVenda("PR_Referencia")
           
                     
           wPegaDescricaoAlternativa = IIf(IsNull(rsItensVenda("DescricaoAlternativa")), rsItensVenda("PR_Descricao"), rsItensVenda("DescricaoAlternativa"))
           If wPegaDescricaoAlternativa = "" Then
               wPegaDescricaoAlternativa = "0"
           End If
           If wPegaDescricaoAlternativa <> "0" Then
               wDescricao = wPegaDescricaoAlternativa
           Else
               wDescricao = Trim(rsItensVenda("pr_descricao"))
           End If
                    
                    
                    wStr16 = ""
                    wStr16 = left$(rsItensVenda("pr_referencia") & Space(7), 7) _
                          & Space(2) & left$(Format(Trim(wDescricao), ">") & Space(55), 55) _
                          & left$(Format(Trim(rsItensVenda("pr_classefiscal")), ">") _
                          & Space(11), 11) & left$(Trim("0" + Format(rsItensVenda("CSTICMS"), "00")) & Space(5), 5) _
                          & left$(Trim(rsItensVenda("pr_unidade")) & Space(2), 2) _
                          & right$(Space(8) & Format(rsItensVenda("QTDE"), "##0"), 8) _
                          & right$(Space(13) & Format(rsItensVenda("vlunit"), "#####0.00"), 13) _
                          & right$(Space(13) & Format(rsItensVenda("VlTotItem"), "#####0.00"), 13) _
                          & right$(Space(4) & Format(rsItensVenda("pr_icmssaida"), "#0"), 4)
                                  

                   Printer.Print wStr16
                      
                      If rsItensVenda("DetalheImpressao") = "D" Then
                         wConta = wConta + 1

                      ElseIf rsItensVenda("DetalheImpressao") = "C" Then
                            
                        Do While wConta < 34
                            wConta = wConta + 1
                            Printer.Print ""
                        Loop
                         
                         wConta = 0

                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                         If wControlaQuebraDaPagina = 3 Then
                            Printer.Print ""
                            wControlaQuebraDaPagina = 0
                         End If

                         Cabecalho rsItensVenda("TipoNota")
                         Printer.Print ""
                         
                         
                       If wContItem = rsComplementoVenda("QTDITEM") Then
                          Call ImprimeCarimbo
                       End If

                      ElseIf rsItensVenda("DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         Call ImprimeCarimbo

                      Else
                         wConta = wConta + 1
                      End If
                       rsItensVenda.MoveNext
            Loop
         Else
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
        Exit Function
    End If
rsItensVenda.Close
RsDados.Close
rsComplementoVenda.Close

End Function
'Public Function EmiteNotafiscal(ByVal Nota As Double, ByVal Serie As String)
'Dim wControlaQuebraDaPagina As Integer
'wControlaQuebraDaPagina = 0
'    For Each NomeImpressora In Printers
'        If UCase(Trim(NomeImpressora.DeviceName)) = UCase(Trim(GLB_ImpressoraNota)) Then
'           ' Seta impressora no sistema
'            Set Printer = NomeImpressora
'            Exit For
'        End If
'    Next
'
'    wSerie = Serie
'    wNotaTransferencia = True
'    wPagina = 1
'
'    Call DadosLoja
'
'    SQL = ""
'    SQL = "Select NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV," _
'        & "NFCAPA.Nf,NFCAPA.BASEICMS, NFCAPA.Serie, NFCAPA.PAGINANF, " _
'        & "NFCAPA.CLIENTE,LOJAS.LO_Telefone,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR," _
'        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.VLRMERCADORIA,Nfcapa.nf,NfCapa.Desconto," _
'        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.lojaOrigem,NFCapa.PgEntra," _
'        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,LOJAS.LO_razao,LOJAS.LO_CGC,NFCAPA.CONDPAG, " _
'        & "LOJAS.LO_Endereco,LOJAS.LO_Municipio,LOJAS.LO_Bairro,LOJAS.LO_Cep,LOJAS.LO_InscricaoEstadual," _
'        & "NfCapa.DataPag,LOJAS.LO_UF,NFCAPA.TOTALNOTAALTERNATIVA,NFCAPA.VALORTOTALCODIGOZERO," _
'        & "NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
'        & "NFITENS.VLTOTITEM,NFITENS.ICMS,NfItens.TipoNota,NfCapa.EmiteDataSaida " _
'        & "From NFCAPA,NFITENS,LOJAS " _
'        & "Where NfCapa.nf= " & Nota & " and NfCapa.Serie in ('" & Serie & "') " _
'        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "' " _
'        & "and NfItens.LojaOrigem=NfCapa.LojaOrigem " _
'        & "and NfItens.Serie=NfCapa.Serie " _
'        & "and NfItens.Nf=NfCapa.NF " _
'        & "and NfCapa.lojaorigem = LOJAS.LO_LOJA"
'
'    RsDados.CursorLocation = adUseClient
'    RsDados.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'    If Not RsDados.EOF Then
'      Cabecalho "T"
'
'      SQL = "Select produtoloja.pr_referencia,produtoloja.pr_descricao, " _
'          & "produtoloja.pr_classefiscal,produtoloja.pr_unidade, " _
'          & "produtoloja.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota," _
'          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao," _
'          & "nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
'          & "from produtoloja,nfitens " _
'          & "where produtoloja.pr_referencia=nfitens.referencia " _
'          & "and nfitens.nf = " & Nota & " and Serie='" & Serie & "' order by nfitens.item"
'
'      rsItensVenda.CursorLocation = adUseClient
'      rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'      If Not rsItensVenda.EOF Then
'         wConta = 0
'         Do While Not rsItensVenda.EOF
'            wPegaDescricaoAlternativa = "0"
'            wDescricao = ""
'            wReferenciaEspecial = rsItensVenda("PR_Referencia")
'            If Wsm = True Then
'                 wPegaDescricaoAlternativa = IIf(IsNull(rsItensVenda("DescricaoAlternativa")), rsItensVenda("PR_Descricao"), rsItensVenda("DescricaoAlternativa"))
'
'                ' If RsDados("UFCliente") = "SP" Then
'                '    wAliqICMSInterEstadual = RsdadosItens("PR_ICMSSaida")
'                ' Else
'                '    wAliqICMSInterEstadual = GLB_AliquotaICMS
'                ' End If
'
'                'wAliqICMSInterEstadual = RsdadosItens("icms")
'
'                 ' Adilson --> Dentro de São Paulo pega do cadastro de produto
'                 '             Fora   de São Paulo pega da rotina AcharIcmsInterEstadual
'
'                 If RsDados("UFCliente") = "SP" Then
'                    wAliqICMSInterEstadual = rsItensVenda("icms")
'                 Else
'                  ' wAliqICMSInterEstadual = rsItensVenda("icmpdv")
'                    wAliqICMSInterEstadual = rsItensVenda("pr_icmssaida")
'                 End If
'
'                   wStr16 = ""
'                   wStr16 = Left$(rsItensVenda("ReferenciaAlternativa") & Space(7), 7) _
'                          & Space(1) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
'                          & Space(16) & Left$(Format(Trim(rsItensVenda("pr_classefiscal")), ">") _
'                          & Space(10), 10) & Left$(Trim(rsItensVenda("Tributacao")) & Space(3), 3) _
'                          & "" & Space(3) & Left$(Trim(rsItensVenda("pr_unidade")) & Space(2), 2) _
'                          & Right$(Space(6) & Format(rsItensVenda("QTDE"), "##0"), 6) _
'                          & Right$(Space(12) & Format(rsItensVenda("PrecoUnitAlternativa"), "#####0.00"), 14) _
'                          & Right$(Space(15) & Format((rsItensVenda("PrecoUnitAlternativa") * rsItensVenda("QTDE")), "#####0.00"), 15) & Space(1) _
'                          & Right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)
'            Else
'
'                   wPegaDescricaoAlternativa = IIf(IsNull(rsItensVenda("DescricaoAlternativa")), rsItensVenda("PR_Descricao"), rsItensVenda("DescricaoAlternativa"))
'                   If wPegaDescricaoAlternativa = "" Then
'                        wPegaDescricaoAlternativa = "0"
'                   End If
'                   If wPegaDescricaoAlternativa <> "0" Then
'                         wDescricao = wPegaDescricaoAlternativa
'                   Else
'                         wDescricao = Trim(rsItensVenda("pr_descricao"))
'                   End If
'
'                   'If RsDados("UFCliente") = "SP" Then
'                   '    wAliqICMSInterEstadual = RsdadosItens("PR_ICMSSaida")
'                   'Else
'                   '    wAliqICMSInterEstadual = GLB_AliquotaICMS
'                   'End If
'
'                   ' Adilson --> Dentro de São Paulo pega do cadastro de produto
'                 '             Fora   de São Paulo pega da rotina AcharIcmsInterEstadual
'
'                 If RsDados("LO_UF") = "SP" Then
'                    wAliqICMSInterEstadual = rsItensVenda("icms")
'                 Else
'                    wAliqICMSInterEstadual = rsItensVenda("icmpdv")
'                 End If
'
'                   wStr16 = ""
'                   wStr16 = Left$(rsItensVenda("pr_referencia") & Space(7), 7) _
'                         & Space(1) & Left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
'                         & Space(16) & Left$(Format(Trim(rsItensVenda("pr_classefiscal")), ">") _
'                         & Space(15), 15) _
'                         & "" & Space(3) & Left$(Trim(rsItensVenda("pr_unidade")) & Space(2), 2) _
'                         & Right$(Space(6) & Format(rsItensVenda("QTDE"), "##0"), 6) _
'                         & Right$(Space(12) & Format(rsItensVenda("vlunit"), "#####0.00"), 14) _
'                         & Right$(Space(15) & Format(rsItensVenda("VlTotItem"), "#####0.00"), 15) & Space(1) _
'                         & Right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)
'
'            End If
'                   Printer.Print wStr16
'
'                      If rsItensVenda("DetalheImpressao") = "D" Then
'                         wConta = wConta + 1
'                         rsItensVenda.MoveNext
'                      ElseIf rsItensVenda("DetalheImpressao") = "C" Then
'                         Do While wConta < 28
'                            wConta = wConta + 1
'                            Printer.Print ""
'                         Loop
'                         rsItensVenda.MoveNext
'
'                         wStr13 = Space(78) & "CX 0" & wNumeroCaixa & Space(3) & "Lj " & RsDados("LojaOrigem") & Space(3) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
'                         Printer.Print wStr13
'
'                         wConta = 0
'                         wPagina = wPagina + 1
'
' '------------------------------------------------------------------------------
'                 'Acerto emissao de nota com mais de um formulario
'                       ' Printer.EndDoc
'
'                         Printer.Print ""
'                         Printer.Print ""
'                         Printer.Print ""
'                         Printer.Print ""
'
'                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
'                         If wControlaQuebraDaPagina = 3 Then
'                            Printer.Print ""
'                            wControlaQuebraDaPagina = 0
'                         End If
''----------------------------------------------------------------------------------
'                         Cabecalho rsItensVenda("TipoNota")
'                      ElseIf rsItensVenda("DetalheImpressao") = "T" Then
'                         wConta = wConta + 1
'                         rsItensVenda.MoveNext
'                         Call FinalizaNota
'                      Else
'                         wConta = wConta + 1
'                         rsItensVenda.MoveNext
'                      End If
'            Loop
'         Else
'            MsgBox "Produto não encontrado", vbInformation, "Aviso"
'         End If
'    Else
'        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
'    End If
'rsItensVenda.Close
'RsDados.Close
'
'
'End Function
Private Sub ImprimeCarimbo()
                 
                       SQL = ""
'                       SQL = "Select CNF_Carimbo,CNF_DetalheImpressao,CNF_TipoCarimbo from CarimboNotaFiscal where " & _
'                             "CNF_Nf = " & RsDados("nf") & " and CNF_Serie = '" & RsDados("Serie") & "' and CNF_Loja = '" & RsDados("Lojaorigem") & "'" & _
'                             "order by cnf_tipocarimbo desc, cnf_sequencia asc"
                       SQL = "Select CNF_Carimbo,CNF_DetalheImpressao,CNF_TipoCarimbo from CarimboNotaFiscal where " & _
                             "CNF_Numeroped = '" & RsDados("numeroped") & "'" & _
                             "order by cnf_tipocarimbo desc, cnf_sequencia asc"
                       
                       RsPegaItensEspeciais.CursorLocation = adUseClient
                       RsPegaItensEspeciais.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

                       If Not RsPegaItensEspeciais.EOF Then
                            Printer.Print ""
                            Do While Not RsPegaItensEspeciais.EOF
                                If Trim(RsPegaItensEspeciais("CNF_tipocarimbo")) = "Z" Then
                                 wStr16 = right$(Space(116) & Trim(RsPegaItensEspeciais("CNF_Carimbo")), 116)
                                Else
                                 wStr16 = Space(5) & left$(RsPegaItensEspeciais("CNF_Carimbo") & Space(116), 116)
                                End If
                                 Printer.Print wStr16

                                 If RsPegaItensEspeciais("CNF_DetalheImpressao") = "D" Then
                                     wConta = wConta + 1

                                 ElseIf RsPegaItensEspeciais("CNF_DetalheImpressao") = "C" Then
                                     
                                     Do While wConta < 34
                                       wConta = wConta + 1
                                       Printer.Print ""
                                     Loop

                                     wConta = 0

                         
                                     wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                                     If wControlaQuebraDaPagina = 3 Then
                                        Printer.Print ""
                                        wControlaQuebraDaPagina = 0
                                     End If

                                     Cabecalho RsDados("tiponota")
                                     Printer.Print ""
                                ElseIf RsPegaItensEspeciais("CNF_DetalheImpressao") = "T" Then
                                       wConta = wConta + 1
                                       Printer.Print ""
                                       Call FinalizaNota(Trim(frmPedido.txtPedido.Text))
                                Else
                                       wConta = wConta + 1
                                End If
                                RsPegaItensEspeciais.MoveNext
                            Loop

                             RsPegaItensEspeciais.Close
'                             Call FinalizaNota(wPedido)
                             Exit Sub
                         Else
                             RsPegaItensEspeciais.Close
                             Call FinalizaNota(Trim(frmPedido.txtPedido.Text))
                         End If

End Sub

Public Function DadosLoja()

    Dim SQL As String
    SQL = "Select CTS_Loja,Loja.* from loja,Controlesistema where lo_loja= CTS_Loja"

    rsInfLoja.CursorLocation = adUseClient
    rsInfLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not rsInfLoja.EOF Then

       wRazao = Trim(rsInfLoja("LO_Razao"))
       Wendereco = Trim(rsInfLoja("lo_ENDERECO")) & ", " & Trim(rsInfLoja("lo_numero"))
       wbairro = rsInfLoja("lo_bairro")
       wCGC = rsInfLoja("lo_CGC")
       wIest = rsInfLoja("lo_INSCRICAOESTADUAL")
       WMunicipio = rsInfLoja("lo_MUNICIPIO")
       westado = rsInfLoja("lo_UF")
       WCep = rsInfLoja("lo_CEP")
       WFone = rsInfLoja("lo_TELEFONE")
       wDDDLoja = rsInfLoja("LO_DDD")
       WFax = rsInfLoja("lo_Fax")
       wLoja = rsInfLoja("CTS_Loja")
       GLB_Loja = rsInfLoja("CTS_Loja")
       wNovaRazao = IIf(IsNull(rsInfLoja("lo_Razao")), "0", rsInfLoja("lo_Razao"))
    
    End If
    rsInfLoja.Close

End Function
Function Cabecalho(ByVal TipoNota As String)
        
    Dim wCgcCliente As String
    Dim impri As Long
    Dim Linha(15) As String
    Dim ContLinha As Integer
    Dim ContParcela As Integer
    
    impri = Printer.Orientation
    wPagina = wPagina + 1
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
    Linha(1) = "          "
    Linha(2) = "          "
    Linha(3) = "          "
    Linha(4) = "          "
    Linha(5) = "          "
    Linha(6) = "          "
    Linha(7) = "          "
    Linha(8) = "          "
    Linha(9) = "          "
    Linha(10) = "          "
    Linha(11) = "          "
    Linha(12) = "          "
    Linha(13) = "          "
    Linha(14) = "          "
    Linha(15) = "          "
    ContLinha = 1
    
    wCondicao = "            "
    Wav = "          "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = ""
    wentrada = "        "
    
    wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
    wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))

    WNatureza = "TRANSFERENCIA"

    If Trim(wLojaVenda) > 0 Then
        If Trim(wLojaVenda) <> Trim(RsDados("LojaOrigem")) Then
            wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
        Else
            wStr6 = ""
        End If
    Else
        wStr6 = ""
    End If
    If Trim(RsDados("AV")) > 1 Then
        If Mid(wCondicao, 1, 9) = "Faturada " Then
            Wav = "AV            : " & Trim(RsDados("AV"))
        End If
    End If
    
     wCondicao = "            "

    
    Linha(ContLinha) = "Pedido " & RsDados("NUMEROPED") & "  Ven " & RsDados("VENDEDOR")
    ContLinha = ContLinha + 1
             
    SQL = "select mo_descricao,mc_valor,mo_grupo from movimentocaixa,modalidade " & _
          "where mc_grupo = mo_grupo and mc_documento = " & RsDados("nf") & " and mc_Serie ='" & RsDados("serie") & _
          "' and mc_loja = '" & Trim(RsDados("lojaorigem")) & "' and mc_grupo like '10%'"

    rdoModalidade.CursorLocation = adUseClient
    rdoModalidade.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
         
    If Not rdoModalidade.EOF Then
        Do While Not rdoModalidade.EOF
         
          If rdoModalidade("mo_grupo") = 10501 Then
            
               SQL = "Select cp_condicao,cp_intervaloParcelas,cp_parcelas from CondicaoPagamento " _
                    & "where  CP_Codigo =" & RsDados("CondPag")

               rdoConPag.CursorLocation = adUseClient
               rdoConPag.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
               wValorParcela = Format((RsDados("totalnota") - RsDados("pgentra")) / rdoConPag("cp_parcelas"), "###,##0.00")
               ContParcela = 1
               wMid = 1
               Linha(ContLinha) = "Faturada " & rdoConPag("cp_parcelas") & " Parc    " & wValorParcela
               ContLinha = ContLinha + 1
               
               Do While Len(rdoConPag("cp_intervaloParcelas")) > wMid
               
                 If rdoConPag("cp_Parcelas") = 1 Then
                     Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy")
                     wMid = wMid + 3
                 ElseIf rdoConPag("cp_Parcelas") Mod 2 = 0 Then
                       Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy") _
                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "dd/mm/yyyy")
                       wMid = wMid + 6
                 Else
                       If Len(rdoConPag("cp_intervaloParcelas")) - 3 > wMid Then
                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy") _
                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "dd/mm/yyyy")
                           wMid = wMid + 6
                       Else
                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy")
                           wMid = wMid + 3
                       End If
                 End If
                 ContLinha = ContLinha + 1
              Loop
              rdoConPag.Close
           Else
              Linha(ContLinha) = rdoModalidade("mo_descricao") & ":   " & Format(rdoModalidade("mc_valor"), "0.00")
              ContLinha = ContLinha + 1
           End If
           rdoModalidade.MoveNext
        Loop
    End If
rdoModalidade.Close

    If RsDados("Pgentra") <> 0 Then
       wentrada = Format(RsDados("Pgentra"), "#####0.00")
       Linha(ContLinha) = "Entrada : " & Format(wentrada, "0.00")
       ContLinha = ContLinha + 1
    End If
    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
       Linha(ContLinha) = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
       ContLinha = ContLinha + 1
    End If
   
    If wPagina = 1 Then
        wCGC = right(String(14, "0") & wCGC, 14)
        wCGC = Format(Mid(wCGC, 1, Len(wCGC) - 6), "###,###,###") & "/" & Mid(wCGC, Len(wCGC) - 5, Len(wCGC) - 10) & "-" & Mid(wCGC, 13, Len(wCGC))
        wCGC = right(String(18, "0") & wCGC, 18)
    End If
  '  wStr0 = Space(110) & wPagina & "/" & RsDados("PAGINANF")  'Inicio Impressão
    wStr0 = Space(110) & "1" & "/" & "1"  'Inicio Impressão

    Printer.Print wStr0

    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 6
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 6
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 6#
    
    If wNovaRazao <> "0" Then
        wStr1 = Space(64) & wNovaRazao
        Printer.Print wStr1
    Else
        Printer.Print ""
    End If
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    wStr1 = (left$(Linha(1) & Space(27), 27)) & Space(10) & left(Format(Trim(UCase(Wendereco)), "<") & Space(34), 34) & left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(25) & right(Format(RsDados("nf"), "######"), 7)
    Printer.Print UCase(wStr1)
    wStr2 = (left$(Linha(2) & Space(27), 27)) & Space(10) & left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & left$(Trim(westado), 2)
    Printer.Print UCase(wStr2)
    wStr3 = (left$(Linha(3) & Space(27), 27)) & Space(10) & "(" & wDDDLoja & ")" & left$(Trim(Format(WFone, "####-####")), 9) & "/(" & wDDDLoja & ")" & left$(Format(WFax, "####-####"), 9) & Space(5) & left$(Format((WCep), "00000-000"), 9)
    Printer.Print UCase(wStr3)
    If wSerie = "CT" Then
        wStr4 = (left$(Linha(4) & Space(27), 27))
    Else
        wStr4 = (left$(Linha(4) & Space(27), 27)) & Space(60) & left(Trim(Format(wCGC, "###,###,###")), 19)
    End If
    Printer.Print wStr4
     wStr4 = (left$(Linha(5) & Space(27), 27))
    
     Printer.Print UCase(wStr4)
     wStr5 = (left$(Linha(6) & Space(32), 32)) & left(Trim(WNatureza) & Space(25), 25) & left$(RsDados("codOper"), 10) & Space(28) & left$(Trim(Format((wIest), "###,###,###,###")), 15)

    Printer.Print wStr5
    wStr5 = (left$(Linha(7) & Space(27), 27))
    Printer.Print wStr5

        wCgcCliente = right(String(14, "0") & Trim(RsDados("LO_cgc")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = right(String(18, "0") & Trim(wCgcCliente), 18)

    
    Printer.Print ""

    wStr6 = (left$(Linha(8) & Space(27), 27)) & Space(5) & left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(RsDados("lo_razao")), ">") & Space(45), 45) & left$(Trim(wCgcCliente) & Space(24), 24) & left$(Format(RsDados("Dataemi"), "dd/mm/yy") & Space(12), 12)
    Printer.Print UCase(wStr6)
    
    wStr6 = (left$(Linha(9) & Space(27), 27))
    Printer.Print UCase(wStr6)
    wStr7 = (left$(Linha(10) & Space(27), 27)) & Space(5) & left$(Format(Trim(RsDados("lo_endereco") & ", " & RsDados("lo_numero")), ">") & Space(42), 42) & left$(Format(Trim(RsDados("lo_bairro")), ">") & Space(18), 18) & right$(Space(12) & Format(RsDados("lo_cep"), "#####-###"), 12) '& Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)


    Printer.Print UCase(wStr7)
    wStr7 = (left$(Linha(11) & Space(27), 27))
    Printer.Print UCase(wStr7)

    wStr8 = (left$(Linha(12) & Space(27), 27)) & Space(5) & left$(Format(Trim(RsDados("lo_municipio")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(RsDados("lo_telefone"))) & Space(15), 15) & left$(Trim(RsDados("lo_UF")), 2) & Space(5) & left$(Trim(Format(RsDados("lo_inscricaoEstadual"), "###,###,###,###")), 15)
    Printer.Print UCase(wStr8)
    Printer.Print ""

    If rdoConPag.State = 1 Then
        rdoConPag.Close
    End If

End Function



'Function Cabecalho(ByVal TipoNota As String)
'    Dim wCgcCliente As String
'    Dim impri As Long
'    Dim Linha(15) As String
'    Dim ContLinha As Integer
'    Dim ContParcela As Integer
'
'    impri = Printer.Orientation
'
'    Printer.ScaleMode = vbMillimeters
'    Printer.ForeColor = "0"
'    Printer.FontSize = 8
'    Printer.FontName = "draft 20cpi"
'    Printer.FontSize = 8
'    Printer.FontBold = False
'    Printer.DrawWidth = 3
'
'    Printer.FontName = "COURIER NEW"
'    Printer.FontSize = 8#
'
'    Linha(1) = "          "
'    Linha(2) = "          "
'    Linha(3) = "          "
'    Linha(4) = "          "
'    Linha(5) = "          "
'    Linha(6) = "          "
'    Linha(7) = "          "
'    Linha(8) = "          "
'    Linha(9) = "          "
'    Linha(10) = "          "
'    Linha(11) = "          "
'    Linha(12) = "          "
'    Linha(13) = "          "
'    Linha(14) = "          "
'    Linha(15) = "          "
'    ContLinha = 1
'
'    wCondicao = "            "
'    Wav = "          "
'    wStr20 = ""
'    wStr19 = "               "
'    wStr7 = ""
'    wentrada = "        "
'
'    wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
'    wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))
'
'
'    If UCase(TipoNota) = "T" Then
'        WNatureza = "TRANSFERENCIA"
'    ElseIf UCase(TipoNota) = "V" Then
'        WNatureza = "VENDA"
'    ElseIf UCase(TipoNota) = "E" Then
'        WNatureza = "DEVOLUCAO"
'    ElseIf UCase(TipoNota) = "S" And (RsDados("CFOAUX") = "5949" Or RsDados("CFOAUX") = "6949") Then
'        WNatureza = "OUTRAS OPER Ñ ESPEC."
'    End If
'
'
'    If Trim(wLojaVenda) > 0 Then
'        If Trim(wLojaVenda) <> Trim(RsDados("LojaOrigem")) Then
'            wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
'        Else
'            wStr6 = ""
'        End If
'    Else
'        wStr6 = ""
'    End If
'    If Trim(RsDados("AV")) > 1 Then
'        If Mid(wCondicao, 1, 9) = "Faturada " Then
'            Wav = "AV            : " & Trim(RsDados("AV"))
'        End If
'    End If
'
'    If Trim(WNatureza) = "TRANSFERENCIA" Then
'        wCondicao = "            "
'    ElseIf Trim(WNatureza) = "DEVOLUCAO" Then
'        wCondicao = "            "
'    End If
'
'    Linha(ContLinha) = "Pedido: " & RsDados("NUMEROPED") & "  Vendedor: " & RsDados("VENDEDOR")
'    ContLinha = ContLinha + 1
'
'    SQL = "select mo_descricao,mc_valor,mo_grupo from movimentocaixa,modalidade " & _
'          "where mc_grupo = mo_grupo and mc_documento = " & RsDados("nf") & " and mc_Serie ='" & RsDados("serie") & _
'          "' and mc_loja = '" & Trim(RsDados("lojaorigem")) & "' and mc_grupo like '10%'"
'
'    rdoModalidade.CursorLocation = adUseClient
'    rdoModalidade.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'    If Not rdoModalidade.EOF Then
'        Do While Not rdoModalidade.EOF
'
'          If rdoModalidade("mo_grupo") = "10501" Then
'
'               SQL = "Select cp_condicao,cp_intervaloParcelas,cp_parcelas from CondicaoPagamento " _
'                    & "where  CP_Codigo =" & RsDados("CondPag")
'
'               rdoConPag.CursorLocation = adUseClient
'               rdoConPag.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'               wValorParcela = Format((RsDados("totalnota") - RsDados("pgentra")) / rdoConPag("cp_parcelas"), "###,##0.00")
'               ContParcela = 1
'               wMid = 1
'               Linha(ContLinha) = "Faturada " & rdoConPag("cp_parcelas") & " Parc    " & wValorParcela
'               ContLinha = ContLinha + 1
'
'               Do While Len(rdoConPag("cp_intervaloParcelas")) > wMid
'
'                 If rdoConPag("cp_Parcelas") = 1 Then
'                     Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "yyyy/mm/dd")
'                     wMid = wMid + 3
'                 ElseIf rdoConPag("cp_Parcelas") Mod 2 = 0 Then
'                       Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "yyyy/mm/dd") _
'                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "yyyy/mm/dd")
'                       wMid = wMid + 6
'                 Else
'                       If Len(rdoConPag("cp_intervaloParcelas")) - 3 > wMid Then
'                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "yyyy/mm/dd") _
'                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "yyyy/mm/dd")
'                           wMid = wMid + 6
'                       Else
'                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "yyyy/mm/dd")
'                           wMid = wMid + 3
'                       End If
'                 End If
'                 ContLinha = ContLinha + 1
'              Loop
'              rdoConPag.Close
'           Else
'              Linha(ContLinha) = rdoModalidade("mo_descricao") & ":   " & Format(rdoModalidade("mo_valor"), "0.00")
'              ContLinha = ContLinha + 1
'           End If
'           rdoModalidade.MoveNext
'        Loop
'    End If
'rdoModalidade.Close
'
'
'    If RsDados("Pgentra") <> 0 Then
'       wentrada = Format(RsDados("Pgentra"), "#####0.00")
' '      wStr18 = "Entrada : " & Format(wentrada, "0.00")
'       Linha(ContLinha) = "Entrada : " & Format(wentrada, "0.00")
'       ContLinha = ContLinha + 1
'    End If
'    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
'  '      wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
'       Linha(ContLinha) = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
'       ContLinha = ContLinha + 1
'    End If
'
'
'    If wPagina = 1 Then
'        wCGC = Right(String(14, "0") & wCGC, 14)
'        wCGC = Format(Mid(wCGC, 1, Len(wCGC) - 6), "###,###,###") & "/" & Mid(wCGC, Len(wCGC) - 5, Len(wCGC) - 10) & "-" & Mid(wCGC, 13, Len(wCGC))
'        wCGC = Right(String(18, "0") & wCGC, 18)
'    End If
'    wStr0 = Space(105) & wPagina & "/" & RsDados("PAGINANF")  'Inicio Impressão
'    Printer.Print wStr0
'    Printer.Print ""
'
'    Printer.ScaleMode = vbMillimeters
'    Printer.ForeColor = "0"
'    Printer.FontSize = 6
'    Printer.FontName = "draft 20cpi"
'    Printer.FontSize = 6
'    Printer.FontBold = False
'    Printer.DrawWidth = 3
'    Printer.FontName = "COURIER NEW"
'    Printer.FontSize = 6#
'
'    If wNovaRazao <> "0" Then
'        wStr1 = Space(64) & wNovaRazao
'        Printer.Print wStr1
'        Printer.Print ""
'    Else
'        Printer.Print ""
'    End If
'    Printer.ScaleMode = vbMillimeters
'    Printer.ForeColor = "0"
'    Printer.FontSize = 8
'    Printer.FontName = "draft 20cpi"
'    Printer.FontSize = 8
'    Printer.FontBold = False
'    Printer.DrawWidth = 3
'    Printer.FontName = "COURIER NEW"
'    Printer.FontSize = 8#
'
'    If Glb_NfDevolucao = True Then
'        WNatureza = "DEVOLUCAO"
'        wStr1 = Space(2) & (Linha(1)) & Space(34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(15) & "X" & Space(16) & Left(Format(RsDados("nf"), "######"), 7)
'    Else
'        wStr1 = Space(2) & (Linha(1)) & Space(34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & Left(Format(RsDados("nf"), "######"), 7)
'    End If
'    Printer.Print wStr1
'    wStr2 = Space(2) & (Linha(2)) & Space(34) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
'    Printer.Print wStr2
'    If wSerie = "CT" Then
'        wStr3 = Space(2) & (Linha(3)) & Space(34) & Space(29) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "####-##'"), 9)
'    Else
'        wStr3 = Space(2) & (Linha(3)) & Space(34) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "####-###"), 9)
'    End If
'    Printer.Print wStr3
'    If wSerie = "CT" Then
'        wStr4 = Space(2) & Linha(4)
'    Else
'        wStr4 = Space(2) & Linha(4) & Space(40) & Space(46) & Left(Trim(Format(wCGC, "###,###,###")), 19)
'    End If
'    Printer.Print wStr4
'    wStr4 = Space(2) & (Linha(5)) & Space(40)
'    Printer.Print wStr4
'
'    If wSerie = "CT" Then
'        If Trim(WNatureza) = "TRANSFERENCIA" Then
'            wStr5 = Space(2) & Linha(6) & Space(36) & Format(Trim(WNatureza), ">") & Space(18) & Left$(RsDados("codOper"), 10)
'        End If
'    Else
'
'        If Trim(Wav) <> "" Then
'            wStr5 = Space(2) & Linha(6) & Space(30) & Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(27) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((wIest), "###,###,###,###")), 15)
'        Else
'         '   wStr5 = Space(2) & Left(Format(linha(8)) & Space(30), 30) & Space(2) & Left(Trim(WNatureza) & Space(26), 26) & Left$(RsDados("CFOP"), 10) & Space(28) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
'        wStr5 = Space(2) & Linha(6) & Space(30) & Space(2) & Left(Trim(WNatureza) & Space(26), 26)
'
'        End If
'    End If
'    Printer.Print wStr5
'
'    wStr5 = Space(2) & (Linha(7))
'    Printer.Print wStr7
'    wStr5 = Space(2) & (Linha(8))
'    Printer.Print wStr7
'
'        wCgcCliente = Right(String(14, "0") & Trim(RsDados("lo_cgc")), 14)
'        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
'        wCgcCliente = Right(String(18, "0") & Trim(wCgcCliente), 18)
'
'    If wSerie = "CT" Then
'        'If wStr6 <> "" Then
'        '    wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("ce_razao")), ">") & Space(50), 50) & Space(6) & Left$(Format(RsDados("Dataemi"), "yyyy/mm/dd"), 12)
'        'Else
'        wStr6 = Space(2) & (Linha(9)) & Space(29) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("lo_Razao")), ">") & Space(45), 45) & Left$(Format(RsDados("Dataemi"), "yyyy/mm/dd"), 12)
''        'End If
'    Else
'        wStr6 = Space(2) & (Linha(9)) & Space(29) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("lo_razao")), ">") & Space(45), 45) & Left$(Trim(wCgcCliente) & Space(24), 24) & Space(1) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy") & Space(12), 12)
'    End If
'
'    Printer.Print wStr6
'
'    wStr6 = Space(2) & (Linha(10))
'    Printer.Print wStr6
'
'    If RsDados("EmiteDataSaida") = "S" Then
'        If wSerie = "CT" Then
'            wStr7 = Space(2) & Linha(10) & Space(29) & Left$(Format(Trim(RsDados("lo_endereco")), ">") & Space(42), 42) & Space(14) & Left$(Format(RsDados("Dataemi"), "yyyy/mm/dd"), 12)
'        Else
'            wStr7 = Space(2) & Linha(10) & Space(29) & Left$(Format(Trim(RsDados("lo_endereco")), ">") & Space(42), 42) & Left$(Format(Trim(RsDados("lo_bairro")), ">") & Space(21), 21) & Right$(Space(11) & RsDados("lo_cep"), 11) & Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)
'        End If
'    Else
'        If wSerie = "CT" Then
'            wStr7 = Space(2) & Linha(10) & Space(29) & Left$(Format(Trim(RsDados("lo_endereco")), ">") & Space(42), 42) '& Space(14) & Left$(Format(RsDados("Dataemi"), "yyyy/mm/dd"), 12)
'        Else
'            wStr7 = Space(2) & Linha(10) & Space(29) & Left$(Format(Trim(RsDados("lo_endereco")), ">") & Space(42), 42) & Left$(Format(Trim(RsDados("lo_bairro")), ">") & Space(21), 21) & Right$(Space(11) & RsDados("lo_cep"), 11) '& Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)
'        End If
'    End If
'    Printer.Print wStr7
'    wStr7 = Space(2) & Linha(11)
'    Printer.Print wStr7
'
'    If wSerie = "CT" Then
'        wStr8 = ""
'    Else
'        wStr8 = Space(2) & Linha(12) & Space(29) & Left$(Format(Trim(RsDados("lo_municipio")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("lo_telefone"))) & Space(15), 15) & Left$(Trim(RsDados("lo_uf")), 2) & Space(5) & Left$(Trim(Format(RsDados("lo_inscricaoEstadual"), "###,###,###,###")), 15)
'    End If
'    Printer.Print ""
'    Printer.Print wStr8
'
'    Printer.Print ""
'    Printer.Print ""
'
'    If rdoConPag.State = 1 Then
'        rdoConPag.Close
'    End If
'
'End Function
'
'
'


'Private Sub FinalizaNota()
'
'
''''''********************************************************************************************
'                Do While wConta < 7
'                   wConta = wConta + 1
'                   Printer.Print ""
'                Loop
'
''''NÃO EXCLUIR SERÁ USADO FUTARAMENTE 28/09/2011 - Isnara
'
''''                SQL = ""
''''                SQL = "Select * from carimbosEspeciais,CarimboNotaFiscal where CE_referencia = CNF_Carimbo And " & _
''''                      "CNF_TipoCarimbo = 'S' And CNF_NumeroPed = " & frmPedido.txtPedido.Text
''''
''''                RsPegaItensEspeciais.CursorLocation = adUseClient
''''                RsPegaItensEspeciais.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''''
''''                If Not RsPegaItensEspeciais.EOF Then
''''                    Do While Not RsPegaItensEspeciais.EOF
''''
''''                      If RsPegaItensEspeciais("CE_Linha12") <> "" Then
''''                         Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha12"), 90)
''''                      End If
''''
''''                      If RsPegaItensEspeciais("CE_Linha1") <> "" Then
''''                        wConta = wConta + 7
''''                        If Trim(RsPegaItensEspeciais("CE_Linha5")) = "" Then
''''                             Printer.Print Space(7) & "______________________________________________________________"
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
''''                             Printer.Print Space(9) & "___________________________________     ____/____/______   "
''''                             Printer.Print Space(9) & "            Assinatura                        Data         "
''''                        Else
''''                             Printer.Print Space(7) & "______________________________________________________________"
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
''''                             Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha5"), 60)
''''                             Printer.Print Space(9) & "___________________________________     ____/____/______   "
''''                             Printer.Print Space(9) & "            Assinatura                        Data         "
''''                        End If
''''                      End If
''''                      RsPegaItensEspeciais.MoveNext
''''                    Loop
''''                End If
''''                RsPegaItensEspeciais.Close
''''
'
'               SQL = ""
'               SQL = "Select * from CarimboNotaFiscal where " & _
'                     "CNF_NumeroPed = " & frmPedido.txtPedido.Text
'
'
Private Sub FinalizaNota(wPedido As String)
     If wNotaTransferencia = False Then
   
        Do While wConta < 13
        wConta = wConta + 1
        Printer.Print ""
        Loop
       
     End If


        wStr9 = right$(Space(9) & Format(RsDados("BaseICMS"), "######0.00"), 9) & right$(Space(25) & Format(RsDados("VLRICMS"), "######0.00"), 12) & Space(34) & right$(Space(10) & Format(RsDados("VlrMercadoria"), "######0.00"), 10)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = right(Space(9) & Format(Space(9) & RsDados("FreteCobr"), "######0.00"), 9) & Space(46) & right(Space(10) & Format(RsDados("TotalNota"), "######0.00"), 10)
        Printer.Print wStr10

     
     wStr11 = Space(2) & "                          "
     Printer.Print wStr11
     wStr12 = Space(2) & "                                                     "
     Printer.Print wStr12
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = right$(Space(5) & Format(RsDados("Volume"), "######0.00"), 5) & Space(5) & "Volume(s)" & Space(25) & right$(Space(7) & Format(RsDados("PesoBR"), "######0.00"), 7) & Space(5) & right$(Space(7) & Format(RsDados("PesoBR"), "######0.00"), 7)
     Printer.Print wStr13
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = Space(105) & right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
     Printer.Print wStr13
     Printer.Print ""
'     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.EndDoc
     

End Sub
    
Function AcharICMSInterEstadual(ByVal Referencia As String, ByVal ChaveIcms As Double) As Boolean

    wIE_icmsAplicado = 0
    wIE_Tributacao = 0
    wIE_Cfo = 0
    wIE_BasedeReducao = 0
    wIE_icmsdestino = 0
    
    SQL = "SELECT * from IcmsInterEstadual where IE_Codigo = " & ChaveIcms
       
    RsICMSInter.CursorLocation = adUseClient
    RsICMSInter.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
           
    If RsICMSInter.EOF Then
        AcharICMSInterEstadual = False
        RsICMSInter.Close
        Exit Function
    Else
        AcharICMSInterEstadual = True
    End If
    
    wIE_icmsAplicado = RsICMSInter("IE_icmsAplicado")
    wIE_Tributacao = RsICMSInter("IE_CST")
    wIE_Cfo = RsICMSInter("IE_Cfop")
    wIE_BasedeReducao = RsICMSInter("IE_BasedeReducao")
    wIE_icmsdestino = RsICMSInter("IE_icmsdestino")
    RsICMSInter.Close

'    SQL = "SELECT * From IcmsInterEstadual Where IE_Codigo = " & ChaveIcms
'    RsICMSInter.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'    If RsICMSInter.EOF Then
'        AcharICMSInterEstadual = False
'        MsgBox "ICMS inter estadual da referencia " & Referencia & " não encontrado" & Chr(10) & "A nota não pode ser impressa", vbCritical, "Aviso"
'        RsICMSInter.Close
'        Exit Function
'    Else
'        AcharICMSInterEstadual = True
'    End If
    
        
End Function

Sub AjustaTela(ByRef Formulario As Form)

  Formulario.top = frmPedido.PicBanner.top
  Formulario.left = frmPedido.PicBanner.left
  Formulario.Width = frmPedido.PicBanner.Width
  Formulario.Height = frmPedido.PicBanner.Height


End Sub

Public Sub LimpaForm()
Dim cObjeto As Control
Dim wColuna As Integer
  
  For Each cObjeto In frmPedido.Controls
      If (TypeOf cObjeto Is TextBox) Then
        cObjeto.Text = ""
      End If
  Next
  
  For Each cObjeto In frmPedido.Controls
      If (TypeOf cObjeto Is CommandButton) Then
        cObjeto.Enabled = True
      End If
  Next
  
  For wColuna = 0 To frmPedido.grdDadosProduto.Cols - 1
    frmPedido.grdDadosProduto.TextMatrix(1, wColuna) = ""
  Next wColuna
  
  
'  For wColuna = 0 To frmPedido.grdPrecos.Cols - 1
'     frmPedido.grdPrecos.TextMatrix(1, wColuna) = ""
'  Next wColuna
  
  frmPedido.grdItensProduto.Rows = 1
  'frmPedido.grdPrecos.Rows = 1
'  frmPedido.fraMenu.Visible = False
  frmPedido.cmdFechaPedido.Visible = False
  frmPedido.cmdBotoes(2).Visible = False
  frmPedido.cmdBotoes(12).Visible = False
  frmPedido.cmdBotoes(6).Visible = False
  frmPedido.cmdBotoes(9).Visible = False
  frmPedido.cmdBotoes(8).Visible = False
  frmPedido.cmdBotoes(10).Visible = False
  frmPedido.cmdTR.Visible = False
  frmPedido.cmdBotoes(7).Visible = False
    
  frmPedido.txtPesquisar.Enabled = False
  frmPedido.txtQuantidade.Enabled = False
  frmPedido.grdItensProduto.Enabled = False
  'frmPedido.grdPrecos.Enabled = False
  'frmPedido.grdPrecos.TextMatrix(0, 0) = "A Vista"
  frmPedido.grdDadosProduto.Enabled = False
  frmPedido.wbFichaTecnica.Visible = False
              
  frmPedido.cmdQtdeItens.Caption = Format(0, "0") '+ "    "
  frmPedido.cmdTotalPedido.Caption = Format(0, "0.00") '+ "           "
  

  frmPedido.txtCondicaoFaturado.Text = ""
  frmPedido.mskDatafaturado.Text = "__/__/____"
  wGuardaPagamento = ""
  auxItens = 0
  wClienteTelaAdicionais = False
              
  'frmPedido.PicBanner.Visible = True
  frmPedido.cmbPedido.Visible = True
  frmPedido.cmdBotoes(1).Visible = True
  frmPedido.cmdBotoes(4).Visible = True
  
  frmPedido.PicBanner.Visible = False
  
'  frmPedido.fraMenu.Visible = False
'  frmPedido.fraMenu.Enabled = False
  'picQuadroGeral.Width = 9975

  frmPedido.txtPedido.Enabled = True
  frmPedido.fradados.Enabled = True
  frmPedido.txtVendedor.Width = frmPedido.txtPedido.Width
  frmPedido.fradados.Width = 2830
  frmPedido.txtVendedor.Enabled = False
  frmPedido.txtPedido.SetFocus
  'frmPedido.PicBanner.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerChamada\BannerChamada")
  GBL_Frete = 0
  auxItens = 0
  auxQtdeItens = 0
  
  
  frmPedido.grdDadosProduto.BackColor = &HE0E0E0
  frmPedido.grdDadosProduto.ForeColor = vbBlack


  wLiberaBloqueioPreco = False
       'frmPedido.WebBrowser1.SetFocus

End Sub
'Ficha Financeira
Sub VerificaSituacaoOnLine()

    SQL = ""
    SQL = "Select CT_BancosOnLine from Controle"
        Set rsLoja = adoCNLoja.OpenResultset(SQL)
    GLB_BancosOnLine = rsLoja("CT_BancosOnLine")

    rsLoja.Close

End Sub
'
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'


Function PesquisaCliente(ByVal tipoPesquisa As Integer, ByVal Cliente As String, ByRef NomerdoResultset) As Boolean
Set NomerdoResultset = New ADODB.Recordset


'
'--------------------------------Pesquisa Pelo Codigo do Cliente (1)-------------------------
'
    DescricaoOperacao "Pesquisando Cliente"
    If tipoPesquisa = 1 Then
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from FIN_Cliente" _
            & "where CE_CodigoCliente = " & Cliente & " "
             NomerdoResultset.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

'
'-------------------------------Pesquisa por cgc ou cpf (2) ---------------------------------
'
    ElseIf tipoPesquisa = 2 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from FIN_Cliente" _
            & "where CE_Cgc = '" & Cliente & "' "
             NomerdoResultset.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
'
'-------------------------------Pesquisa Pelo Nome Cliente (3) ---------------------------------
'
    ElseIf tipoPesquisa = 3 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_razao,CE_CodigoCliente from FIN_Cliente" _
            & "where CE_Razao like '" & UCase(Cliente) & "%' order by CE_Razao"
            NomerdoResultset.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
'
'-------------------------------Pesquisa Cliente Tela frmCadCliente(4) --------------------------
'
    ElseIf tipoPesquisa = 4 Then
        SQL = ""
        SQL = ""
        SQL = "Select * from FIN_Cliente" _
            & "where CE_CodigoCliente = " & Cliente & " order by CE_CodigoCliente"
            'Set NomerdoResultset = adoCNLoja.OpenResultset(SQL)
            NomerdoResultset.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Else
        Exit Function
    End If
    If Not NomerdoResultset.EOF Then
        PesquisaCliente = True
    Else
        PesquisaCliente = False
    End If
    DescricaoOperacao "Pronto"
    
End Function

Sub LimpaTR()

    SQL = ""
    SQL = "update nfcapa set ValorTotalCodigoZero = 0, TotalNotaAlternativa = 0, valormercadoriaAlternativa = 0 " & _
          "where Numeroped = " & frmPedido.txtPedido.Text
   adoCNLoja.Execute (SQL)
 
    
    SQL = ""
    SQL = "update nfitens set ReferenciaAlternativa = 0, DescricaoAlternativa = '', valormercadoriaAlternativa = 0 " & _
          "where Numeroped = " & frmPedido.txtPedido.Text
   adoCNLoja.Execute (SQL)

End Sub


Function PegaSerieNota() As String

    
       SQL = ""
       SQL = "Select CTS_SerieNota from ControleSistema"
     
       rdoSerie.CursorLocation = adUseClient
       rdoSerie.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

       If Not rdoSerie.EOF Then
           PegaSerieNota = rdoSerie("CTS_SerieNota")
       End If
       rdoSerie.Close

End Function



'Sub TrocaBannerTopo1()
'
'    If NroBanner = 1 Then
'       frmPedido.webInternet2.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerTopo1\BannerTopo1b.swf")
'       NroBanner = 2
'    ElseIf NroBanner = 2 Then
'       frmPedido.webInternet2.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerTopo1\BannerTopo1c.swf")
'       NroBanner = 3
'    ElseIf NroBanner = 3 Then
'       frmPedido.webInternet2.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerTopo1\BannerTopo1d.swf")
'       NroBanner = 4
'    ElseIf NroBanner = 4 Then
'       frmPedido.webInternet2.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerTopo1\BannerTopo1a.swf")
'       NroBanner = 1
'    End If
'
'    EsperarTrocaBanner 7.11
'
'End Sub

Sub EsperarTrocaBanner(ByVal Tempo As Long)
    
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
        DoEvents
    Loop
 
End Sub

Public Function ImprimirCotacaoBola(ByVal Pedido As Double)

Dim wNomeVendedor As String
Dim wCondPag As String
Dim wGuardaPagamento As String

Open GLB_ImpCotacao For Output As #1
 
    Screen.MousePointer = 11
   
    ValorlItem = 0
    ValorDesconto = 0
    SubTotal = 0

   Print #1, "________________________________________"
   Print #1, "COTACAO DE VENDA    "
   Print #1, " "
   Print #1, "NUMERO: " & Pedido & " DATA: " & Format(Date, "dd/mm/yyyy")
   Print #1, "________________________________________"

   SQL = "Select NFCapa.*,Loja.*,fin_Cliente.* From NFCapa, Loja, fin_Cliente " _
          & " Where Cliente = CE_CodigoCliente and LojaOrigem = LO_Loja and NumeroPed = " & Pedido

    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
   Print #1, ; RsDados("LO_Razao")
   Print #1, ; RsDados("LO_Endereco") & ", " & RsDados("LO_numero")
   Print #1, ; RsDados("LO_CEP") & " - " & RsDados("LO_Bairro")
   Print #1, ; RsDados("LO_Municipio") & " - " & RsDados("LO_UF")
   Print #1, ; "TELEFONE: "; RsDados("LO_DDD") & RsDados("LO_Telefone")
   Print #1, "========================================"
   Print #1, " "
   Print #1, "DESCRICAO DO PRODUTO                    "
   Print #1, "CODIGO  PRODUTO  QTDxUNIT.   VALOR TOTAL"
   Print #1, "________________________________________"
   
            
             
   If Not RsDados.EOF Then
      wPegaDesconto = RsDados("Desconto")
      wPegaFrete = RsDados("FreteCobr")
     
      If RsDados("CondPag") = "85" Then
            wGuardaPagamento = " - " & frmPedido.mskDatafaturado.Text
      ElseIf RsDados("CondPag") > 2 And RsDados("CondPag") < 100 Then
            wGuardaPagamento = " - " & frmPedido.txtCondicaoFaturado.Text
      ElseIf RsDados("CondPag") > 99 Then
            wGuardaPagamento = " - " & frmTrocaModalidadeVenda.grdPrecos.TextMatrix(grdPrecos.Row, 0)
      Else
            wGuardaPagamento = " "
      End If
      
      wCondPag = RTrim(LTrim(IIf(IsNull(RsDados.Fields("modalidadevenda")), "A VISTA", RsDados.Fields("modalidadevenda"))))
      
   End If
             
   SQL = "Select * from Nfitens " _
        & "Where numeroPed = " & Pedido

   RsDadosItens.CursorLocation = adUseClient
   RsDadosItens.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
 
       
       If Not RsDadosItens.EOF Then
          Do While Not RsDadosItens.EOF
             SQL = "Select PR_Descricao from Produtoloja Where PR_Referencia ='" & RsDadosItens("Referencia") & "'"
             rdoProduto.CursorLocation = adUseClient
             rdoProduto.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

             ValorlItem = (RsDadosItens("vlunit") * RsDadosItens("Qtde"))
             SubTotal = (SubTotal + ValorlItem)
           
             Print #1, Trim(rdoProduto("PR_Descricao"))
             Print #1, RsDadosItens("referencia") _
             & Space(3) & right(Space(4) & Format(RsDadosItens("Qtde"), "###0"), 4) & "x" _
             & Format(RsDadosItens("vlunit"), "###,###,###.00") & Space(5) _
             & right(Space(10) & Format(ValorlItem, "###,###,###.00"), 14)
             rdoProduto.Close
             RsDadosItens.MoveNext
          Loop
       End If
       
       RsDadosItens.Close
        
       SQL = "Select CTS_ValidadeCotacao From ControleSistema"
       rdoValidadeCotacao.CursorLocation = adUseClient
       rdoValidadeCotacao.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
       Print #1, " "
       Print #1, "SUB-TOTAL " & Space(16) & right(Space(10) & Format(SubTotal, "###,###,###,##0.00"), 14)
       Print #1, "DESCONTO  " & Space(16) & right(Space(10) & Format(RsDados("Desconto"), "###,###,###,##0.00"), 14)
       Print #1, "TOTAL     " & Space(16) & right(Space(10) & Format((SubTotal - RsDados("Desconto")), "###,###,##0.00"), 14)

       Print #1, "________________________________________"
       Print #1, "COND PAGTO " & wCondPag & wGuardaPagamento
       Print #1, "VALIDADE "; Format(DateAdd("D", rdoValidadeCotacao("CTS_ValidadeCotacao"), Date), "dd/mm/yyyy")
       Print #1, "VENDEDOR " & Trim(frmPedido.txtVendedor.Text)
       Print #1, " "
       Print #1, "========================================"
       Print #1, Trim(RsDados("ce_razao"))
       Print #1, Trim(RsDados("Ce_Endereco")) & ", " & Trim(RsDados("ce_numero"))
       Print #1, Trim(RsDados("ce_Cep")) & " - " & Trim(RsDados("ce_bairro"))
       Print #1, Trim(RsDados("ce_municipio")) & " - " & Trim(RsDados("ce_estado"))
       Print #1, "TELEFONE: " & Trim(RsDados("ce_telefone"))
       Print #1, "========================================"
       Print #1, " "
       Print #1, " "
       Print #1, " "
       Print #1, " "
       Print #1, " "
       Print #1, " "
       Print #1, " "
       Print #1, " "
       rdoValidadeCotacao.Close
       RsDados.Close
       Printer.EndDoc
       Close #1
       Screen.MousePointer = 0




End Function

Public Sub sairDoSistema()
    'sairDoSistema
    'Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
    'Call criaIconeBarra(TrayDelete, frmPedido.Hwnd, frmPedido.Caption, frmPedido.Icon)
    End
End Sub


Public Sub campoSelecionadoComCaracter(campo As TextBox)
    If campo.Text <> "" Then
        campo.SelStart = 0
        campo.SelLength = Len(campo.Text)
    End If
End Sub


Function encerraTransferencia2(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String) As Boolean
    Dim SerieProd As String
            
        wVerificaTM = False
        wQuantdadeTotalItem = 0
        wAnexo = ""
        wAnexo1 = ""
        wAnexo2 = ""
        wQuantItensCapaNF = 0
        wCFO2 = " "
        wCFO1 = " "
        wChaveICMS = 0
        GLB_TotalIcmsCalculado = 0
        GLB_ValorCalculadoICMS = 0
        GLB_BasedeCalculoICMS = 0
        GLB_AliquotaAplicadaICMS = 0
        GLB_AliquotaICMS = 0
        GLB_BaseTotalICMS = 0
        GLB_Tributacao = 0
        wCFOItem = 0
        wUltimoItem = 1
        wComissaoVenda = 0
        wSomaVenda = 0
        wSomaMargem = 0
        wCarimbo5 = ""
        wCarimbo2 = ""
        wST20 = "N"
        wST60 = "N"
        EncerraVenda = True
        SerieProd = ""
        wRecebeCarimboAnexo = ""
        wQuantItensNF = 0
        'If ConsistenciaNota(NumeroDocumento, SerieDocumento) = False Then
            'EncerraVenda = False
            'Exit Function
        'End If


'If rsCapaNF.State = 1 Then
  'rsCapaNF.Close
'End If

Dim RsCapaNF As New ADODB.Recordset
Dim rsItensNF As New ADODB.Recordset

SQL = "Select nfcapa.*, fin_Estado.*,fin_Cliente.* from nfcapa, fin_Estado, fin_cliente where nfcapa.numeroped = " & _
       NumeroDocumento & " and nfcapa.cliente = fin_cliente.ce_codigocliente " & _
      "And fin_cliente.ce_estado = fin_Estado.UF_Estado"
             
             RsCapaNF.CursorLocation = adUseClient
             RsCapaNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        SQL = "Select produtoloja.*, nfitens.* from produtoloja,nfitens " _
              & "where nfitens.numeroped = " & NumeroDocumento & "" _
              & " and pr_referencia = nfitens.referencia order by NfItens.Item"
          
              rsItensNF.CursorLocation = adUseClient
              rsItensNF.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

          
          
          wLoja = RsCapaNF("lojaorigem")
          wTM = RsCapaNF("TM")

'aqui
          
          If Not rsItensNF.EOF Then
             Do While Not rsItensNF.EOF
                  wChaveICMSItem = wChaveICMS
                  If Trim(wCarimbo5) = "" Then
                    If rsItensNF("PR_substituicaotributaria") = "N" _
                       And rsItensNF("PR_codigoreducaoicms") > 0 Then
                        wST20 = "S"
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
                                Exit Function
                          End If
                    End If
                    
                    
                                      
                        wCFOItem = wIE_Cfo
                        GLB_AliquotaAplicadaICMS = wIE_icmsAplicado
                        GLB_Tributacao = wIE_Tributacao
                        GLB_CFOP = wIE_Cfo
                        wAnexoIten = rsItensNF("PR_CodigoReducaoICMS")
                        
                        If wAnexoIten <> 0 Then
                            If wAnexoIten = 1 Then
                                wAnexo1 = rsItensNF("Item") & "," & wAnexo1
                            ElseIf wAnexoIten = 2 Then
                                wAnexo2 = rsItensNF("Item") & "," & wAnexo2
                            End If
                        End If
                        
                        
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
                       
                            wAnexoAux = ""
                            If rsItensNF("pr_codigoreducaoicms") <> 0 Then
                               wAnexoAux = wAnexoAux & "," & Format(rsItensNF("ITEM"), "0")
                            End If
                        
                            If wCFOItem = 5102 Or wCFOItem = 6102 Then
                                wCFO1 = wCFOItem
                            ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                                wCFO2 = wCFOItem
                                If Trim(wCFO2) = 6405 Then
                                   wCFO2 = 6404
                                End If
                            End If
                        
                        If Trim(wCFO1) = "" And Trim(wCFO2) = "" And RsCapaNF("TipoNota") <> "S" Then
                            wCFO1 = wCFOItem
                     
                        End If
                   
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
'aqui
                    wQuantItensCapaNF = RsCapaNF("QtdItem")
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
     
                    If wRomaneio = True Then
                       GLB_BasedeCalculoICMS = 0
                       GLB_ValorCalculadoICMS = 0
                    End If

                    rdoCNLoja.BeginTrans
                                        
                    SQL = "UPDATE nfitens set baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
                    & "Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & " ,TipoNota = 'V'," _
                    & "DetalheImpressao = '" & wDetalheImpressao & "', CSTICMS = " & GLB_Tributacao & ", " _
                    & "CFOP = " & GLB_CFOP & ", ICMSAplicado = " & ConverteVirgula(wIE_icmsdestino) _
                    & " where nfitens.numeroped = " & NumeroDocumento _
                    & " and Referencia = '" & rsItensNF("PR_Referencia") & "' and Item=" & rsItensNF("Item") & ""
                    rdoCNLoja.Execute (SQL)
                
                    If Err.Number = 0 Then
                        rdoCNLoja.CommitTrans
                    Else
                        rdoCNLoja.RollbackTrans
                    End If
                    
                rsItensNF.MoveNext
                End If
             Loop
     ' End If 'estava com '
        If wRomaneio = True Then
           wRomaneio = False
        End If
        
        rsItensNF.Close
 
   
' -------------------------------------- INSERIR CARIMBOS --------------------------------------------------

           


            rdoCNLoja.BeginTrans

            SQL = ""
            SQL = "update CarimboNotafiscal set CNF_Serie = '" & RsCapaNF("Serie") & "', CNF_NF = " & RsCapaNF("nf") & _
                  " , CNF_Situacaoprocesso = 'A' , CNF_DataProcesso = '" & Format(Date, "yyyy/mm/dd") & "' " & _
                  " where CNF_NumeroPed = " & NumeroDocumento
              rdoCNLoja.Execute (SQL)
  
            If Err.Number = 0 Then
                 rdoCNLoja.CommitTrans
            Else
                 rdoCNLoja.RollbackTrans
            End If


            If RsCapaNF("desconto") > 0 Then
            
               rdoCNLoja.BeginTrans
                            
               SQL = ""
               SQL = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                      " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                      ",1,' DESCONTO:    " & Format(RsCapaNF("desconto"), "#####0.00") & _
                      "' , 'Z',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (SQL)

                If Err.Number = 0 Then
                  rdoCNLoja.CommitTrans
                Else
                 rdoCNLoja.RollbackTrans
                End If
            End If
            
            If wST20 = "S" Then
            
                wSequenciaS = wSequenciaS + 1
                  
                SQL = ""
                SQL = "Select CE_linha12 from CarimbosEspeciais where ce_Referencia = '9999991'"
                rsCarimbo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                  
                 rdoCNLoja.BeginTrans
                  
                SQL = ""
                SQL = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                  " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                  "," & wSequenciaS & " ,'" & rsCarimbo("CE_linha12") & "' , 'S',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (SQL)
                rsCarimbo.Close
            
                If Err.Number = 0 Then
                    rdoCNLoja.CommitTrans
                Else
                    rdoCNLoja.RollbackTrans
                End If
                
            End If
            
            If wST60 = "S" Then
            
                wSequenciaS = wSequenciaS + 1
                            
                SQL = ""
                SQL = "Select CE_linha12 from CarimbosEspeciais where ce_Referencia = '9999992' "
                rsCarimbo.CursorLocation = adUseClient
                rsCarimbo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
                rdoCNLoja.BeginTrans
            
                SQL = ""
                SQL = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                      " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                      "," & wSequenciaS & " ,'" & rsCarimbo("CE_linha12") & "' , 'S',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (SQL)
            
                If Err.Number = 0 Then
                    rdoCNLoja.CommitTrans
                Else
                    rdoCNLoja.RollbackTrans
                End If
            
                rsCarimbo.Close
                
                End If
                    
           Else
               MsgBox "Não foi possível acessar os carimbos fiscais", vbCritical, "AVISO"
               rsItensNF.Close
               RsCapaNF.Close
               Exit Function

          End If
          
            SQL = ""
            SQL = "select count(*) as somacarimbo from carimbonotafiscal where cnf_Loja = '" & RsCapaNF("Lojaorigem") & "' and CNF_NF = " & RsCapaNF("nf") & _
                         " and cnf_serie = '" & RsCapaNF("serie") & "' " & _
                         " and CNF_NumeroPed = " & RsCapaNF("numeroped") & " "
            rsCarimbo.CursorLocation = adUseClient
            rsCarimbo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            If Not rsCarimbo.EOF Then
               wTotalCarimbo = rsCarimbo("somacarimbo")
            End If
            rsCarimbo.Close
            
          
            SQL = ""
            SQL = "select * from carimbonotafiscal where cnf_Loja = '" & RsCapaNF("Lojaorigem") & "' and CNF_NF = " & RsCapaNF("nf") & _
                  " and cnf_serie = '" & RsCapaNF("serie") & "' " & _
                  " and CNF_NumeroPed = " & RsCapaNF("numeroped") & _
                  " order by cnf_tipocarimbo desc, cnf_sequencia asc"
            rsCarimbo.CursorLocation = adUseClient
            rsCarimbo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

            If Not rsCarimbo.EOF Then
                
                wRestoItens = ((RsCapaNF("QTDItem")) Mod 12)
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

                   rdoCNLoja.BeginTrans
                   
                   SQL = ""
                   SQL = "update CarimboNotafiscal set CNF_DetalheImpressao = '" & wDetalheImpressao & "', cnf_data = '" & Format(RsCapaNF("dataemi"), "yyyy/mm/dd") & "'" & _
                         " where cnf_Loja = '" & rsCarimbo("cnf_Loja") & "' and cnf_nf = " & rsCarimbo("cnf_nf") & _
                         " and cnf_serie = '" & rsCarimbo("cnf_serie") & "' and cnf_tipocarimbo = '" & rsCarimbo("cnf_tipocarimbo") & "' " & _
                         " and cnf_sequencia = '" & rsCarimbo("cnf_sequencia") & "' " & _
                         " and CNF_NumeroPed = " & RsCapaNF("numeroped") & " "
                         rdoCNLoja.Execute (SQL)
                         
                   If Err.Number = 0 Then
                        rdoCNLoja.CommitTrans
                   Else
                        rdoCNLoja.RollbackTrans
                   End If
                   
                   rsCarimbo.MoveNext
                Loop
             End If
             rsCarimbo.Close
             
'-------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
             
             SQL = "update nfitens set BaseICMS = 0 where BaseICMS is null and numeroped = " & NumeroDocumento
             rdoCNLoja.Execute (SQL)
             
             SQL = "Select sum(BASEICMS) as BaseICMS from nfitens where numeroped = " & NumeroDocumento
             rsCarimbo.CursorLocation = adUseClient
             rsCarimbo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
             SQL = "select top 1 CFOP from nfitens where numeroped = " & NumeroDocumento
             rsItensNF.CursorLocation = adUseClient
             rsItensNF.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             rdoCNLoja.BeginTrans

             SQL = "UPDATE nfcapa set " _
                & "Paginanf = " & ConverteVirgula(wUltimoItem) & ",BaseICMS = " _
                & ConverteVirgula(rsCarimbo("BaseICMS")) & ", " _
                & "ECF  = " & GLB_ECF & ",TipoNota = 'V' , CodOper = " & rsItensNF("CFOP") & ", " _
                & "CFOAUX = " & rsItensNF("CFOP") _
                & "where nfcapa.numeroped = " & NumeroDocumento & ""
                rdoCNLoja.Execute (SQL)
                
             If Err.Number = 0 Then
                rdoCNLoja.CommitTrans
             Else
                rdoCNLoja.RollbackTrans
             End If
             
             rsItensNF.Close
             rsCarimbo.Close
                
'--------------------------------------  ATUALIZA ESTOQUE LOJA ----------------------------------------------------
             
            rdoCNLoja.BeginTrans
             
            SQL = ""
            SQL = "UPDATE EstoqueLoja Set EL_Estoque = (EL_Estoque - QTDE) FROM NFItens, EstoqueLoja " _
                 & "Where EL_Referencia = Referencia and NumeroPed = " & NumeroDocumento
             
            rdoCNLoja.Execute SQL
            
            If Err.Number = 0 Then
               rdoCNLoja.CommitTrans
            Else
               rdoCNLoja.RollbackTrans
            End If
        
Exit Function
ErroEncerraTransferencia:
     MsgBox Err.Number & " - " & Err.description & vbLf & _
           "Não foi possível encerrar a nota fiscal de venda.", vbCritical, "AVISO"

    rsItensNF.Close
    RsCapaNF.Close

    Exit Function


    rsItensNF.Close
    RsCapaNF.Close
'End Sub
            
    
End Function

Public Sub verificaAppExecucao()
    If App.PrevInstance Then
       MsgBox App.EXEName + " Já está executando", vbCritical
       End
    End If
End Sub


'------------------------------------------------------------------------------------EMERSON

Private Sub criaTermo()
    
    Dim CotacaoHTL As String
    Dim CotacaoDado As String
    Dim teste, teste1 As Integer
   teste = 1
    Dim campo As String
    
    CotacaoHTL = codigoCotacaohtml
    
    CotacaoDado = "Loja"

    campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, loja(teste))
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
         teste = teste + 1
    Loop
    
    CotacaoDado = "CliVen"

    campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, loja(teste))
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
        teste = teste + 1
    Loop
     CotacaoDado = "CodVen"

    campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, loja(teste))
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
        teste = teste + 1
    Loop
    
     
     CotacaoDado = "Valo"
    campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, loja(teste))
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
      teste = teste + 1
    Loop
    
    CotacaoDado = "QtoI"
    campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, total)
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Loop
         
     CotacaoDado = "TPVal"
    campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, loja(teste))
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
      teste = teste + 1
    Loop
    
             CotacaoDado = "ImagemVe"

    campo = campoCotacao2(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, vendedor)
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
        teste = teste + 1
    Loop
            CotacaoDado = "ImagemLo"

    campo = campoCotacao2(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, imagemLogo)
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
        teste = teste + 1
    Loop
    CotacaoDado = "TabelaTe"
       campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Do While campo <> ""
        CotacaoHTL = Replace(CotacaoHTL, CotacaoDado & campo, tabelaHtml)
        campo = campoCotacao(CotacaoHTL, CotacaoDado)
    Loop
    
    
    Open "C:\Sistemas\DMAC Venda\Cotacao\Cotacao.html " For Output As #1
    Print #1, CotacaoHTL
    Close #1
    
    FrmCotacao.WebNavegador.Navigate "C:\Sistemas\DMAC Venda\Cotacao\Cotacao.html"
    
End Sub
'-----------------------------------------------------------------------------------------
Public Function Replace(Source As String, Find As String, ReplaceStr As String, _
    Optional ByVal Start As Long = 1, Optional Count As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
    Dim findLen As Long
    Dim replaceLen As Long
    Dim Index As Long
    Dim counter As Long
    
    findLen = Len(Find)
    replaceLen = Len(ReplaceStr)
    If findLen = 0 Then Err.Raise 5
    
    If Start < 1 Then Start = 1
    Index = Start
    
    Replace = Source
    
    Do
        Index = InStr(Index, Replace, Find, Compare)
        If Index = 0 Or Count = 0 Then Exit Do
        If findLen = replaceLen Then
            Mid$(Replace, Index, findLen) = ReplaceStr
        Else
            Replace = left$(Replace, Index - 1) & ReplaceStr & Mid$(Replace, _
            Index + findLen)
        End If
        Index = Index + replaceLen
        counter = counter + 1
    Loop Until counter = Count
    
    If Start > 1 Then Replace = Mid$(Replace, Start)
    
End Function
'--------------------------------------------------------------------------------------
Private Function campoCotacao(codigoCotacao As String, campo As String) As String
    
    If codigoCotacao Like "*" & campo & "*" Then
        Dim inicioCampo, fimCampo As Integer
        
        inicioCampo = (InStr(codigoCotacao, campo)) + (Len(campo))
        fimCampo = (InStr(inicioCampo, codigoCotacao, "<")) - inicioCampo
        
        If inicioCampo + fimCampo <> 0 Then
            campoCotacao = Mid$(codigoCotacao, inicioCampo, fimCampo)
        End If
        
    Else
        campoCotacao = ""
    End If
    
End Function
Private Function campoCotacao2(codigoCotacao As String, campo As String) As String
    
    If codigoCotacao Like "*" & campo & "*" Then
        Dim inicioCampo, fimCampo As Integer
        
        inicioCampo = (InStr(codigoCotacao, campo)) + (Len(campo))
        fimCampo = (InStr(inicioCampo, codigoCotacao, """")) - inicioCampo
        
        If inicioCampo + fimCampo <> 0 Then
            campoCotacao2 = Mid$(codigoCotacao, inicioCampo, fimCampo)
        End If
        
    Else
        campoCotacao2 = ""
    End If
    
End Function

'--------------------------------------------------------------------------------------
Public Sub obterCertificado()
    Dim fso As New FileSystemObject
    Dim mensagemArquivoTXT As TextStream
  
    Set mensagemArquivoTXT = fso.OpenTextFile _
    ("C:\Sistemas\DMAC Venda\Cotacao\configCotacao")
    codigoCotacaohtml = mensagemArquivoTXT.ReadAll
    mensagemArquivoTXT.Close
    criaTermo

End Sub


Public Sub dados(ByVal Pedido As Double)

    Dim valorICMS As Double
    Dim baseICMS As Double
    Dim SQL As String
    Dim cnpj As String
    Dim Linha As Integer
     Dim pagina As Integer
     Dim fso As New FileSystemObject

   
    SQL = "Select * from Loja where LO_Loja='" & AchaLojaControle & "'"
    rsInfLoja.CursorLocation = adUseClient
    rsInfLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    i = 1
    If Not rsInfLoja.EOF Then
    ReDim Preserve loja(i) As String
                loja(i) = rsInfLoja("LO_Razao")
                i = i + 1
     ReDim Preserve loja(i) As String
     cnpj = rsInfLoja("LO_cgc")
       loja(i) = Format(Mid(cnpj, 1, 8), "##,###,###") & "/" & Mid(cnpj, 9, 4) & "-" & Mid(cnpj, 13, 2)
      i = i + 1
    ReDim Preserve loja(i) As String
               loja(i) = rsInfLoja("LO_Endereco") & ", " & rsInfLoja("LO_Numero")
               i = i + 1
    ReDim Preserve loja(i) As String
                loja(i) = Format(rsInfLoja("LO_Cep"), "#####-###")
               i = i + 1
    ReDim Preserve loja(i) As String
                loja(i) = rsInfLoja("LO_Municipio")
               i = i + 1
    ReDim Preserve loja(i) As String
                loja(i) = rsInfLoja("LO_UF")
                 i = i + 1
    ReDim Preserve loja(i) As String
               loja(i) = Format(rsInfLoja("LO_Telefone"), "(##)####-####")
           i = i + 1
           

'               ReDim Preserve loja(i) As String
'               loja(i) = rsInfLoja("lo_televendas")
'           i = i + 1
'                  ReDim Preserve loja(i) As String
'               loja(i) = rsInfLoja("LO_emailoja")
'           i = i + 1
                     ReDim Preserve loja(i) As String
               loja(i) = rsInfLoja("LO_site")
           i = i + 1
    End If
    rsInfLoja.Close
    ReDim Preserve loja(i) As String
                loja(i) = Pedido
                 i = i + 1

        ReDim Preserve loja(i) As String
                loja(i) = Date
                 i = i + 1
     SQL = "SELECT Cliente From NFCapa Where Numeroped = " & Pedido
    
          rsComplemento.CursorLocation = adUseClient
          rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
    If Not rsComplemento.EOF Then
       wValorComplemento = Trim(rsComplemento.Fields("Cliente"))
'    Else
'    ReDim Preserve Loja(i) As String
'       wValorComplemento = 999999
'        Loja(i) = 999999
'       i = i + 1
    End If

    rsComplemento.Close
        sql1 = ""
   
        sql1 = "Select CE_CodigoCliente, CE_Razao, CE_Telefone, CE_Fax from FIN_Cliente Where CE_CodigoCliente= '" & wValorComplemento & "'"
            rsCliente.CursorLocation = adUseClient
            rsCliente.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsCliente.EOF Then
     ReDim Preserve loja(i) As String
        loja(i) = rsCliente("CE_CodigoCliente") & " - " & rsCliente("CE_Razao")
       i = i + 1
        ReDim Preserve loja(i) As String
        loja(i) = Format(rsCliente("CE_Telefone"), "(##)####-####")
       i = i + 1
        ReDim Preserve loja(i) As String
        loja(i) = Format(rsCliente("CE_Fax"), "(##)####-####")
       i = i + 1
        
'    Else
'     ReDim Preserve Loja(i) As String
'    Loja(i) = "999999-CONSUMIDOR"
'       i = i + 1
'        ReDim Preserve Loja(i) As String
'        Loja(i) = "0000-0000"
'       i = i + 1
'        ReDim Preserve Loja(i) As String
'        Loja(i) = "0000-0000"
'       i = i + 1
        
    End If
    
    
    rsCliente.Close
    qtotal = 0
    
    SQL = "EXEC SP_FIN_Calcula_ICMS_NFCAPA '" & frmPedido.txtPedido.Text & "', '" & Format(Date, "YYYY/MM/DD") & "'"
    adoCNLoja.Execute SQL
    
     SQL = "Select FO_NomeFantasia,C.Cliente,C.BASEICMS,C.vlricms,VE_Codigo, VE_Nome,VDE_Email,VDE_ASSINATURA, PR_Descricao, I.Qtde,I.desconto as descontoporitem, I.Vlunit,(I.VLUnit * I.Qtde) as VLUnit2,PR_Referencia, PR_ClasseFiscal, C.Desconto as Desconto " _
         & ",C.FRETECOBR,C.TOTALNOTA,PR_Unidade,PR_ICMSSAIDA,PR_ST " _
        & "From ProdutoLoja, NFItens as I, NFCapa as C, Vende,Vende_Detalhe,Fornecedor " _
        & "Where PR_Referencia = I.Referencia and VE_Codigo = C.Vendedor and I.NumeroPed = C.NumeroPed and " _
        & " I.DataEmi = C.DataEmi and PR_CodigoFornecedor=FO_CodigoFornecedor and  VDE_CODIGO=VE_Codigo and C.NumeroPed = " & Pedido
        
    rdoPedido.CursorLocation = adUseClient
    rdoPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    'Dados do Vendedor
    ReDim Preserve loja(i) As String
        loja(i) = rdoPedido("VDE_Email")
        i = i + 1
     ReDim Preserve loja(i) As String
        loja(i) = rdoPedido("VE_Codigo") & " - " & UCase$(rdoPedido("VE_Nome"))
        i = i + 1

        vendedor = "C:\Sistemas\DMAC Venda\Cotacao\Imagem\" & rdoPedido("VE_Codigo")
If Not fso.FileExists(vendedor) Then
vendedor = ""
End If
    
imagemLogo = "C:\Sistemas\DMAC Venda\Cotacao\Imagem\" & GLB_logoPedido
If Not fso.FileExists(imagemLogo) Then
imagemLogo = ""
End If
        
     
        'Valor do Frete
          ReDim Preserve loja(i) As String
        loja(i) = Format(rdoPedido("FRETECOBR"), "##,###,###0.00")
        i = i + 1
        'Valor Total de Desconto
          ReDim Preserve loja(i) As String
        loja(i) = Format(rdoPedido("Desconto"), "##,###,###0.00")
        i = i + 1
    ReDim Preserve loja(i) As String
        'BASE ICMS
        loja(i) = Format((rdoPedido("BASEICMS") - rdoPedido("Desconto")), "##,###,###0.00")
        i = i + 1
    ReDim Preserve loja(i) As String
        'Valor ICMS
       If IsNull(rdoPedido("vlricms")) Then
       loja(i) = Format(0, "##,###,###0.00")
        i = i + 1
       Else
        loja(i) = Format(rdoPedido("vlricms"), "##,###,###0.00")
        i = i + 1
        End If
  'valor Total da Cotação
    ReDim Preserve loja(i) As String
        loja(i) = Format((rdoPedido("TOTALNOTA") - rdoPedido("Desconto")), "##,###,###0.00")
        valparcela = loja(i)
        i = i + 1
 
   
  
  n = 1
  Linha = 1
  pagina = 1
  cont = 0
  tabelaHtml = criainiciotabela
  Do While Not rdoPedido.EOF
    
    cont = cont + 1
    
        VlTotItem = rdoPedido("VLUnit2") - rdoPedido("descontoporitem")
        VlTotItem = Format(VlTotItem, "##,###,###0.00")
If Linha <= 32 And pagina = 1 Then
    tabelaHtml = tabelaHtml + "<tr class=alinha1>" _
    & "<td align=left>" & cont & "</td>" _
    & "<td align=left>" & rdoPedido("PR_Referencia") & "</td>" _
   & "<td align=left>" & rdoPedido("FO_NomeFantasia") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_Descricao") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_ClasseFiscal") & "</td>" _
   & "<td >" & rdoPedido("Qtde") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_Unidade") & "</td>" _
   & "<td >" & rdoPedido("PR_ICMSSAIDA") & "</td>" _
   & "<td >" & rdoPedido("PR_ST") & "</td>" _
   & "<td >" & Format(rdoPedido("Vlunit"), "##,###,###0.00") & "</td>" _
& "<td >" & Format(rdoPedido("descontoporitem"), "##,###,###0.00") & "</td>" _
& "<td >" & VlTotItem & "</td>" _
   & "</tr>"
ElseIf Linha <= 38 And pagina >= 2 Then
   tabelaHtml = tabelaHtml + "<tr class=alinha1>" _
    & "<td align=left>" & cont & "</td>" _
    & "<td align=left>" & rdoPedido("PR_Referencia") & "</td>" _
   & "<td align=left>" & rdoPedido("FO_NomeFantasia") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_Descricao") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_ClasseFiscal") & "</td>" _
   & "<td >" & rdoPedido("Qtde") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_Unidade") & "</td>" _
   & "<td >" & rdoPedido("PR_ICMSSAIDA") & "</td>" _
   & "<td >" & rdoPedido("PR_ST") & "</td>" _
   & "<td >" & Format(rdoPedido("Vlunit"), "##,###,###0.00") & "</td>" _
& "<td >" & Format(rdoPedido("descontoporitem"), "##,###,###0.00") & "</td>" _
& "<td >" & VlTotItem & "</td>" _
   & "</tr>"
Else
 tabelaHtml = tabelaHtml + "</table>"
    tabelaHtml = tabelaHtml + "<FONT SIZE = 2>ORCAMENTO: " & Pedido & "   - PAGINA: " & pagina & "</FONT><br><br><br><br><br><br><br><br>" + criainiciotabela
       tabelaHtml = tabelaHtml + "<tr class=alinha1>" _
    & "<td align=left>" & cont & "</td>" _
    & "<td align=left>" & rdoPedido("PR_Referencia") & "</td>" _
   & "<td align=left>" & rdoPedido("FO_NomeFantasia") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_Descricao") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_ClasseFiscal") & "</td>" _
   & "<td >" & rdoPedido("Qtde") & "</td>" _
   & "<td align=left>" & rdoPedido("PR_Unidade") & "</td>" _
   & "<td >" & rdoPedido("PR_ICMSSAIDA") & "</td>" _
   & "<td >" & rdoPedido("PR_ST") & "</td>" _
   & "<td >" & Format(rdoPedido("Vlunit"), "##,###,###0.00") & "</td>" _
& "<td >" & Format(rdoPedido("descontoporitem"), "##,###,###0.00") & "</td>" _
& "<td >" & VlTotItem & "</td>" _
   & "</tr>"
Linha = 1
pagina = pagina + 1
End If
Linha = Linha + 1
  rdoPedido.MoveNext

  Loop
 
If Linha <= 32 And pagina = 1 Then
Do While Linha <= 32
tabelaHtml = tabelaHtml + linhabranca
Linha = Linha + 1
Loop
ElseIf Linha <= 40 And pagina >= 2 Then
Do While Linha <= 38
tabelaHtml = tabelaHtml + linhabranca
Linha = Linha + 1
Loop
End If
  tabelaHtml = tabelaHtml + "</table><FONT SIZE = 2>   ORCAMENTO: " & Pedido & "   - PAGINA: " & pagina & "</FONT><BR>"
  qtotal = cont
 
    rdoPedido.Close
  
    SQL = "select modalidadevenda from nfcapa where NumeroPed = " & Pedido
             
             rsComplemento.CursorLocation = adUseClient
             rsComplemento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
         
    SQL = ""
    SQL = "Select CTS_ValidadeCotacao From ControleSistema"
    rdoValidadeCotacao.CursorLocation = adUseClient
    rdoValidadeCotacao.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        ReDim Preserve loja(i) As String
        loja(i) = Format(DateAdd("D", rdoValidadeCotacao("CTs_ValidadeCotacao"), Date), "dd/mm/yyyy")
        i = i + 1

    rdoValidadeCotacao.Close
       
       
        If Not rsComplemento.EOF Then

        ReDim Preserve loja(i) As String
        loja(i) = IIf(IsNull(rsComplemento.Fields("modalidadevenda")), "A VISTA", rsComplemento.Fields("modalidadevenda"))
        
        If Trim(loja(i)) = "Faturado" Then
         loja(i) = loja(i) & " - " & wPagamentototal
         loja(i) = Replace(loja(i), "X", "  ")
         
        Else
        wPagamentototal = valparcela
        valparcela = (wPagamentototal / wPagamento)
        valparcela = Format(valparcela, "##,###,###0.00")
        wPagamentototal = Format(wPagamentototal, "##,###,###0.00")
        If Trim(loja(i)) = "Cartão" Then
        loja(i) = "Cart&atilde;o"
        End If
        
         loja(i) = Trim(loja(i)) & " - " & wPagamento & " X " & valparcela & " Total: " & wPagamentototal
        End If
        i = i + 1
            
        Else
        
        ReDim Preserve loja(i) As String
        loja(i) = "A VISTA"
        loja(i) = loja(i) & " - " & wPagamento
        i = i + 1
        End If
        rsComplemento.Close
total = qtotal

  
    

    
    obterCertificado
    
End Sub

Private Function criainiciotabela() As String

criainiciotabela = "<table width=1100 border=1px  cellspacing=0 cellpadding=1>" _
     & "<tr>" _
     & "<th class=titulo width=64>ITEM</th>" _
     & "<th class=titulo width=68>REFER&Ecirc;CIA</th>" _
     & "<th class=titulo width=68>MARCA</th>" _
     & "<th class=titulo width=313>DESCRI&Ccedil;&Atilde;O</th>" _
     & "<th class=titulo width=63>NCM</th>" _
     & "<th class=titulo width=64>QTD</th>" _
     & "<th class=titulo width=67>UN</th>" _
     & "<th class=titulo width=73>%ICM</th>" _
     & "<th class=titulo width=70>CST</th>" _
     & "<th class=titulo width=83>PRE&Ccedil;O UNITARIO</th>" _
     & "<th class=titulo width=68>DESCONTO</th>" _
     & "<th class=titulo width=73>VALOR TOTAL</th>" _
   & "</tr>"

End Function

Private Function linhabranca() As String

linhabranca = "<tr>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "<td > &nbsp; </td>" _
            & "</tr>"

End Function


Public Function endIMG(nomeBotao As String) As String
    
    Dim Arquivo As String
    Dim enderecoArquivo As String
    
    enderecoArquivo = "c:\sistemas\dmac VENDA\imagens\lojas\" & GLB_logoPedido & "_" & nomeBotao
    Arquivo = Dir(enderecoArquivo, vbDirectory)
    
    If Arquivo = Empty Then
        enderecoArquivo = "c:\sistemas\dmac VENDA\imagens\" & nomeBotao
    End If
    
    endIMG = enderecoArquivo
    
End Function



Public Function validaDadosCliente(codigoCliente As String) As Boolean

    Dim adoValidaCliente As New ADODB.Recordset
    Dim SQL As String
    
 
    If codigoCliente = Empty Then
        MsgBox "Não foi informado o codigo do Cliente!", vbExclamation, "Cliente"
        validaDados = False
        Exit Function
    End If
    
    adoCNLoja.CommitTrans
    
    SQL = "exec SP_GLB_Valida_Cliente '" & codigoCliente & "'"
    adoCNLoja.Execute SQL
    
    SQL = "select count(campoErrado) as campoErrado from temp_Fin_Cliente_Erro"
    adoValidaCliente.CursorLocation = adUseClient
    adoValidaCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If adoValidaCliente("campoErrado") = 0 Then
        validaDadosCliente = True
    Else
        validaDadosCliente = False
    End If

    adoValidaCliente.Close
    adoCNLoja.BeginTrans
    
End Function
