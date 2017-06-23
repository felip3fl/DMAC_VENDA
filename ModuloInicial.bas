Attribute VB_Name = "ModuloInicial"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Main()

verificaAppExecucao

lsDSN = "Driver={Microsoft Access Driver (*.mdb)};" & _
          "Dbq=c:\sistemas\DMACini.mdb;" & _
          "Uid=Admin; Pwd=astap36"
  adoCNAccess.Open lsDSN


  Sql = "Select count(*) as QtdeDeLojasINI from ConexaoSistema"
   
  rdoConexaoINI.CursorLocation = adUseClient
  rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
 
        If Not rdoConexaoINI.EOF Then
            
          Sql = "Select * from ParametroSistema"
          rdoParametroINI.CursorLocation = adUseClient
          rdoParametroINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
           
          If Not rdoParametroINI.EOF Then
          
             GLB_ImpCotacao = rdoParametroINI("GLB_ImpCotacaoResumo")
             GLB_ECF = rdoParametroINI("CXA_ECF")
             GLB_ImpressoraNota = rdoParametroINI("GLB_ImpNotaFiscal")
             'GLB_Impr00 = rdoParametroINI("GLB_Imp00")
             GLB_Cotacao = Trim(rdoParametroINI("BCO_Cotacao"))
             Glb_AlteraResolucao = rdoParametroINI("GLB_AlteraResolucao")

             rdoParametroINI.Close
           Else
             MsgBox "Problemas no banco de dados de inicializacao", vbCritical, "Aviso"
             rdoParametroINI.Close
             rdoConexaoINI.Close
             End
             Exit Sub
           End If
           
           
           If rdoConexaoINI("QtdeDeLojasINI") = 1 Then
           
              rdoConexaoINI.Close
              
              Sql = "Select * from ConexaoSistema"
                     rdoConexaoINI.CursorLocation = adUseClient
                     rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
                     
                     If Not rdoConexaoINI.EOF Then
                    
                        GLB_Servidor = Trim(rdoConexaoINI("GLB_ServidorRetaguarda"))
                        GLB_Loja = Trim(rdoConexaoINI("GLB_Loja"))
                        GLB_Banco = Trim(rdoConexaoINI("GLB_BancoRetaguarda"))
                        GLB_Servidorlocal = Trim(rdoConexaoINI("GLB_ServidorLocal"))
                        Glb_BancoLocal = Trim(rdoConexaoINI("GLB_BancoLocal"))
                        'GLB_Usuario = Trim(rdoConexaoINI("GLB_Usuario"))
                        'GLB_Senha = Trim(rdoConexaoINI("GLB_Senha"))
                        
                        rdoConexaoINI.Close
                     Else
                        MsgBox "Problemas no banco de dados de inicializacao", vbCritical, "Aviso"
                        rdoConexaoINI.Close
                        End
                        Exit Sub
                     End If
                     
           Else
              rdoConexaoINI.Close
              frmInicio.Show
              Exit Sub
           End If
        Else
           MsgBox "Banco de dados de inicializacao Vazio", vbCritical, "Aviso"
           adoCNAccess.Close
           Unload frmInicio
           End
           Exit Sub
        End If

ConectaODBC
adoCNAccess.Close

If GLB_ConectouOK = True Then
'       mdiTraderBalcao.Show
       
          'frmPedido.Show
          ShellExecute hwnd, "open", ("C:\Sistemas\DMAC Venda\limpaCache"), "", "", sw_hide
          frmTrocaVersao.Show
          frmBandeja.Show
          'On Error Resume Next
          'tmrTroca.Interval = 1
          
      Else
        MsgBox "Erro ao conectar-se ao Banco de Dados da Loja", vbCritical, "Atenção"
        Exit Sub
 End If
End Sub

'Public Function ConectaODBC(ByRef RdoVar, ByVal Banco As String, ByVal Servidor As String, ByVal Usuario As String, ByVal Senha As String) As Boolean
  Sub ConectaODBC()

'  =========  Conexao  ADO com SQL Server 2000 ========
On Error GoTo ConexaoErro:

'adoCNLoja.Provider = "SQLOLEDB"
'adoCNLoja.Properties("Data Source").Value = GLB_Servidorlocal
'adoCNLoja.Properties("Initial Catalog").Value = Glb_BancoLocal
'adoCNLoja.Properties("User ID").Value = GLB_Usuario
'adoCNLoja.Properties("Password").Value = GLB_Senha

    If ConexaoDLLAdo.abrirConexaoADO(adoCNLoja, GLB_Servidorlocal, Glb_BancoLocal) Then
        GLB_ConectouOK = True
        Exit Sub
    End If

'adoCNLoja.Open

'GLB_ConectouOK = True

'Exit Sub
ConexaoErro:
MsgBox "Erro ao abrir banco de localizacao! "

    GLB_ConectouOK = False
  
   
Exit Sub
  
End Sub

Sub ConectaODBCMatriz()
 
 On Error GoTo ConexaoErro

'  =========  Conexao ADO com SQL Server 2000 ========
If rdoCNMatriz.State <> 1 Then
'rdoCNMatriz.Provider = "SQLOLEDB"
'rdoCNMatriz.Properties("Data Source").Value = GLB_Servidor
'rdoCNMatriz.Properties("Initial Catalog").Value = GLB_Banco
'rdoCNMatriz.Properties("User ID").Value = GLB_Usuario
'rdoCNMatriz.Properties("Password").Value = GLB_Senha

    If ConexaoDLLAdo.abrirConexaoADO(rdoCNMatriz, GLB_Servidor, GLB_Banco) Then
        GLB_ConectouOK = True
        'Exit Sub
    End If

    'rdoCNMatriz.Open
    
End If

GLB_ConectouOK = True
Exit Sub

ConexaoErro:
    MsgBox "Erro ao conectar-se ao Banco de Dados da Matriz", vbCritical, "Atenção"
    GLB_ConectouOK = False
  
End Sub



