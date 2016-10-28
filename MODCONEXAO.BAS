Attribute VB_Name = "ModConexao"
Option Explicit
Global TipoProcesso As Integer
Global Servidor As String
Global Banco As String
Dim SQL As String

'Function ConectaODBCRetaguarda(ByRef RdoVar, ByVal Usuario As String, ByVal Senha As String) As Boolean
        
'        On Error GoTo ConexaoErro
    
'        With RdoVar
'            Servidor = "DemeoNovo"
'            Banco = "Demeo"
'
'            .Connect = "Driver={SQL Server};" _
'                    & "Server=" & Trim(Servidor) & ";" _
'                    & "DataBase=" & Trim(Banco) & ";" _
'                    & "MaxBufferSize=512;" _
'                    & "PageTimeout=5;" _
'                    & "UID=" & Usuario & ";" _
'                    & "PWD=" & Senha & ";"
'
'            .LoginTimeout = 10
'            .CursorDriver = rdUseClientBatch
'            .EstablishConnection rdDriverNoPrompt
'        End With
'
'        ConectaODBCRetaguarda = True
'        Exit Function
'

'ConexaoErro:

'    ConectaODBCRetaguarda = False

'End Function

'Function ConectaOdbcBalcao(ByRef RdoVar, ByVal Usuario As String, ByVal Senha As String) As Boolean
    
'        On Error GoTo ConexaoErro
'
'        With RdoVar
'            Servidor = "Demeoserv"
'            Banco = "Loja048"
'
'            .Connect = "Dsn=" & Trim(Servidor) & ";" _
'                    & "Server=" & Trim(Servidor) & ";" _
'                    & "DataBase=" & Trim(Banco) & ";" _
'                    & "MaxBufferSize=512;" _
'                    & "PageTimeout=5;" _
'                    & "UID=" & Usuario & ";" _
'                    & "PWD=" & Senha & ";"
'
'            .LoginTimeout = 10
'            .CursorDriver = rdUseClientBatch
'            .EstablishConnection rdDriverNoPrompt
'        End With
'
'        ConectaOdbcBalcao = True
'        Exit Function
'
'ConexaoErro:
'    ConectaOdbcBalcao = False

'End Function
Function Skin()
End Function
