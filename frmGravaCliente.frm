VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGravaCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Grava Cliente"
   ClientHeight    =   1620
   ClientLeft      =   4065
   ClientTop       =   2325
   ClientWidth     =   2625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   45
      OleObjectBlob   =   "frmGravaCliente.frx":0000
      Top             =   1485
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      TabIndex        =   3
      Top             =   -75
      Width           =   2610
      Begin MSMask.MaskEdBox mskCodCliente 
         Height          =   285
         Left            =   1395
         TabIndex        =   0
         Top             =   420
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16711680
         MaxLength       =   6
         Format          =   "0"
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCodCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Cliente:"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   450
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   2610
      Begin VB.CommandButton cmdRetorna 
         Caption         =   "&Retorna"
         Height          =   390
         Left            =   1815
         TabIndex        =   2
         Top             =   180
         Width           =   750
      End
      Begin VB.CommandButton cmdGrava 
         Caption         =   "&Gravar"
         Height          =   390
         Left            =   1065
         TabIndex        =   1
         Top             =   180
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmGravaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim RSPegaCliente As rdoResultset
'Dim VerificaCliente As rdoResultset

Private Sub cmdGrava_Click()

    SQL = ""
    SQL = "Select * From Cliente Where CE_CodigoCliente = " & Val(mskCodCliente) & ""
    RSPegaCliente.CursorLocation = adUseClient
    RSPegaCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    'Set RSPegaCliente = Conexao.OpenResultset(SQL)

    On Error Resume Next

    If Not RSPegaCliente.EOF Then
        
        SQL = ""
        SQL = "Select * From Cliente Where CE_CodigoCliente =  " & Val(mskCodCliente) & ""
          VerificaCliente.CursorLocation = adUseClient
          VerificaCliente.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       ' Set VerificaCliente = rdoCNLoja.OpenResultset(SQL)
        rdoCNLoja.BeginTrans
        If VerificaCliente.EOF Then
            SQL = ""
            SQL = "Insert into Cliente ([CE_CodigoCliente], [CE_CGC], [CE_InscricaoEstadual], [CE_Razao], " _
                & "[CE_Endereco], [CE_Bairro], [CE_Municipio], [CE_Estado], [CE_CEP], [CE_Telefone], [CE_Fax], " _
                & "[CE_EMail], [CE_TipoPessoa], [CE_Praca], [CE_PagamentoCarteira], [CE_EnderecoCobranca], " _
                & "[CE_BairroCobranca], [CE_MunicipioCobranca], [CE_EstadoCobranca], [CE_CEPCobranca], [CE_LimiteCredito], " _
                & "[CE_DataLimiteCredito], [CE_MaiorCompra], [CE_DataMaiorCompra], [CE_UltimaCompra], [CE_DataUltimaCompra], " _
                & "[CE_UltimoPagamento], [CE_DataUltimoPagamento], [CE_MaiorAtraso], [CE_QuantidadeCompras], [CE_JurosCartorio], [CE_DataCadastro], [CE_DataCancelamento], [CE_Alteracao], [CE_Situacao], [CE_HoraManutencao])" _
                & "Values(" & RSPegaCliente("CE_CodigoCliente") & ", '" & RSPegaCliente("CE_CGC") & "', '" & RSPegaCliente("CE_InscricaoEstadual") & "', '" & RSPegaCliente("CE_Razao") & "', " _
                & "'" & RSPegaCliente("CE_Endereco") & "', '" & RSPegaCliente("CE_Bairro") & "', '" & RSPegaCliente("CE_Municipio") & "', '" & RSPegaCliente("CE_Estado") & "', '" & RSPegaCliente("CE_CEP") & "', '" & RSPegaCliente("CE_Telefone") & "', '" & RSPegaCliente("CE_Fax") & "', " _
                & "'" & RSPegaCliente("CE_EMail") & "', '" & RSPegaCliente("CE_TipoPessoa") & "', " & RSPegaCliente("CE_Praca") & ", '" & RSPegaCliente("CE_PagamentoCarteira") & "', '" & RSPegaCliente("CE_EnderecoCobranca") & "', " _
                & "'" & RSPegaCliente("CE_BairroCobranca") & "', '" & RSPegaCliente("CE_MunicipioCobranca") & "', '" & RSPegaCliente("CE_EstadoCobranca") & "', '" & RSPegaCliente("CE_CEPCobranca") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_LimiteCredito"), "0.00")) & ", " _
                & "'" & Format(RSPegaCliente("CE_DataLimiteCredito"), "mm/dd/yyyy") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_MaiorCompra"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataMaiorCompra"), "mm/dd/yyyy") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_UltimaCompra"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataUltimaCompra"), "mm/dd/yyyy") & "', " _
                & "" & ConverteVirgula(Format(RSPegaCliente("CE_UltimoPagamento"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataUltimoPagamento"), "mm/dd/yyyy") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_MaiorAtraso"), "0.00")) & ", " & RSPegaCliente("CE_QuantidadeCompras") & ", " & ConverteVirgula(Format(RSPegaCliente("CE_JurosCartorio"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataCadastro"), "mm/dd/yyyy") & "', '" & Format(RSPegaCliente("CE_DataCancelamento"), "mm/dd/yyyy") & "', '" & RSPegaCliente("CE_Alteracao") & "', '" & RSPegaCliente("CE_Situacao") & "', '" & Format(RSPegaCliente("CE_HoraManutencao"), "mm/dd/yyyy hh:mm") & "')"
            rdoCNLoja.Execute (SQL)
        Else
            SQL = "Delete Cliente where CE_CodigoCliente=" & Val(mskCodCliente) & ""
            rdoCNLoja.Execute (SQL)
            
            If Err.Number = 0 Then
                rdoCNLoja.CommitTrans
            Else
                rdoCNLoja.RollbackTrans
            End If
            
            rdoCNLoja.BeginTrans
            SQL = ""
            SQL = "Insert into Cliente ([CE_CodigoCliente], [CE_CGC], [CE_InscricaoEstadual], [CE_Razao], " _
                & "[CE_Endereco], [CE_Bairro], [CE_Municipio], [CE_Estado], [CE_CEP], [CE_Telefone], [CE_Fax], " _
                & "[CE_EMail], [CE_TipoPessoa], [CE_Praca], [CE_PagamentoCarteira], [CE_EnderecoCobranca], " _
                & "[CE_BairroCobranca], [CE_MunicipioCobranca], [CE_EstadoCobranca], [CE_CEPCobranca], [CE_LimiteCredito], " _
                & "[CE_DataLimiteCredito], [CE_MaiorCompra], [CE_DataMaiorCompra], [CE_UltimaCompra], [CE_DataUltimaCompra], " _
                & "[CE_UltimoPagamento], [CE_DataUltimoPagamento], [CE_MaiorAtraso], [CE_QuantidadeCompras], [CE_JurosCartorio], [CE_DataCadastro], [CE_DataCancelamento], [CE_Alteracao], [CE_Situacao], [CE_HoraManutencao])" _
                & "Values(" & RSPegaCliente("CE_CodigoCliente") & ", '" & RSPegaCliente("CE_CGC") & "', '" & RSPegaCliente("CE_InscricaoEstadual") & "', '" & RSPegaCliente("CE_Razao") & "', " _
                & "'" & RSPegaCliente("CE_Endereco") & "', '" & RSPegaCliente("CE_Bairro") & "', '" & RSPegaCliente("CE_Municipio") & "', '" & RSPegaCliente("CE_Estado") & "', '" & RSPegaCliente("CE_CEP") & "', '" & RSPegaCliente("CE_Telefone") & "', '" & RSPegaCliente("CE_Fax") & "', " _
                & "'" & RSPegaCliente("CE_EMail") & "', '" & RSPegaCliente("CE_TipoPessoa") & "', " & RSPegaCliente("CE_Praca") & ", '" & RSPegaCliente("CE_PagamentoCarteira") & "', '" & RSPegaCliente("CE_EnderecoCobranca") & "', " _
                & "'" & RSPegaCliente("CE_BairroCobranca") & "', '" & RSPegaCliente("CE_MunicipioCobranca") & "', '" & RSPegaCliente("CE_EstadoCobranca") & "', '" & RSPegaCliente("CE_CEPCobranca") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_LimiteCredito"), "0.00")) & ", " _
                & "'" & Format(RSPegaCliente("CE_DataLimiteCredito"), "mm/dd/yyyy") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_MaiorCompra"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataMaiorCompra"), "mm/dd/yyyy") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_UltimaCompra"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataUltimaCompra"), "mm/dd/yyyy") & "', " _
                & "" & ConverteVirgula(Format(RSPegaCliente("CE_UltimoPagamento"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataUltimoPagamento"), "mm/dd/yyyy") & "', " & ConverteVirgula(Format(RSPegaCliente("CE_MaiorAtraso"), "0.00")) & ", " & RSPegaCliente("CE_QuantidadeCompras") & ", " & ConverteVirgula(Format(RSPegaCliente("CE_JurosCartorio"), "0.00")) & ", '" & Format(RSPegaCliente("CE_DataCadastro"), "mm/dd/yyyy") & "', '" & Format(RSPegaCliente("CE_DataCancelamento"), "mm/dd/yyyy") & "', '" & RSPegaCliente("CE_Alteracao") & "', '" & RSPegaCliente("CE_Situacao") & "', '" & Format(RSPegaCliente("CE_HoraManutencao"), "mm/dd/yyyy hh:mm") & "')"
            rdoCNLoja.Execute (SQL)
        End If
        If Err.Number = 0 Then
            rdoCNLoja.CommitTrans
            MsgBox "Cliente gravado com sucesso.", vbInformation, "Balcao 2000"
            mskCodCliente.SelStart = 0
            mskCodCliente.SelLength = Len(mskCodCliente)
            mskCodCliente.SetFocus
        Else
            rdoCNLoja.RollbackTrans
            MsgBox "Problemas gravação do cliente. Contate o CPD.", vbInformation, "Balcao 2000"
            mskCodCliente.SelStart = 0
            mskCodCliente.SelLength = Len(mskCodCliente)
            mskCodCliente.SetFocus
        End If
    Else
        MsgBox "Código do Cliente não existe.", vbInformation, "Balcao 2000"
        mskCodCliente.SelStart = 0
        mskCodCliente.SelLength = Len(mskCodCliente)
        mskCodCliente.SetFocus
    End If
    
    RSPegaCliente.Close
    VerificaCliente.Close
    
End Sub

Private Sub cmdRetorna_Click()

    'Conexao.Close
    rdoCNMatriz.Close
    rdoCNMatriz.ConnectionString = ""
    Unload Me
    

    frmConsCliente.cmdImportar.Enabled = False

End Sub

Private Sub Form_Load()
  Left = 10700
  Top = 6500
  
  Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
  Skin1.ApplySkin Me.hwnd
  
    Screen.MousePointer = 11
    On Error Resume Next
       ConectaODBCMatriz
        If GLB_ConectouOK = False Then
            Screen.MousePointer = 0
            MsgBox "Erro ao conectar-se ao Banco de Dados da Matriz", vbCritical, "Atenção"
            Exit Sub
        End If
    Screen.MousePointer = 0
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmConsCliente.cmdImportar.Enabled = False
    

End Sub

Private Sub mskCodCliente_GotFocus()
    
    mskCodCliente.SelStart = 0
    mskCodCliente.SelLength = Len(mskCodCliente)
    mskCodCliente.SetFocus

End Sub
