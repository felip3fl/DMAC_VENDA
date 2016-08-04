VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmInicio 
   Caption         =   "Conectar"
   ClientHeight    =   1935
   ClientLeft      =   8610
   ClientTop       =   3495
   ClientWidth     =   3060
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   3060
   Begin VB.ComboBox cmb_Caixa 
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   765
      Width           =   1035
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   345
      Left            =   300
      OleObjectBlob   =   "frmInicio.frx":23FA
      TabIndex        =   4
      Top             =   795
      Width           =   660
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   375
      Left            =   1530
      TabIndex        =   3
      Top             =   1305
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   285
      TabIndex        =   2
      Top             =   1305
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   330
      Left            =   315
      OleObjectBlob   =   "frmInicio.frx":2462
      TabIndex        =   1
      Top             =   375
      Width           =   570
   End
   Begin VB.ComboBox cmb_loja 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   330
      Width           =   1035
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmb_loja_Click()


SQL = "Select CXA_NumeroCaixa from ParametroSistema order by CXA_NumeroCaixa"

rdoParametroINI.CursorLocation = adUseClient
rdoParametroINI.Open SQL, adoCNAccess, adOpenForwardOnly, adLockPessimistic

        If Not rdoParametroINI.EOF Then
            cmb_Caixa.Clear
            Do While Not rdoParametroINI.EOF
                cmb_Caixa.AddItem Trim(rdoParametroINI("CXA_NumeroCaixa"))
                rdoParametroINI.MoveNext
            Loop

            Screen.MousePointer = 0
            cmb_Caixa.ListIndex = 0
        End If

rdoParametroINI.Close
End Sub

Private Sub cmb_loja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1_Click
End If

End Sub

Private Sub Command1_Click()

SQL = "Select * from ConexaoSistema where GLB_Loja = '" & Trim(cmb_loja.Text) & "'"

rdoConexaoINI.CursorLocation = adUseClient
rdoConexaoINI.Open SQL, adoCNAccess, adOpenForwardOnly, adLockPessimistic

  If Not rdoConexaoINI.EOF Then
       GLB_Servidor = Trim(rdoConexaoINI("GLB_ServidorRetaguarda"))
       GLB_Loja = Trim(rdoConexaoINI("GLB_Loja"))
       GLB_Banco = Trim(rdoConexaoINI("GLB_BancoRetaguarda"))
       GLB_Servidorlocal = Trim(rdoConexaoINI("GLB_ServidorLocal"))
       Glb_BancoLocal = Trim(rdoConexaoINI("GLB_BancoLocal"))
       'GLB_Usuario = Trim(rdoConexaoINI("GLB_Usuario"))
       'GLB_Senha = Trim(rdoConexaoINI("GLB_Senha"))
       rdoConexaoINI.Close
  End If

  ConectaODBC
  
  ShellExecute Hwnd, "open", ("C:\Sistemas\DMAC Venda\limpaCache"), "", "", 1
    
Continua:

    If GLB_ConectouOK = True Then
       'Me.Visible = False
       Call DadosLoja
       'frmPedido.Show
       frmBandeja.Show
       frmPedido.ZOrder
       Unload Me

       Else
           MsgBox "Erro ao conectar-se ao Banco de Dados", vbCritical, "Atenção"
           Exit Sub
       End If


End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
'Skin1.LoadSkin "c:\WINDOWS\system\skin.skn"
'Skin1.ApplySkin Me.Hwnd

SQL = "Select GLB_LOJA from ConexaoSistema GROUP BY GLB_LOJA"
 
rdoConexaoINI.CursorLocation = adUseClient
rdoConexaoINI.Open SQL, adoCNAccess, adOpenForwardOnly, adLockPessimistic
 
        If Not rdoConexaoINI.EOF Then
            cmb_loja.Clear
            Do While Not rdoConexaoINI.EOF
                cmb_loja.AddItem Trim(rdoConexaoINI("GLB_LOJA"))
                rdoConexaoINI.MoveNext
            Loop
            
            Screen.MousePointer = 0
            
             
        End If

rdoConexaoINI.Close

cmb_loja.ListIndex = 0

Exit Sub
ConexaoErro:
MsgBox "Erro ao abrir banco de Dados da Loja! "
End

End Sub

