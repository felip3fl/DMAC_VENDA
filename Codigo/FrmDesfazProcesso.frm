VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmDesfazProcesso 
   Caption         =   "Desfaz Processo"
   ClientHeight    =   2910
   ClientLeft      =   3705
   ClientTop       =   2655
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   3480
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FrmDesfazProcesso.frx":0000
      Top             =   2265
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ok"
      Height          =   390
      Left            =   2115
      TabIndex        =   8
      Top             =   2325
      Width           =   1305
   End
   Begin VB.TextBox txtpedido 
      Height          =   285
      Left            =   2970
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   105
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3375
      Begin VB.CheckBox ChkCliente 
         Caption         =   "Cliente"
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   1755
         Width           =   1710
      End
      Begin VB.CheckBox ChkCarimbo 
         Caption         =   "Carimbo"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   1545
         Width           =   960
      End
      Begin VB.CheckBox ChkTR 
         Caption         =   "TR"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   1275
         Width           =   1665
      End
      Begin VB.CheckBox ChkVendaDistancia 
         Caption         =   "Venda Distância"
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   1080
         Width           =   1590
      End
      Begin VB.CheckBox ChkTransferencia 
         Caption         =   "Transferencia"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   855
         Width           =   2805
      End
      Begin VB.CheckBox ChkFrete 
         Caption         =   "Frete"
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   2670
      End
      Begin VB.CheckBox ChkFormaPagto 
         Caption         =   "Pagamento"
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   180
         TabIndex        =   2
         Top             =   390
         Width           =   2625
      End
      Begin VB.CheckBox ChkDesconto 
         Caption         =   "Desconto"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   135
         Width           =   2685
      End
   End
End
Attribute VB_Name = "FrmDesfazProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCodigo As Integer
Dim wSequencia As Integer
Dim SQL As String

Private Sub CmdOK_Click()

On Error GoTo erronoUpdate

wCodigo = 1

If ChkDesconto.Value = 1 Then
   frmPedido.cmdBotoes(9).Enabled = True
   wSequencia = 15
   Call DesfazProcesso
End If
If ChkFormaPagto = 1 Then

   wSequencia = 4
   Call DesfazProcesso
   wSequencia = 17
   Call DesfazProcesso
End If
If ChkFrete.Value = 1 Then
   frmPedido.cmdBotoes(10).Visible = True
   wSequencia = 16
   Call DesfazProcesso
End If
If ChkTransferencia = 1 Then
   frmPedido.cmdBotoes(9).Visible = True
   

   frmPedido.cmdBotoes(10).Visible = True
   frmPedido.cmdBotoes(12).Visible = True
'   frmPedido.cmdTR.Visible = True
   frmPedido.cmdBotoes(6).Visible = True
   frmPedido.cmdBotoes(7).Visible = True
   wSequencia = 1
   Call DesfazProcesso
   wSequencia = 2
   Call DesfazProcesso
   wSequencia = 3
   Call DesfazProcesso
   wSequencia = 8
   Call DesfazProcesso
   wSequencia = 14
   Call DesfazProcesso
   
   
   adoCNLoja.BeginTrans
   Screen.MousePointer = vbHourglass
 
   SQL = "Update ItensVenda set ITV_Situacao = 'I' where ITV_Numeropedido = " _
         & txtpedido.Text
         adoCNLoja.Execute SQL
         Screen.MousePointer = vbNormal
         adoCNLoja.CommitTrans
End If

If ChkVendaDistancia = 1 Then
'   frmPedido.cmdTR.Visible = True
   frmPedido.cmdBotoes(6).Visible = True
 '  frmPedido.cmdBotoes(8).Visible = True
   wSequencia = 9
   Call DesfazProcesso
   wSequencia = 10
   Call DesfazProcesso
End If
If ChkTR = 1 Then
 '   frmPedido.cmdTR.Visible = True
    frmPedido.cmdBotoes(6).Visible = True
 '   frmPedido.cmdBotoes(8).Visible = True
   wSequencia = 19
   Call DesfazProcesso
   
   adoCNLoja.BeginTrans
   Screen.MousePointer = vbHourglass
 
   SQL = "Update ItensVenda set ITV_DescricaoAlternativa = '' where ITV_Numeropedido = " _
         & txtpedido.Text
         adoCNLoja.Execute SQL
         Screen.MousePointer = vbNormal
         adoCNLoja.CommitTrans
   
   adoCNLoja.BeginTrans
   Screen.MousePointer = vbHourglass
 
   SQL = "Update ItensVenda set ITV_PrecoAlternativo = 0.00 where ITV_Numeropedido = " _
         & txtpedido.Text
         adoCNLoja.Execute SQL
         Screen.MousePointer = vbNormal
         adoCNLoja.CommitTrans
   
End If

If ChkCliente.Value = 1 Then
   'frmPedido.cmdBotoes(9).Enabled = True
   wSequencia = 6
   Call DesfazProcesso
'   frmPedido.cmdTR.Visible = True
   frmPedido.cmdBotoes(6).Enabled = True
'   frmPedido.cmdBotoes(8).Enabled = True
   frmPedido.cmdBotoes(10).Visible = True
End If

If ChkCarimbo = 1 Then
'   frmPedido.cmdBotoes(8).Enabled = True
   frmPedido.cmdBotoes(7).Enabled = True
   wSequencia = 22
   Call DesfazProcesso
   wSequencia = 23
   Call DesfazProcesso
   wSequencia = 24
   Call DesfazProcesso
   wSequencia = 25
   Call DesfazProcesso
End If


'frmPedido.cmdBotoes(8).Enabled = True

inibebotoes (FrmDesfazProcesso.txtpedido)

Unload Me

Exit Sub

erronoUpdate:
MsgBox "Erro na atualização do pedido " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
Exit Sub
End Sub

Private Sub DesfazProcesso()
SQL = "Select * from ComplementoVenda where COV_NumeroPedido = " & txtpedido.Text _
      & " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " _
      & wSequencia
      rsComplementoVenda.CursorLocation = adUseClient
      rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

  On Error GoTo ErronaDelecao
  
  If rsComplementoVenda.EOF Then
     rsComplementoVenda.Close
     Exit Sub
  Else
     adoCNLoja.BeginTrans
     Screen.MousePointer = vbHourglass
     SQL = ""
     SQL = "Delete ComplementoVenda " _
           & " Where COV_Numeropedido = " & txtpedido.Text & _
           " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " & wSequencia
            adoCNLoja.Execute SQL
            Screen.MousePointer = vbNormal
            adoCNLoja.CommitTrans
            rsComplementoVenda.Close
            
            Exit Sub
  End If
  
ErronaDelecao:
MsgBox "Erro ao Deletar Complemento de Venda Sequencia ==> " & wSequencia & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
rsComplementoVenda.Close
End Sub




Private Sub Form_Load()
left = (Screen.Width - Width) / 2
top = (Screen.Height - Height) / 2

  
'  Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
'  Skin1.ApplySkin Me.hwnd


End Sub

