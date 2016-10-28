VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmFrete 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Frete"
   ClientHeight    =   5625
   ClientLeft      =   8025
   ClientTop       =   2760
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   30
      ScaleHeight     =   45
      ScaleWidth      =   6360
      TabIndex        =   10
      Top             =   4845
      Width           =   6360
   End
   Begin VB.TextBox txtTotalGeral 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   2010
      Width           =   2730
   End
   Begin VB.TextBox txtFrete 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1215
      TabIndex        =   7
      Text            =   "0"
      Top             =   1185
      Width           =   1635
   End
   Begin VB.Frame fraTipoFrete 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   195
      TabIndex        =   2
      Top             =   1110
      Width           =   765
      Begin VB.OptionButton optValor 
         BackColor       =   &H00505050&
         Caption         =   "Valor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   4
         Top             =   135
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton optPercentual 
         BackColor       =   &H00505050&
         Caption         =   "%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   3
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.TextBox txtTotalPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   150
      TabIndex        =   1
      Text            =   "0"
      Top             =   345
      Width           =   2730
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3285
      OleObjectBlob   =   "frmFrete.frx":0000
      Top             =   3720
   End
   Begin VB.TextBox txtPedido 
      Height          =   285
      Left            =   2970
      TabIndex        =   0
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5175
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmFrete.frx":0234
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblValorTotalFrete 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total do Pedido"
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
      Height          =   210
      Left            =   165
      TabIndex        =   9
      Top             =   1800
      Width           =   1890
   End
   Begin VB.Label lblFrete 
      BackStyle       =   0  'Transparent
      Caption         =   "Frete"
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
      Height          =   210
      Left            =   195
      TabIndex        =   6
      Top             =   885
      Width           =   960
   End
   Begin VB.Label lblValorTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Pedido"
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
      Height          =   210
      Left            =   165
      TabIndex        =   5
      Top             =   135
      Width           =   1395
   End
End
Attribute VB_Name = "frmFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wFrete As Double
Dim wCodigo As Integer
Dim wSequencia As Integer
Dim wValorDados As String
Dim SQL As String

Private Sub cmdGrava_Click()
  
    wCodigo = 1
  wSequencia = 16
  wValorDados = Format(wFrete, "0.00")
  
 If Not IsNumeric(wValorDados) Then
    Unload Me
    Exit Sub
 End If
 
 If Numeros(wValorDados) = "" Then
    Unload Me
    Exit Sub
 End If
 
 If Numeros(wValorDados) <= 0 Then
    Unload Me
    Exit Sub
 End If
 
'  SQL = "Select * from ComplementoVenda where COV_NumeroPedido = " & txtPedido.Text _
'      & " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " _
'      & wSequencia
'      rsComplementoVenda.CursorLocation = adUseClient
'      rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

  On Error GoTo erronaInclusao
  
'  If rsComplementoVenda.EOF Then
'     adoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & "," & wCodigo & "," & wSequencia & ",'" & wValorDados & "'" ' , rdExecDirect
'     rsComplementoVenda.Close
'     frmPedido.cmdBotoes(8).Enabled = False
'     frmPedido.cmdBotoes(10).visible = False
'     Unload Me
'     Exit Sub
'  Else
'     adoCNLoja.BeginTrans
'     Screen.MousePointer = vbHourglass
'     SQL = ""
'     SQL = "Update ComplementoVenda set COV_ValorComplemento = '" & wValorDados & "'" _
'           & " Where COV_Numeropedido = " & txtPedido.Text & _
'           " and COV_CodigoComplemento = " & wCodigo & " and COV_SequenciaComplemento = " & wSequencia
'            adoCNLoja.Execute SQL
'            Screen.MousePointer = vbNormal
'            adoCNLoja.CommitTrans
'            rsComplementoVenda.Close
'            frmPedido.cmdBotoes(8).Enabled = False
'            frmPedido.cmdBotoes(10).visible = False
'            Unload Me
'            Exit Sub
'  End If

  Screen.MousePointer = vbHourglass
  SQL = ""
  SQL = "FRETECOBR = " & ConverteVirgula(Format(wFrete, "##0.00"))
  adoCNLoja.BeginTrans
  adoCNLoja.Execute "Exec SP_GravaComplementoVenda " & txtpedido.Text & ",1,1,'" & SQL & "'"
  adoCNLoja.CommitTrans
  GBL_Frete = wFrete
  
 
  frmPedido.cmdTotalPedido.Caption = Format(txtTotalPedido.Text + GBL_Frete, "###,###,###,##0.00")

   'Do While Len(frmPedido.cmdTotalPedido.Caption) <= 12
       'frmPedido.cmdTotalPedido.Caption = frmPedido.cmdTotalPedido.Caption ' + " "
   'Loop
  
  
  Screen.MousePointer = vbNormal
  Unload Me
 ' frmPedido.picQuadroGeral.Width = 11550
  frmPedido.txtPesquisar.SetFocus
  
  Exit Sub
  
erronaInclusao:
       MsgBox "Erro na atualizazão do Pedido de Venda com valor do Frete " & vbLf & _
              Err.description, vbCritical, "Aviso"
       adoCNLoja.RollbackTrans
       Screen.MousePointer = vbNormal
End Sub

Private Sub cmdGrava_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub cmdRetorna_Click()
 Unload Me
' frmPedido.picQuadroGeral.Width = 11550
 frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub cmdRetorna_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_Activate()
    Call AjustaTela(frmFrete)
  
 ' Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
    
  txtpedido.Text = frmPedido.txtpedido
  txtFrete.SelStart = 0
  txtFrete.SelLength = Len(txtFrete.Text)
  GBL_Frete = 0
  
  
  SQL = ""
  SQL = "FRETECOBR = " & ConverteVirgula(Format(0, "##0.00"))
  adoCNLoja.Execute "Exec SP_GravaComplementoVenda " & txtpedido.Text & ",1,1,'" & SQL & "'"
  
 SQL = "Select (sum(vltotitem) - sum(desconto)) as vltotitem  From Nfitens Where NumeroPed = " & frmPedido.txtpedido.Text
  rsComplementoVenda.CursorLocation = adUseClient
  rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

  frmPedido.cmdTotalPedido.Caption = Format(rsComplementoVenda("vltotitem") + GBL_Frete, "###,###,###,##0.00")
  txtTotalPedido.Text = Trim(frmPedido.cmdTotalPedido.Caption)
    
 rsComplementoVenda.Close
   'Do While Len(frmPedido.cmdTotalPedido.Caption) <= 12
       'frmPedido.cmdTotalPedido.Caption = frmPedido.cmdTotalPedido.Caption ' + " "
   'Loop


End Sub

Private Sub optPercentual_Click()
txtFrete.SetFocus
End Sub

Private Sub optValor_Click()
txtFrete.SetFocus
End Sub

Private Sub txtFrete_Change()

If IsNumeric(txtFrete.Text) = False Then
        txtFrete.Text = ""
        txtFrete.SelStart = 0
        txtFrete.SelLength = Len(txtFrete.Text)
        txtFrete.SetFocus
    ElseIf txtFrete.Text < 0 Then
        txtFrete.Text = ""
        txtFrete.SelStart = 0
        txtFrete.SelLength = Len(txtFrete.Text)
        txtFrete.SetFocus
    End If
End Sub

Private Sub txtFrete_GotFocus()
txtFrete.Text = ""
txtFrete.SelStart = 0
txtTotalGeral.Text = txtTotalPedido.Text
txtFrete.SelLength = Len(txtFrete.Text)
End Sub

Private Sub txtFrete_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)

If KeyAscii = 46 Then
   txtFrete.Text = ""
   txtFrete.SelStart = 0
   txtFrete.SelLength = Len(txtFrete.Text)
   txtFrete.SetFocus
   Exit Sub
End If

If KeyAscii = vbKeyReturn Then
   cmdGrava.SetFocus
End If
If KeyAscii = 27 Then
   Unload Me

End If

End Sub

Private Sub txtFrete_LostFocus()
 
txtFrete.Text = Format(txtFrete.Text, "###,###,###,##0.00")
 
If txtFrete.Text <> "" Then
  If IsNumeric(txtFrete.Text) = False Then
     txtFrete.Text = ""
     txtFrete.SetFocus
     txtFrete.SelStart = 0
     txtFrete.SelLength = Len(txtFrete.Text)
     Exit Sub
  End If
Else
Exit Sub
End If
 
  If optValor.Value = True Then
     wFrete = txtFrete.Text
     txtTotalGeral.Text = Format((txtTotalPedido.Text + wFrete), "###,###,###,##0.00")
  ElseIf optPercentual.Value = True Then
     txtTotalGeral.Text = Format((txtTotalPedido.Text + ((txtTotalPedido.Text * txtFrete.Text) / 100)), "###,###,###,##0.00")
  End If
'  wFrete = ConverteVirgula1(Format((txtTotalPedido.text - txtTotalGeral.Text), "###,###,###,##0.00"))
  wFrete = Format((txtTotalGeral.Text - txtTotalPedido.Text), "###,###,###,##0.00")
      
End Sub


Private Sub txtTotalGeral_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtTotalGeral_KeyPress(KeyAscii As Integer)
    cmdGrava.SetFocus
    If KeyAscii = 27 Then
    Unload Me
    End If
    
End Sub

