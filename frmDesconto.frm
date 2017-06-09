VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmDesconto 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Desconto"
   ClientHeight    =   7260
   ClientLeft      =   2985
   ClientTop       =   2190
   ClientWidth     =   13260
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFormaDesconto 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   3990
      TabIndex        =   20
      Top             =   0
      Width           =   2325
      Begin VB.CheckBox ChcSimula 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Caption         =   "Simular Desconto"
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
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame fraFormaDesconto 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      ForeColor       =   &H00404040&
      Height          =   615
      Index           =   0
      Left            =   4935
      TabIndex        =   17
      Top             =   240
      Width           =   1380
      Begin VB.OptionButton optPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Caption         =   "No Pedido"
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
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Caption         =   "No Item"
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
         Height          =   195
         Left            =   30
         TabIndex        =   18
         Top             =   375
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   14
      Top             =   4875
      Width           =   6165
   End
   Begin VB.Frame FrmeSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   165
      TabIndex        =   11
      Top             =   2355
      Visible         =   0   'False
      Width           =   2430
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   465
         Width           =   2310
      End
      Begin VB.Label lblSenhaGerente 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha do Gerente"
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
         Left            =   75
         TabIndex        =   13
         Top             =   195
         Width           =   2100
      End
   End
   Begin VB.Frame fraTipoDesconto 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3990
      TabIndex        =   8
      Top             =   240
      Width           =   975
      Begin VB.OptionButton optValor 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Caption         =   "Valor"
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
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optPercentual 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Caption         =   "%"
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
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   465
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
      TabIndex        =   4
      Text            =   "0"
      Top             =   375
      Width           =   3465
   End
   Begin VB.TextBox txtTotalGeral 
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
      TabIndex        =   3
      Text            =   "0"
      Top             =   1155
      Width           =   3465
   End
   Begin VB.TextBox txtDesconto 
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
      Left            =   3990
      TabIndex        =   0
      Top             =   1155
      Width           =   2325
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   450
      OleObjectBlob   =   "frmDesconto.frx":0000
      Top             =   5910
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   150
      TabIndex        =   1
      Text            =   "0"
      Top             =   5910
      Visible         =   0   'False
      Width           =   270
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5250
      TabIndex        =   2
      Top             =   5040
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
      MICON           =   "frmDesconto.frx":0234
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensPedido 
      Height          =   2295
      Left            =   150
      TabIndex        =   23
      Top             =   1800
      Width           =   6165
      _cx             =   10874
      _cy             =   4048
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
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDesconto.frx":0250
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
   End
   Begin VB.Label lblMagenLinha 
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
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblMagenTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5265
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label txtDescricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do item selecionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   150
      TabIndex        =   15
      Top             =   4170
      Width           =   3240
   End
   Begin VB.Label lblTotalPedido 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pedido"
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
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   2100
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pedido com Desconto"
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
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   930
      Width           =   2385
   End
   Begin VB.Label lblDesconto 
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto"
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
      Left            =   3990
      TabIndex        =   5
      Top             =   930
      Width           =   1425
   End
End
Attribute VB_Name = "frmDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I As Integer
Dim wCodigo As Integer
Dim wSequencia As Double
Dim wValorDados As String
Dim wTotalPedido As Double
Dim wDescontoMax As Double
Dim wDescontoRat As Double
Dim wTotal As Double
Dim wDesconto As Double

Dim Sql As String
Dim resposta As String
Dim GLB_Senha As String
Dim wReferencia As String

Dim wCond As String
Dim flgpad As Integer
Dim wProdutoPromocao As Boolean
Dim wComDesconto As String
Dim wDescontoPedido As Double
Dim wTotalPed As Double
Dim wLimiteDescontoAtingido As Boolean
Dim wLiberaBloqueio As Boolean
  Dim rsMargem As New ADODB.Recordset
  Dim margemB As String
  Dim wTipoBloqueio As String


Private Sub ChcSimula_Click()
margemB = ""
lblMagenLinha.Visible = False
lblMagenTotal.Visible = False
cmdGrava.Enabled = False
txtDesconto.Text = ""
optPedido.Value = True

    If ChcSimula.Value = 1 Then
        wProdutoPromocao = False
        Call CarregaItensPedido
        liquidotodas
        optPedido_Click
        ChcSimula.Enabled = False
        
          Sql = "Select CTS_MostraMgSimulador from controlesistema"
        rsMargem.CursorLocation = adUseClient
        rsMargem.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        margemB = Trim(rsMargem("CTS_MostraMgSimulador"))
        rsMargem.Close
        
    Else
    Call CarregaItensPedido
    optItem_Click
    I = 0
        For I = 1 To grdItensPedido.Rows - 1
            If grdItensPedido.TextMatrix(I, 0) = "P" Then
                wProdutoPromocao = True
             End If
              optItem.Value = True
              If ChcSimula.Value = 0 And I <> grdItensPedido.Rows Then
          grdItensPedido.TextMatrix(I, 5) = Format(0, "###,###,###,##0.00")
      grdItensPedido.TextMatrix(I, 6) = Format(0, "###,###,###,##0.00")
      End If
        Next I
        
      
  

    End If
   
If margemB = "N" Then
        
         grdItensPedido.ColHidden(11) = False
         grdItensPedido.ColHidden(12) = True
Else
                 
         grdItensPedido.ColHidden(11) = True
         grdItensPedido.ColHidden(12) = False
End If
      
    

End Sub

Private Sub cmdGrava_Click()
   
  wCodigo = 1
  wSequencia = 15
  wValorDados = Format(wDesconto, "0.00")
 
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
 
 wTotalPedido = Format(txtTotalPedido.Text, "###,###,###,##0.00")
 If wDesconto > wTotalPedido Then
    txtDesconto.Enabled = True
    fraTipoDesconto.Enabled = True
    'fraFormaDesconto.Enabled = True
    txtDesconto.SelStart = 0
    txtDesconto.SelLength = Len(txtDesconto.Text)
    txtDesconto.Text = ""
    txtDesconto.SetFocus
    txtTotalGeral.Text = txtTotalPedido.Text
    Exit Sub
 End If


  On Error GoTo erronaInclusao
  
 
  Screen.MousePointer = vbHourglass
  
  Sql = ""
  wReferencia = ""
        For I = 1 To grdItensPedido.Rows - 1
          Sql = " UPDATE NFItens Set Desconto = " & ConverteVirgula(Format(grdItensPedido.TextMatrix(I, 5), "#0.00")) & _
                " Where NumeroPed = " & txtpedido.Text & " And referencia = " & grdItensPedido.TextMatrix(I, 1)
'          adoCNLoja.BeginTrans
          adoCNLoja.Execute Sql
'          adoCNLoja.CommitTrans
        Next I
  
    
  Sql = ""
    
   Sql = " UPDATE NFCapa Set Desconto = " & ConverteVirgula(wDesconto) & _
         " Where NumeroPed = " & txtpedido.Text
          adoCNLoja.Execute Sql
    
  Screen.MousePointer = vbNormal
  
  

  frmPedido.cmdTotalPedido.Caption = Format(txtTotalGeral.Text + GBL_Frete, "###,###,###,##0.00")
  'frmPedido.cmdTotalPedidoGE.Caption = "+ G.E " & _
  frmPedido.cmdTotalPedido.Caption + frmGarantiaEstendida.valorGarantiaEstendida
  frmPedido.cmdTotalPedidoGE.Caption = "+ G.E " & frmGarantiaEstendida.valorGarantiaEstendida

     'Do While Len(frmPedido.cmdTotalPedido.Caption) <= 12
         'frmPedido.cmdTotalPedido.Caption = frmPedido.cmdTotalPedido.Caption ' + " "
     'Loop

  Unload Me
  frmPedido.txtPesquisar.SetFocus
  Exit Sub
  
 FrmeSenha.Visible = False
    'fraFormaDesconto.Enabled = True
    fraTipoDesconto.Enabled = True
    grdItensPedido.Enabled = True
 cmdGrava.Enabled = False
  
erronaInclusao:
MsgBox "Erro na atualização do Complemento de Venda com valor do desconto " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
rsComplementoVenda.Close
End Sub

Private Sub cmdGrava_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cmdGrava.Enabled = False
        fraTipoDesconto.Enabled = True
        'fraFormaDesconto.Enabled = True
        txtDesconto.Enabled = True
        txtDesconto.SetFocus
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
  Dim rsDesconto As New ADODB.Recordset

  wLiberaBloqueio = False
  cmdGrava.Enabled = False
  wProdutoPromocao = False
  
  Sql = "Select sum(vltotitem) as vltotitem From Nfitens Where NumeroPed = " & frmPedido.txtpedido.Text
  rsDesconto.CursorLocation = adUseClient
  rsDesconto.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  frmPedido.cmdTotalPedido.Caption = Format(rsDesconto("vltotitem") + GBL_Frete, "###,###,###,##0.00")
  
  'Do While Len(frmPedido.cmdTotalPedido.Caption) <= 12
    'frmPedido.cmdTotalPedido.Caption = frmPedido.cmdTotalPedido.Caption '+ " "
  'Loop
   
  txtTotalPedido.Text = Format(rsDesconto("vltotitem"), "###,###,###,##0.00")
  txtTotalGeral.Text = txtTotalPedido.Text
  txtpedido.Text = frmPedido.txtpedido.Text
  rsDesconto.Close

  

  Sql = "Update Nfcapa set desconto = 0 Where NumeroPed = " & frmPedido.txtpedido.Text
  adoCNLoja.Execute Sql


  Sql = "Update Nfitens set desconto = 0 Where NumeroPed = " & frmPedido.txtpedido.Text
  adoCNLoja.Execute Sql
  
  

  grdItensPedido.Rows = 1

  Call CarregaItensPedido

''  SQL = "Select pr_classe, 'N' as liberabloqueio from nfitens as i, nfcapa as c, produtoloja " _
''        & "where pr_referencia = i.referencia and i.numeroped = " & frmPedido.txtpedido.Text & " " _
''        & "and i.numeroped = c.numeroped"
''   '''CONTINENTAL
  Sql = "Select pr_classe, liberabloqueio from nfitens as i, nfcapa as c, produtoloja " _
        & "where pr_referencia = i.referencia and i.numeroped = " & frmPedido.txtpedido.Text & " " _
        & "and i.numeroped = c.numeroped"
        
  rsDesconto.CursorLocation = adUseClient
  rsDesconto.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsDesconto.EOF Then
     Do While Not rsDesconto.EOF
            If rsDesconto("pr_classe") = "P" And (rsDesconto("liberabloqueio") <> "S" Or rsDesconto("liberabloqueio") <> "T") And ChcSimula.Value = 0 Then
                 wProdutoPromocao = True
                 optItem.Value = True
                 For I = 1 To grdItensPedido.Rows - 1
                    If (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then
                      grdItensPedido.TextMatrix(I, 0) = "N"
                    End If
                 Next I
            Else
                 optPedido.Value = True
                 For I = 1 To grdItensPedido.Rows - 1
                    If grdItensPedido.TextMatrix(I, 0) = "N" And grdItensPedido.TextMatrix(I, 12) = "S" Then
                      grdItensPedido.TextMatrix(I, 0) = "S"
                    End If
                 Next I
            End If
            rsDesconto.MoveNext
     Loop
  End If
 '   txtDesconto.SetFocus
  rsDesconto.Close


End Sub

Private Sub Form_Load()
    Call AjustaTela(frmDesconto)
  
  txtDesconto.SelStart = 0
  txtDesconto.SelLength = Len(txtDesconto.Text)
'  cmdGrava.Enabled = False
  grdItensPedido.Rows = 1
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If rsComplementoVenda.State = 1 Then
    rsComplementoVenda.Close
 End If
 grdItensPedido.Rows = 1

End Sub



Private Sub grdItensPedido_Click()
'    If grdItensPedido.Col = 0 Then
'       If optPedido.Value = False Then

'       Else
'          grdItensPedido.Editable = flexEDNone
'       End If
'    Else
'       grdItensPedido.Editable = flexEDNone
'    End If

   
   liquidolinha
     
     
     
End Sub

Private Sub grdItensPedido_DblClick()
  If optItem.Value = True Then
    If (Trim(grdItensPedido.TextMatrix(grdItensPedido.Row, 0)) = "S" Or Trim(grdItensPedido.TextMatrix(grdItensPedido.Row, 0)) = "T") Then
       grdItensPedido.TextMatrix(grdItensPedido.Row, 0) = "N"
    ElseIf Trim(grdItensPedido.TextMatrix(grdItensPedido.Row, 0)) = "N" Then
        If (grdItensPedido.TextMatrix(grdItensPedido.Row, 12) = "S" Or wTipoBloqueio = "E") Then
            grdItensPedido.TextMatrix(grdItensPedido.Row, 0) = "S"
        Else
            MsgBox "Desconto totalmente bloqueado nessa referência", vbExclamation
        End If
       
    End If
  End If
    
End Sub

Private Sub grdItensPedido_EnterCell()
'    lblDescricao.Caption = grdItensPedido.TextMatrix(grdItensPedido.Row, 5)
    txtDescricao.Caption = grdItensPedido.TextMatrix(grdItensPedido.Row, 4)
    If ChcSimula.Value = 1 Then
        liquidolinha
    End If
    If grdItensPedido.TextMatrix(grdItensPedido.Row, 0) = "N" Then
     lblMagenLinha.Caption = ""
    End If
    

End Sub

Private Sub grdItensPedido_KeyDown(KeyCode As Integer, Shift As Integer)
 
If KeyCode = vbKeyF2 Then
  cmdGrava_Click
 ElseIf KeyCode = 27 Then
    Unload Me
 ElseIf KeyCode = 13 Then
txtDesconto.SetFocus
    End If

End Sub

Private Sub lblTotalPedido_Click()
lblMagenLinha.Caption = ""
lblMagenTotal.Caption = ""
End Sub

Private Sub optItem_Click()
lblMagenLinha.Caption = ""
lblMagenTotal.Caption = ""

'    For i = 1 To grdItensPedido.Rows - 1
'        grdItensPedido.Cell(flexcpChecked, i, 0) = 2
'    Next i

If ChcSimula.Value = 1 Then
For I = 1 To grdItensPedido.Rows - 1
       If grdItensPedido.TextMatrix(I, 0) = "P" Then
        grdItensPedido.TextMatrix(I, 0) = "N"
       End If
       Next I
End If


    For I = 1 To grdItensPedido.Rows - 1
       If (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then
        grdItensPedido.TextMatrix(I, 0) = "N"
       End If
    Next I
    grdItensPedido.Enabled = True
    grdItensPedido.Row = 1
    txtDesconto.Text = ""
    txtDesconto.SelStart = 0
    txtTotalGeral.Text = txtTotalPedido.Text
    txtDesconto.SelLength = Len(txtDesconto.Text)
    cmdGrava.Enabled = False
End Sub

Private Sub optPedido_Click()

 If wProdutoPromocao = True And ChcSimula.Value = 0 Then
    MsgBox "Não é permitido dar desconto para produto em promoção."
    optItem.Value = True
 Else
'    For i = 1 To grdItensPedido.Rows - 1
'        grdItensPedido.Cell(flexcpChecked, i, 0) = 1
'    Next i
If ChcSimula.Value = 1 Then
For I = 1 To grdItensPedido.Rows - 1
       If grdItensPedido.TextMatrix(I, 0) = "P" Then
        grdItensPedido.TextMatrix(I, 0) = "S"
       End If
       Next I
End If


    For I = 1 To grdItensPedido.Rows - 1
       If grdItensPedido.TextMatrix(I, 0) = "N" Then
          grdItensPedido.TextMatrix(I, 0) = "S"
       End If
    Next I

    txtDesconto.Visible = True
    txtDesconto.Text = ""
    txtDesconto.SelStart = 0
    txtTotalGeral.Text = txtTotalPedido.Text
    txtDesconto.SelLength = Len(txtDesconto.Text)
    cmdGrava.Enabled = False
 End If
  

End Sub

Private Sub optPercentual_Click()
lblMagenLinha.Caption = ""
lblMagenTotal.Caption = ""
    If wProdutoPromocao = False Then
       optPedido.SetFocus
    Else
      optItem.SetFocus
    End If
    txtDesconto.Text = ""
    txtDesconto.SelStart = 0
    txtTotalGeral.Text = txtTotalPedido.Text
    txtDesconto.SelLength = Len(txtDesconto.Text)
    cmdGrava.Enabled = False
End Sub

Private Sub optValor_Click()
lblMagenLinha.Caption = ""
lblMagenTotal.Caption = ""
    If wProdutoPromocao = False Then
       optPedido.SetFocus
    Else
       optItem.SetFocus
    End If
    
    txtDesconto.Text = ""
    txtDesconto.SelStart = 0
    txtTotalGeral.Text = txtTotalPedido.Text
    txtDesconto.SelLength = Len(txtDesconto.Text)
    cmdGrava.Enabled = False
End Sub

Private Sub SkinLabel1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtDesconto_Change()
   
If IsNumeric(txtDesconto.Text) = False Then
   txtDesconto.Text = ""
   txtDesconto.SelStart = 0
   txtDesconto.SelLength = Len(txtDesconto.Text)
   cmdGrava.Enabled = False
'   txtDesconto.SetFocus
'ElseIf txtDesconto.Text <= 0 Then
'       txtDesconto.Text = ""
'       txtDesconto.SelStart = 0
'       txtDesconto.SelLength = Len(txtDesconto.Text)
'       txtDesconto.SetFocus
End If

End Sub



Private Sub txtDesconto_GotFocus()
   cmdGrava.Enabled = False
   campoSelecionadoComCaracter txtDesconto
End Sub

Private Sub txtDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = vbKeyF2 Then
  cmdGrava_Click
 ElseIf KeyCode = 27 Then
    Unload Me
 ElseIf KeyCode = vbEnter Then
cmdGrava.SetFocus
   Exit Sub
End If
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)

Dim wDif As Double
Dim Linha As Integer
Dim wValorAuxDesc As Double
Dim bloqueioID As Boolean
wDif = 0
wValorAuxDesc = 0

 
 If KeyAscii = 13 Then

   If Trim(txtDesconto.Text) = "" Then
       MsgBox "Informe o valor do desconto.", vbCritical, "Atenção"
       txtDesconto.SetFocus
       Exit Sub
   End If
   
   If (Val(txtDesconto.Text) > 99 Or Val(txtDesconto.Text) < 0) And optPercentual.Value = True Then
       MsgBox "Porcentagem de desconto invalida", vbCritical, "Atenção"
       txtDesconto.SetFocus
       Exit Sub
   End If
   
   If CDec(txtDesconto.Text) >= CDec(txtTotalGeral.Text) And optValor.Value = True Then
       MsgBox "Valor do desconto não pode ser maior ou igual ao valor do Pedido", vbCritical, "Atenção"
       txtDesconto.SetFocus
       Exit Sub
   End If
   
   wDescontoPedido = txtDesconto.Text
   wTotalPed = txtTotalPedido.Text
   If (optValor.Value = True And wDescontoPedido > Format(wTotalPed, "##0.00")) _
      Or (optPercentual.Value = True And wDescontoPedido > 100) Then
      
      MsgBox "Não é permitido desconto maior que o valor do pedido"
      txtDesconto.SelStart = 0
      txtDesconto.SelLength = Len(txtDesconto.Text)
      txtDesconto.Text = ""
      txtDesconto.SetFocus
      Exit Sub
   End If

   Call LerControle
   wReferencia = ""
   
If rsComplementoVenda.State = 1 Then
    rsComplementoVenda.Close
 End If
    
    
    
    Sql = ""
    Sql = "Select Referencia, PR_Descricao, VLUNIT, desconto,vltotitem, Qtde, PR_classe From NFItens, ProdutoLoja " _
          & "Where PR_Referencia = Referencia and TipoNota = 'PD' and " _
          & "NumeroPed = " & txtpedido.Text & " Order By Referencia"

    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    For I = 1 To grdItensPedido.Rows - 1
      grdItensPedido.TextMatrix(I, 5) = Format(0, "###,###,###,##0.00")
      grdItensPedido.TextMatrix(I, 6) = Format(0, "###,###,###,##0.00")
    Next I
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   If optItem.Value = True Then    '**************** DESCONTO POR ITEM

     If optValor.Value = True Then    '************ DESCONTO POR VALOR
        wTotalPedido = 0
        For I = 1 To grdItensPedido.Rows - 1
'          If grdItensPedido.Cell(flexcpChecked, i, 0) = 1 Then
          If (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then
             wTotalPedido = Format(wTotalPedido + grdItensPedido.TextMatrix(I, 3), "###,###,###,##0.00")
          Else
             grdItensPedido.TextMatrix(I, 5) = "0,00"
          End If
        Next I
        
        If wTotalPedido = 0 Then
           MsgBox "Selecione um produto para aplicar o desconto.", vbInformation, "Atenção"
           If rdoControle.status = 0 Then rdoControle.Close
           txtDesconto.SetFocus
           txtDesconto.SelStart = 0
           txtDesconto.SelLength = Len(txtDesconto.Text)
           
           Exit Sub
        End If
        
        wDesconto = txtDesconto.Text
        wDesconto = Format((wDesconto / wTotalPedido) * 100, "##0.0000")
        
            
        For I = 1 To grdItensPedido.Rows - 1
 '         If grdItensPedido.Cell(flexcpChecked, i, 0) = 1 Then
          
          If (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then
             grdItensPedido.TextMatrix(I, 6) = Format((grdItensPedido.TextMatrix(I, 3) - ((grdItensPedido.TextMatrix(I, 3) * wDesconto) / 100)), "###,###,###,##0.00")
             grdItensPedido.TextMatrix(I, 5) = Format(((grdItensPedido.TextMatrix(I, 3) * wDesconto) / 100), "##0.00")
            Linha = I
          Else
             grdItensPedido.TextMatrix(I, 6) = Format(grdItensPedido.TextMatrix(I, 3), "###,###,###,##0.00")
             grdItensPedido.TextMatrix(I, 5) = Format(0, "###,###,###,##0.00")
          End If
        Next I
        
        wDif = 0
        wValorAuxDesc = 0
        For I = 1 To grdItensPedido.Rows - 1
            wValorAuxDesc = wValorAuxDesc + grdItensPedido.TextMatrix(I, 5)
        Next I
       
        wDesconto = Format(txtDesconto.Text, "##0.00")
        wDif = Format(wDesconto - wValorAuxDesc, "#0.00")
            
       If wDif <> 0 Then
          grdItensPedido.TextMatrix(Linha, 5) = Format((grdItensPedido.TextMatrix(Linha, 5) + (wDif)), "##0.00")
          grdItensPedido.TextMatrix(Linha, 6) = Format((grdItensPedido.TextMatrix(Linha, 6) - (wDif)), "##0.00")
       End If


        wTotalPedido = 0
        wTotalPedido = Format((txtTotalPedido.Text - wDesconto), "###,###,###,##0.00")
        
     Else     '************ DESCONTO POR PERCENTUAL
        wValorAuxDesc = 0
        For I = 1 To grdItensPedido.Rows - 1
'          If grdItensPedido.Cell(flexcpChecked, i, 0) = 1 Then
          If (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then
             grdItensPedido.TextMatrix(I, 6) = Format((grdItensPedido.TextMatrix(I, 3) - ((grdItensPedido.TextMatrix(I, 3) * txtDesconto) / 100)), "###,###,###,##0.00")
             grdItensPedido.TextMatrix(I, 5) = Format(((grdItensPedido.TextMatrix(I, 3) * txtDesconto) / 100), "##0.00")
             wValorAuxDesc = wValorAuxDesc + Format(grdItensPedido.TextMatrix(I, 5), "##0.00")
          Else
             grdItensPedido.TextMatrix(I, 5) = "0,00"
             grdItensPedido.TextMatrix(I, 6) = grdItensPedido.TextMatrix(I, 3)
          End If
        Next I

        If wValorAuxDesc = 0 Then
           MsgBox "Selecione um produto para aplicar o desconto.", vbInformation, "Atenção"
           If rdoControle.status = 0 Then rdoControle.Close
           txtDesconto.SetFocus
           txtDesconto.SelStart = 0
           txtDesconto.SelLength = Len(txtDesconto.Text)
           
           Exit Sub
        Else
            wTotalPedido = (txtTotalPedido.Text - wValorAuxDesc)
        End If


     End If
   Else
   
   
   
   rsComplementoVenda.Close

'**************** DESCONTO GERAL

    wLimiteDescontoAtingido = False
    
    If optValor.Value = True Then
        wTotal = txtTotalPedido.Text
        wDesconto = txtDesconto.Text
        wDesconto = Format(((wDesconto / wTotal) * 100), "###,###,###,##0.000000")
        
        For I = 1 To grdItensPedido.Rows - 1
            'If wDesconto >= (grdItensPedido.TextMatrix(i, 3) / grdItensPedido.TextMatrix(i, 4)) * 100 Then
                'wLimiteDescontoAtingido = True
                'pintaLinhaGrid grdItensPedido, i
                'grdItensPedido.TextMatrix(i, 7) = Format((grdItensPedido.TextMatrix(i, 4) - ((grdItensPedido.TextMatrix(i, 4) * wDesconto) / 100)), "##0.00")
                'grdItensPedido.TextMatrix(i, 6) = Format(((grdItensPedido.TextMatrix(i, 4) * wDesconto) / 100), "##0.00")
            'Else
                grdItensPedido.TextMatrix(I, 6) = Format((grdItensPedido.TextMatrix(I, 3) - ((grdItensPedido.TextMatrix(I, 3) * wDesconto) / 100)), "##0.00")
                grdItensPedido.TextMatrix(I, 5) = Format(((grdItensPedido.TextMatrix(I, 3) * wDesconto) / 100), "##0.00")
                Linha = I
                
                If grdItensPedido.TextMatrix(I, 12) = "N" Or wTipoBloqueio <> "E" Then
                    bloqueioID = True
                End If
            'End If
        Next I
        'If wLimiteDescontoAtingido Then
            'MsgBox "Esse pedido possui item(s) que excedeu o limite máximo do desconto", vbCritical, "Atenção"
            'CarregaItensPedido
            'pintaLinhaGrid grdItensPedido, i
            'rdoControle.Close
            'Exit Sub
        'End If
        
        
        wDif = 0
        wValorAuxDesc = 0
        For I = 1 To grdItensPedido.Rows - 1
           wValorAuxDesc = wValorAuxDesc + grdItensPedido.TextMatrix(I, 5)
        Next I
        wDesconto = Format(txtDesconto.Text, "##0.00")
        wDif = Format(wDesconto - wValorAuxDesc, "#0.00")
        
    Else
        wValorAuxDesc = 0
        For I = 1 To grdItensPedido.Rows - 1
             grdItensPedido.TextMatrix(I, 6) = Format((grdItensPedido.TextMatrix(I, 3) - ((grdItensPedido.TextMatrix(I, 3) * txtDesconto.Text) / 100)), "##0.00")
             grdItensPedido.TextMatrix(I, 5) = Format(((grdItensPedido.TextMatrix(I, 3) * txtDesconto) / 100), "##0.00")
             wValorAuxDesc = wValorAuxDesc + grdItensPedido.TextMatrix(I, 5)
             Linha = I
             If grdItensPedido.TextMatrix(I, 12) = "N" Or wTipoBloqueio <> "E" Then
                  bloqueioID = True
             End If
        Next I
        
        wDesconto = wValorAuxDesc
    End If

       If wDif <> 0 Then
          grdItensPedido.TextMatrix(Linha, 5) = Format((grdItensPedido.TextMatrix(Linha, 5) + (wDif)), "##0.00")
          grdItensPedido.TextMatrix(Linha, 6) = Format((grdItensPedido.TextMatrix(Linha, 6) - (wDif)), "##0.00")
       End If


    wTotalPedido = 0
    wTotalPedido = Format((txtTotalPedido.Text - wDesconto), "###,###,###,##0.00")
  
  End If
  
   txtTotalGeral.Text = Format(wTotalPedido, "###,###,###,##0.00")
   If optPercentual.Value = True Then
      wDescontoMax = rdoControle("CTS_DescontoVendedor")
   Else
      If optValor.Value = True Then
         wDescontoMax = ((txtTotalPedido * rdoControle("CTS_DescontoVendedor")) / 100)
      End If
   End If
   
   GLB_Senha = Trim(rdoControle("CTS_SenhaDesconto"))
   
If ChcSimula.Value = 0 Then
   If wLiberaBloqueio = False Then
        If validaDesconto(grdItensPedido) = True Then
            senhaGerente
        End If
   Else
        If bloqueioID Then
            MsgBox "Há item(s) totalmente bloqueado nesse pedido", vbExclamation
        Else
            senhaGerente
        End If
   End If
End If
   rdoControle.Close

   Call validaTotal(grdItensPedido)
   
'  rsComplementoVenda.Close
  'End If
  
ElseIf KeyAscii = 46 Then
   txtDesconto.SelStart = 0
   txtDesconto.SelLength = Len(txtDesconto.Text)
   txtDesconto.Text = ""
   txtDesconto.SetFocus

ElseIf KeyAscii = 27 Then
   Unload Me
End If
If KeyAscii = 13 Then
        liquidotodas
    liquidolinha
End If
 
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If UCase(txtSenha.Text) = UCase(GLB_Senha) Then
      cmdGrava.Enabled = True
      ChcSimula.Enabled = False
      txtTotalGeral.Visible = True
      FrmeSenha.Visible = False
            'fraFormaDesconto.Enabled = True
            fraTipoDesconto.Enabled = True
            grdItensPedido.Enabled = True
      wTotal = 0
'      Esperar 1
      txtDesconto.Enabled = False
      fraTipoDesconto.Enabled = False
      'fraFormaDesconto.Enabled = False
      grdItensPedido.Enabled = False
      For I = 1 To grdItensPedido.Rows - 1
         wTotal = wTotal + grdItensPedido.TextMatrix(I, 6)
      Next I
      
   Else
      txtSenha.SetFocus
      txtSenha.SelStart = 0
      txtSenha.SelLength = Len(txtSenha.Text)
   End If
   txtTotalGeral.Text = Format(wTotalPedido, "###,###,###,##0.00")
    
ElseIf KeyAscii = 27 Then
   txtDesconto.Enabled = True
   fraTipoDesconto.Enabled = True
   'fraFormaDesconto.Enabled = True
   grdItensPedido.Enabled = True
   txtTotalGeral.Visible = True
   FrmeSenha.Visible = False
      'fraFormaDesconto.Enabled = True
      fraTipoDesconto.Enabled = True
      grdItensPedido.Enabled = True
   txtDesconto.Text = ""
   txtDesconto.SetFocus
    
   For I = 1 To grdItensPedido.Rows - 1
      grdItensPedido.TextMatrix(I, 6) = grdItensPedido.TextMatrix(I, 3)
      grdItensPedido.TextMatrix(I, 4) = "0,00"
   Next I
   txtTotalGeral.Text = txtTotalPedido.Text
   Exit Sub
End If


End Sub

Private Sub txtSenha_LostFocus()
txtSenha.Text = ""
'If GetAsyncKeyState(vbKeyTab) <> 0 Then
'   txtSenha.Text = ""
'   FrmeSenha.Visible = False
'   txtTotalGeral.Visible = True
'   txtDesconto.Enabled = True
'   grdItensPedido.Enabled = True
'   For i = 1 To grdItensPedido.Rows - 1
'      grdItensPedido.TextMatrix(i, 7) = grdItensPedido.TextMatrix(i, 4)
'      grdItensPedido.TextMatrix(i, 6) = "0,00"
'   Next i
'   txtDesconto.SetFocus
'   Exit Sub
'End If

End Sub

' Private Sub txtTotalGeral_KeyPress(KeyAscii As Integer)
' cmdGrava.SetFocus
' End Sub
Private Sub LerControle()
Sql = "Select * from ControleSistema"
      rdoControle.CursorLocation = adUseClient
      rdoControle.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic

End Sub

Private Sub CarregaItensPedido()

    Dim rsComplementoVenda As New ADODB.Recordset
 
    grdItensPedido.Rows = 1
    
    Sql = "Select i.Referencia, PR_Descricao, i.VLUNIT,pr_customedioliquido1,pr_precovendaliquido1, i.desconto, i.vltotitem, i.Qtde, " _
          & "pr_classe, c.liberabloqueio, i.DESCRAT AS CP_DESCONTO, CP_PermitirDesconto as PermitirDesconto " _
          & "From NFItens as i, NFCapa as c, ProdutoLoja, CondicaoPagamento " _
          & "Where PR_Referencia = i.Referencia and i.TipoNota = 'PD' and " _
          & "i.NumeroPed = " & txtpedido.Text & " and c.numeroped = i.numeroped and " _
          & "cp_id = pr_indicepreco and c.condpag = cp_codigo " _
          & "Order By Referencia"
''''VOLTA CONTINENTAL
''
''    SQL = "Select i.Referencia, PR_Descricao, i.VLUNIT,pr_customedioliquido1,pr_precovendaliquido1, i.desconto, i.vltotitem, i.Qtde, " _
''          & "pr_classe, 'N' as liberabloqueio, cp_desconto " _
''          & "From NFItens as i, NFCapa as c, ProdutoLoja, CondicaoPagamento " _
''          & "Where PR_Referencia = i.Referencia and i.TipoNota = 'PD' and " _
''          & "i.NumeroPed = " & txtpedido.Text & " and c.numeroped = i.numeroped and " _
''          & "cp_id = pr_indicepreco and c.condpag = cp_codigo " _
''          & "Order By Referencia"

    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    If rsComplementoVenda.EOF = False Then
        
       wTipoBloqueio = rsComplementoVenda("liberabloqueio")
       If (rsComplementoVenda("liberabloqueio") = "S" Or rsComplementoVenda("liberabloqueio") = "T" Or rsComplementoVenda("liberabloqueio") = "E") Then wLiberaBloqueio = True
        wTipoBloqueio = rsComplementoVenda("liberabloqueio")
        
       Do While Not rsComplementoVenda.EOF
         If rsComplementoVenda("PR_Classe") = "P" And ChcSimula.Value = 0 Then
            If (wTipoBloqueio = "S" Or wTipoBloqueio = "T" Or wTipoBloqueio = "E") Then
                wComDesconto = "N"
            Else
                wComDesconto = "P"
            End If
         ElseIf rsComplementoVenda("desconto") > 0 Then
            wComDesconto = "S"
         Else
            wComDesconto = "N"
         End If
       
          grdItensPedido.AddItem wComDesconto & _
                Chr(9) & rsComplementoVenda("Referencia") & _
                Chr(9) & rsComplementoVenda("Qtde") & _
                Chr(9) & Format(rsComplementoVenda("vltotitem"), "###,###,###,##0.00") & _
                Chr(9) & Format(rsComplementoVenda("PR_Descricao"), "###,###,###,##0.00") & _
                Chr(9) & Format(rsComplementoVenda("desconto"), "###,###,###,##0.00") & _
                Chr(9) & Format((rsComplementoVenda("vltotitem") - rsComplementoVenda("desconto")), "###,###,###,##0.00") & _
                Chr(9) & rsComplementoVenda("PR_Classe") & _
                Chr(9) & Format(rsComplementoVenda("pr_precovendaliquido1"), "###,###,###,##0.00") & _
                Chr(9) & Format(rsComplementoVenda("pr_customedioliquido1"), "###,###,###,##0.00") & _
                Chr(9) & Format(rsComplementoVenda("cp_desconto"), "###,###,###,##0.00") & _
                Chr(9) & Format(rsComplementoVenda("cp_desconto"), "###,###,###,##0.00") & _
                Chr(9) & rsComplementoVenda("PermitirDesconto")


           'if rsComplementoVenda("cp_desconto") > (rsComplementoVenda("cp_desconto"))

'          grdItensPedido.AddItem wComDesconto & Chr(9) & rsComplementoVenda("Referencia") & Chr(9) & rsComplementoVenda("Qtde") & _
'                Chr(9) & Format(rsComplementoVenda("vlunit"), "###,###,###,##0.00") & _
'                Chr(9) & Format(rsComplementoVenda("vltotitem"), "###,###,###,##0.00") & _
'                Chr(9) & Format(rsComplementoVenda("PR_Descricao"), "###,###,###,##0.00") & _
'                Chr(9) & Format(rsComplementoVenda("desconto"), "###,###,###,##0.00") & _
'                Chr(9) & Format((rsComplementoVenda("vltotitem") - rsComplementoVenda("desconto")), "###,###,###,##0.00") & _
'                Chr(9) & rsComplementoVenda("PR_Classe")

          rsComplementoVenda.MoveNext
       Loop
'
'    Else
'       MsgBox "Nenhum produto encontrado para este pedido.", vbInformation, "Atenção"
'       rsComplementoVenda.Close
'       txtTotalGeral.Enabled = False
'       txtDesconto.Enabled = False
'
'       Exit Sub
    End If
    rsComplementoVenda.Close
    
 '   For i = 1 To grdItensPedido.Rows - 1
 '     grdItensPedido.TextMatrix(i, 0) = "S"
 '   Next i
'    lblDescricao.Caption = ""
    txtDescricao.Caption = "Descrição do item selecionado"
    liquidotodas
'grdItensPedido.CellForeColor
End Sub

Private Sub pintaLinhaGrid(grid, linhaGrid As Integer)
    Dim I As Integer
    
    grid.Row = linhaGrid
    For I = 0 To grid.Cols - 1
        grid.Col = I
        grid.CellForeColor = vbRed
    Next I
    grid.Row = 0
    grid.Col = 0
    
End Sub

Private Function validaDesconto(grid) As Boolean
    Dim I As Integer
    Dim wValorMaximo As Double
    Dim wValorAplicado As Double
    
    validaDesconto = True
    
    For I = 1 To grid.Rows - 1
    
        wValorMaximo = (grid.TextMatrix(I, 10) * grid.TextMatrix(I, 3)) / 100
        wValorAplicado = grid.TextMatrix(I, 5)
    
        If wValorAplicado > wValorMaximo Then
            validaDesconto = False
            pintaLinhaGrid grid, I
            'grid.TextMatrix(i, 7) = Format((grid.TextMatrix(i, 4) - ((grid.TextMatrix(i, 4) * wDesconto) / 100)), "##0.00")
            'grid.TextMatrix(i, 6) = Format(((grid.TextMatrix(i, 4) * wDesconto) / 100), "##0.00")
        End If
    Next I
    
    If validaDesconto = False Then
        MsgBox "Esse pedido possui item(s) que excedeu o limite máximo do desconto", vbCritical, "Atenção"
        Call CarregaItensPedido
        
        'frmPedido.cmdTotalPedido.Caption = Format(rsDesconto("vltotitem") + GBL_Frete, "###,###,###,##0.00")
        'txtTotalPedido.Text = Format(rsDesconto("vltotitem"), "###,###,###,##0.00")
        txtTotalGeral.Text = txtTotalPedido.Text
        txtpedido.Text = frmPedido.txtpedido.Text
        cmdGrava.Enabled = False
        
        'txtDesconto.SetFocus
        campoSelecionadoComCaracter txtDesconto
        
    End If

End Function
Private Sub senhaGerente()
    If txtDesconto.Text <> "" Then
        If txtDesconto.Text > wDescontoMax Or senhaGerenteItem() = True Then
        FrmeSenha.left = 115
        FrmeSenha.Visible = True
            fraTipoDesconto.Enabled = False
            grdItensPedido.Enabled = False
        txtDesconto.Enabled = False
        FrmeSenha.top = 670
        txtSenha.SetFocus
        ElseIf ChcSimula.Value = 0 Then
        cmdGrava.Enabled = True
        ChcSimula.Enabled = False
        cmdGrava.SetFocus
        End If
    
    End If
End Sub

Private Function senhaGerenteItem() As Boolean

    Dim wTotalLiquido As Double
    Dim wMaximoDesconto As Double

    senhaGerenteItem = False
    'rdoControle ("CTS_DescontoVendedor")
    
    For I = 1 To grdItensPedido.Rows - 1
        wMaximoDesconto = ((grdItensPedido.TextMatrix(I, 3) * rdoControle("CTS_DescontoVendedor")) / 100)
        wTotalLiquido = grdItensPedido.TextMatrix(I, 5)
        
        If wTotalLiquido > wMaximoDesconto Then
            senhaGerenteItem = True
            'pintaLinhaGrid grdItensPedido, i
        End If
    Next I
    
'    If senhaGerenteItem = False Then
'        MsgBox "Esse pedido possui item(s) que excedeu o limite máximo do desconto", vbCritical, "Atenção"
'        Call CarregaItensPedido
'
'        txtTotalGeral.Text = txtTotalPedido.Text
'        txtPedido.Text = frmPedido.txtPedido.Text
'        cmdGrava.Enabled = False
'
'        campoSelecionadoComCaracter txtDesconto
'
'    End If
    
End Function


Private Function validaTotal(grid) As Boolean
    Dim I As Integer
    Dim wTotalLiquido As Double
    
    validaTotal = True
    
    For I = 1 To grid.Rows - 1
        wTotalLiquido = grid.TextMatrix(I, 6)
        
        If wTotalLiquido < 0 Then
            validaTotal = False
            pintaLinhaGrid grid, I
        End If
    Next I
    
    If validaTotal = False Then
        MsgBox "Esse pedido possui item(s) que excedeu o limite máximo do desconto", vbCritical, "Atenção"
        Call CarregaItensPedido
        
        txtTotalGeral.Text = txtTotalPedido.Text
        txtpedido.Text = frmPedido.txtpedido.Text
        cmdGrava.Enabled = False

        campoSelecionadoComCaracter txtDesconto
        
    End If

End Function


'Calcula  a  Margem  Linha
Private Sub liquidolinha()
Dim prliquido, prcusto, totaliq, desc, valor As Double
Dim qtde As Integer
prliquido = 0
prcusto = 0
If grdItensPedido.Row <> 0 Then
            
            'Quantidade  de  Itens  Total  da  linha
            qtde = CInt(grdItensPedido.TextMatrix(grdItensPedido.Row, 2))
    
            'Preço Liquido do  Iten  tirado  do  Grid/Tabela-Protudoloja bd-DmaC_lOJA
            If grdItensPedido.TextMatrix(grdItensPedido.Row, 8) <> "" Then ' verefica  se  o  campo não  está Vazio
                 prliquido = CDbl(grdItensPedido.TextMatrix(grdItensPedido.Row, 8))
            End If
    
            'Custo Liquido do  Iten  tirado  do  Grid/Tabela-Protudoloja bd-DmaC_lOJA
            If grdItensPedido.TextMatrix(grdItensPedido.Row, 9) <> "" Then ' verefica  se  o  campo não  está  Vazio
                prcusto = CDbl(grdItensPedido.TextMatrix(grdItensPedido.Row, 9))
            End If
            
            'Preço  Total  da venda por  Linha
            totaliq = CDbl(grdItensPedido.TextMatrix(grdItensPedido.Row, 6))
            
             If (wTipoBloqueio = "S" Or wTipoBloqueio = "N" Or wTipoBloqueio = "E") Then
                    If txtDesconto.Text <> "" And optPercentual.Value = True Then 'verefica  se o  campo não  está vizio e o Option do  Prescentual  esta  true
                        desc = (prliquido * txtDesconto.Text) / 100
                        prliquido = prliquido - desc

                    ElseIf optValor.Value = True And txtDesconto.Text <> "" Then 'verefica  se o  campo não  está vizio e o Option dovalor  está  true
                            
                           desc = (grdItensPedido.TextMatrix(grdItensPedido.Row, 5) * 100) / grdItensPedido.TextMatrix(grdItensPedido.Row, 3)
                           desc = (prliquido * desc) / 100
                           prliquido = prliquido - desc
                    End If
                    
            End If
        If totaliq <> 0 Then
            valor = ((((qtde * prliquido) - (qtde * prcusto)) * 100) / totaliq)
            If ChcSimula.Value = 1 Then 'verefica  se  e  Simulação
            lblMagenLinha.Visible = False
                lblMagenLinha.Caption = Format(valor, "###,###,###,##0.00")
                lblMagenLinha.Visible = True
            Else
                lblMagenLinha.Visible = False
            End If
        End If
    End If
    
End Sub

'Calcula  Margem  no  Grid  Todo
Private Sub liquidotodas()

Dim prliquido, prcusto, totaliq, desc, valor, Tliqpre, Tliqcus As Double
Dim qtde, Linhas As Integer
prliquido = 0
prcusto = 0
    Tliqpre = 0
    Tliqcus = 0
    Linhas = grdItensPedido.Rows - 1
    
    'Loop  Para  ler todas as Linha  o  Grid
    I = 1
    Do While (I <= Linhas)
                
                 'Quantidade  de  Itens  Total  da  linha
                qtde = CInt(grdItensPedido.TextMatrix(I, 2))
        
                'Preço Liquido do  Iten  tirado  do  Grid/Tabela-Protudoloja bd-DmaC_lOJA
                If grdItensPedido.TextMatrix(I, 8) <> "" Then
                     prliquido = CDbl(grdItensPedido.TextMatrix(I, 8))
                End If
                
                'Custo Liquido do  Iten  tirado  do  Grid/Tabela-Protudoloja bd-DmaC_lOJA
                If grdItensPedido.TextMatrix(I, 9) <> "" Then
                    prcusto = CDbl(grdItensPedido.TextMatrix(I, 9))
                End If
                    
                  'Preço  Total  da venda por  Linha
                  totaliq = CDbl(grdItensPedido.TextMatrix(I, 6))
            
                    
                If txtDesconto.Text <> "" And optPercentual.Value = True And (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then 'verefica  se o  campo não  está vizio e o Option do  Prescentual  esta  true
                        desc = (prliquido * txtDesconto.Text) / 100
                        prliquido = prliquido - desc

                    ElseIf optValor.Value = True And txtDesconto.Text <> "" And (grdItensPedido.TextMatrix(I, 0) = "S" Or grdItensPedido.TextMatrix(I, 0) = "T") Then 'verefica  se o  campo não  está vizio e o Option dovalor  está  true
                            
                           desc = (grdItensPedido.TextMatrix(I, 5) * 100) / grdItensPedido.TextMatrix(I, 3)
                           desc = (prliquido * desc) / 100
                           prliquido = prliquido - desc
                    End If
                   
              grdItensPedido.TextMatrix(I, 13) = Format((((qtde * prliquido) - (qtde * prcusto)) * 100) / totaliq, "###,###,###,##0.00")
              Tliqpre = (Tliqpre + (qtde * prliquido))
              Tliqcus = (Tliqcus + (qtde * prcusto))
    I = I + 1
    Loop
    
    
    valor = (((Tliqpre - Tliqcus) * 100) / txtTotalGeral.Text)

If ChcSimula.Value = 1 Then
 lblMagenTotal.Caption = Format(valor, "###,###,###,##0.00")
 lblMagenTotal.Visible = True
 Else
  lblMagenTotal.Visible = False
 End If
 

End Sub





