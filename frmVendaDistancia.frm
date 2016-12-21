VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmVendaDistancia 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Venda a Distância"
   ClientHeight    =   5535
   ClientLeft      =   1785
   ClientTop       =   2250
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   5
      Top             =   4875
      Width           =   6165
   End
   Begin VB.TextBox txtVendedorLojaVenda 
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
      MaxLength       =   3
      TabIndex        =   4
      Top             =   4245
      Width           =   2265
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5805
      OleObjectBlob   =   "frmVendaDistancia.frx":0000
      Top             =   3840
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
      Left            =   5505
      TabIndex        =   0
      Top             =   3855
      Visible         =   0   'False
      Width           =   300
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdLojas 
      Height          =   3450
      Left            =   150
      TabIndex        =   1
      Top             =   405
      Width           =   6345
      _cx             =   11192
      _cy             =   6085
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmVendaDistancia.frx":0234
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
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5250
      TabIndex        =   6
      Top             =   5055
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
      MICON           =   "frmVendaDistancia.frx":0281
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblVendedorLojaVenda 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor Loja Venda"
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
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   3990
      Width           =   2460
   End
   Begin VB.Label lblLojaVenda 
      BackStyle       =   0  'Transparent
      Caption         =   "Loja Venda"
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
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   2775
   End
End
Attribute VB_Name = "frmVendaDistancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCodigo As Integer
Dim wSequencia As Integer
Dim wValorCampo As String
Dim SQL As String
Dim status As Integer
Dim rsTipoCliente As New ADODB.Recordset


Private Sub cmdGrava_Click()
    
  wSequencia = 10
  wValorCampo = txtVendedorLojaVenda.Text
  SQL = ""
  SQL = "LojaVenda = ''" & Trim(grdLojas.TextMatrix(grdLojas.Row, 0)) & "'', OutraLoja = ''" & Trim(grdLojas.TextMatrix(grdLojas.Row, 0)) & "''," & _
        "VendedorLojaVenda = " & wValorCampo & ",OutroVend = " & wValorCampo
  adoCNLoja.Execute "exec SP_GravaComplementoVenda " & txtPedido.Text & ",1," & wSequencia & ",'" & SQL & "'" ', rdExecDirect

  Unload Me
  
  
End Sub

Private Sub cmdRetorna_Click()
 Unload Me
 frmPedido.txtPesquisar.SetFocus

End Sub

Private Sub Form_Activate()
'ricardo
  VerificaCliente
End Sub

Private Sub Form_Load()
    
    Call AjustaTela(frmVendaDistancia)
    
  'Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
 ' Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
  


  txtVendedorLojaVenda.Enabled = False

  txtPedido.Text = frmPedido.txtPedido.Text
  wCodigo = 1
  SQL = "Select * from Loja where " _
      & "LO_OrdemLoja <> 888 and LO_Loja not in('Conso','CMCS','CMCE','CD','CMC') Order By LO_OrdemLoja"
  
  rsCarregaLoja.CursorLocation = adUseClient
  rsCarregaLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  If Not rsCarregaLoja.EOF Then
     Do While Not rsCarregaLoja.EOF
        grdLojas.AddItem rsCarregaLoja("LO_Loja") & Chr(9) _
        & rsCarregaLoja("LO_Endereco")
        rsCarregaLoja.MoveNext
     Loop
  End If

  rsCarregaLoja.Close
End Sub
Private Function VerificaCliente() As Boolean

    'Metodo que verifca se o cliente é do tipo consumidor
   VerificaCliente = True
   
    SQL = ""
    SQL = "select * from nfcapa where numeroped = '" & (frmPedido.txtPedido.Text) & "' and vendedor = '" & Mid(frmPedido.txtVendedor.Text, 1, 3) & "' and LojaOrigem = '" & RTrim(wLoja) & "'"
      

        rsTipoCliente.CursorLocation = adUseClient
        rsTipoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        
            If rsTipoCliente("Cliente") = "999999" Then
                MsgBox "Você não pode fazer venda a distância para Cliente consumidor", vbInformation, "Atenção"
                 VerificaCliente = False
                 Unload Me
                  frmPedido.txtPesquisar.SetFocus
            End If
            
    
        rsTipoCliente.Close
    
End Function

Private Sub grdLojas_Click()
 '   fraVendedor.Enabled = True
    'cmdRetorna.Enabled = True
    txtVendedorLojaVenda.Enabled = True
    txtVendedorLojaVenda.SetFocus
'    cmdGrava.Enabled = True
  
End Sub

Private Sub grdLojas_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
       Unload Me
 frmPedido.txtPesquisar.SetFocus
   End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub txtVendedorLojaVenda_Change()

    If IsNumeric(txtVendedorLojaVenda.Text) = False Then
        txtVendedorLojaVenda.Text = ""
        txtVendedorLojaVenda.SelStart = 0
        txtVendedorLojaVenda.SelLength = Len(txtVendedorLojaVenda.Text)
        txtVendedorLojaVenda.SetFocus
    ElseIf txtVendedorLojaVenda.Text <= 0 Then
        txtVendedorLojaVenda.Text = ""
        txtVendedorLojaVenda.SelStart = 0
        txtVendedorLojaVenda.SelLength = Len(txtVendedorLojaVenda.Text)
        txtVendedorLojaVenda.SetFocus
    End If


End Sub

Private Sub txtVendedorLojaVenda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      If txtVendedorLojaVenda <> "" Then
        cmdGrava.Enabled = True
        cmdGrava.SetFocus
        Exit Sub
      Else
 '       txtVendedorLojaVenda.SetFocus
        cmdGrava.Enabled = False
        Exit Sub
      End If
    End If
       
    
If KeyAscii = 46 Then
      txtVendedorLojaVenda.Text = 0
      txtVendedorLojaVenda.SelStart = 0
      txtVendedorLojaVenda.SelLength = Len(txtVendedorLojaVenda.Text)
      txtVendedorLojaVenda.SetFocus
      Exit Sub
   End If
   
   If KeyAscii = 44 Then
      txtVendedorLojaVenda.Text = 0
      txtVendedorLojaVenda.SelStart = 0
      txtVendedorLojaVenda.SelLength = Len(txtVendedorLojaVenda.Text)
      txtVendedorLojaVenda.SetFocus
      Exit Sub
   End If
   

   
If KeyAscii = 27 Then
    Unload Me
 frmPedido.txtPesquisar.SetFocus
 ElseIf KeyAscii = vbKeyF2 Then
 cmdGrava_Click
 
End If
End Sub

Private Sub txtVendedorLojaVenda_LostFocus()

    If IsNumeric(txtVendedorLojaVenda.Text) = False Then
        txtVendedorLojaVenda.Text = ""
        txtVendedorLojaVenda.SelStart = 0
        txtVendedorLojaVenda.SelLength = Len(txtVendedorLojaVenda.Text)
        cmdGrava.Enabled = False
 '       txtVendedorLojaVenda.SetFocus
    ElseIf txtVendedorLojaVenda.Text <= 0 Then
        txtVendedorLojaVenda.Text = ""
        txtVendedorLojaVenda.SelStart = 0
        txtVendedorLojaVenda.SelLength = Len(txtVendedorLojaVenda.Text)
'        txtVendedorLojaVenda.SetFocus
        cmdGrava.Enabled = False
    Else
        cmdGrava.Enabled = True
    End If
    
End Sub
