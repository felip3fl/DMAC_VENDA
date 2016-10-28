VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmConsultaPedido 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Consulta Pedido"
   ClientHeight    =   5595
   ClientLeft      =   6990
   ClientTop       =   4080
   ClientWidth     =   6555
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtVendedor 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   150
      MaxLength       =   3
      TabIndex        =   0
      Top             =   885
      Width           =   1605
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   2
      Top             =   4875
      Width           =   6165
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdPedidos 
      Height          =   2910
      Left            =   150
      TabIndex        =   1
      Top             =   1440
      Width           =   2580
      _cx             =   4551
      _cy             =   5133
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsultaPedido.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   2
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
   Begin VB.Label lblPagamento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   6165
   End
   Begin VB.Label lblNovoPedido 
      BackColor       =   &H00B63C18&
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
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
      Left            =   150
      TabIndex        =   4
      Top             =   4515
      Width           =   2595
   End
   Begin VB.Label lblVendedor 
      BackColor       =   &H00B63C18&
      BackStyle       =   0  'Transparent
      Caption         =   " Vendedor"
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
      Left            =   150
      TabIndex        =   3
      Top             =   630
      Width           =   1095
   End
End
Attribute VB_Name = "frmConsultaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim wNroPedido As Long
Dim wNroPedidoNovo As Long

Private Sub cmdRetorna_Click()
 'If grdPedidos.Visible = True Then
    grdPedidos.Visible = False
    lblVendedor.Visible = True
    txtVendedor.Visible = True
    txtVendedor.SetFocus
  'Else
   Unload Me
   frmPedido.txtPedido.SetFocus
  'End If
End Sub

Private Sub Form_Load()
   Call AjustaTela(frmConsultaPedido)
     lblNovoPedido.Caption = ""
     grdPedidos.Visible = False
     lblVendedor.Visible = True
     txtVendedor.Visible = True
End Sub


Private Sub grdPedidos_DblClick()
  If MsgBox("Deseja Replicar este Pedido = " & grdPedidos.TextMatrix(grdPedidos.Row, 0), _
     vbYesNo + vbQuestion, "Atenção") = vbYes Then
     SQL = "Select * from ControleSistema"
     rsPegaNumeroPedido.CursorLocation = adUseClient
     rsPegaNumeroPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

     If Not rsPegaNumeroPedido.EOF Then
        adoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        SQL = ""
        SQL = "Update ControleSistema set CTS_NumeroPedido=(CTS_NumeroPedido + 1)"
              adoCNLoja.Execute SQL
              Screen.MousePointer = vbNormal
              adoCNLoja.CommitTrans
     
        lblNovoPedido.Caption = "Novo Pedido : " & rsPegaNumeroPedido("CTS_NumeroPedido")
        wNroPedidoNovo = rsPegaNumeroPedido("CTS_NumeroPedido")
        wNroPedido = grdPedidos.TextMatrix(grdPedidos.Row, 0)
        rsPegaNumeroPedido.Close
        adoCNLoja.Execute "Exec SP_Replicar_Pedido_Venda " & wNroPedido & "," & wNroPedidoNovo

        Exit Sub
     Else
        MsgBox "Erro no Controle do Sistema avise o CPD"
        rsPegaNumeroPedido.Close
     End If
   End If
End Sub

Private Sub grdPedidos_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 27 Then
     grdPedidos.Row = 1
     grdPedidos.Visible = False
     lblVendedor.Visible = True
     txtVendedor.Visible = True
     txtVendedor.SetFocus
     cmdRetorna_Click

  End If
     
     
End Sub

Private Sub txtVendedor_GotFocus()
    txtVendedor.SelStart = 0
    txtVendedor.SelLength = Len(txtVendedor.Text)
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   If txtVendedor.Text = "" Then
     MsgBox "Informe o número do vendedor"
     Exit Sub
   End If
 
   SQL = "select ve_codigo from vende where ve_codigo = " & RTrim(LTrim(txtVendedor.Text))
   rsVendedor.CursorLocation = adUseClient
   rsVendedor.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
    If Not rsVendedor.EOF Then
        
      'lblVendedor.Visible = False
      'txtVendedor.Visible = False
      grdPedidos.Visible = True
      grdPedidos.Rows = 1
       
      SQL = "select numeroped,totalnota from nfcapa where tiponota = 'PA' and dataemi = '" & Format(Date, "yyyy/mm/dd") & _
      "' and vendedor = " & RTrim(LTrim(txtVendedor.Text))
      rsPedidosAbertos.CursorLocation = adUseClient
      rsPedidosAbertos.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic


      If Not rsPedidosAbertos.EOF Then
         Do While Not rsPedidosAbertos.EOF
           grdPedidos.AddItem rsPedidosAbertos("Numeroped") & Chr(9) & _
           Format(rsPedidosAbertos("totalnota"), "###,###,###,##0.00")
           rsPedidosAbertos.MoveNext
         Loop
      End If
      rsPedidosAbertos.Close

    Else
       MsgBox "Vendedor não cadastrado."
    End If
    rsVendedor.Close
  End If
  
  
  If KeyAscii = 27 Then
   Unload Me
   frmPedido.txtPedido.SetFocus
  End If
End Sub


