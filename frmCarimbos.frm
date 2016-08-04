VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmCarimbos 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Carimbos Nota Fiscal"
   ClientHeight    =   6510
   ClientLeft      =   8025
   ClientTop       =   3885
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   3
      Top             =   5835
      Width           =   6165
   End
   Begin VB.TextBox TxtCarimbo 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   150
      TabIndex        =   0
      Top             =   390
      Width           =   6165
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   45
      OleObjectBlob   =   "frmCarimbos.frx":0000
      Top             =   4980
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
      Left            =   645
      TabIndex        =   1
      Top             =   5010
      Visible         =   0   'False
      Width           =   300
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdCarimbo 
      Height          =   4500
      Left            =   150
      TabIndex        =   2
      ToolTipText     =   "Se desejar excluir um item, clique duas vezes sobre o item a ser excluido."
      Top             =   900
      Width           =   6165
      _cx             =   10874
      _cy             =   7937
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      FormatString    =   $"frmCarimbos.frx":0234
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
      WallPaperAlignment=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carimbo"
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
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   6165
   End
End
Attribute VB_Name = "frmCarimbos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCarimbo As New ADODB.Recordset
Dim rsCarimbo2 As New ADODB.Recordset
Dim wPesquisaCodigo As Integer
Dim wSEQ As Integer


Private Sub cmdRetorna_Click()
 Unload Me
 frmPedido.txtPesquisar.SetFocus

End Sub

Private Sub cmdRetorna_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   frmPedido.txtPesquisar.SetFocus
End If
Unload Me
End Sub

Private Sub Form_Activate()
  txtPedido.Text = frmPedido.txtPedido.Text
  Call CarregaCarimbos
End Sub

Private Sub Form_Load()
    Call AjustaTela(frmCarimbos)
  
'  Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
   
End Sub

Private Sub CarregaCarimbos()

  grdCarimbo.Rows = 1
  TxtCarimbo.Text = ""
  
  SQL = "Select * From CarimboNotaFiscal " _
        & "Where CNF_NumeroPed = " & txtPedido.Text & " And CNF_TipoCarimbo = 'I' " _
        & "Order By CNF_Sequencia"
  rsCarimbo.CursorLocation = adUseClient
  rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsCarimbo.EOF Then
     Do While Not rsCarimbo.EOF
        grdCarimbo.AddItem rsCarimbo("CNF_sequencia") _
        & Chr(9) & rsCarimbo("cnf_carimbo")
        rsCarimbo.MoveNext
     Loop
     grdCarimbo.Enabled = True
     grdCarimbo.Editable = flexEDNone
     grdCarimbo.Row = 1

  End If
  rsCarimbo.Close
End Sub

Private Sub grdCarimbo_DblClick()

 On Error GoTo erronaUpdate

 If grdCarimbo.Row = 0 Then
    Exit Sub
 Else
    If grdCarimbo.TextMatrix(grdCarimbo.Row, 0) < 1 Then
        Exit Sub
    End If
 End If
 
 If MsgBox("Deseja Excluir o Item = " & grdCarimbo.TextMatrix(grdCarimbo.Row, 0), _
           vbYesNo + vbQuestion, "Atenção") = vbYes Then
    adoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    

    SQL = "Delete CarimboNotaFiscal Where CNF_NumeroPed = " & txtPedido.Text & " and " _
          & "CNF_Sequencia = " & grdCarimbo.TextMatrix(grdCarimbo.Row, 0)
         
         adoCNLoja.Execute (SQL)
         Screen.MousePointer = vbNormal
         adoCNLoja.CommitTrans

    Call CarregaCarimbos
      
    Exit Sub
 Else
    Exit Sub
 End If

erronaUpdate:
MsgBox "Erro na Exclusão Item " & Err.description, vbCritical, "Aviso"
'adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal


End Sub

Private Sub TxtCarimbo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     
     
  
  If TxtCarimbo = "" Then
     MsgBox "Inserir Carimbo"
     Exit Sub
  End If
  
  If Len(TxtCarimbo.Text) > 60 Then
     TxtCarimbo = ""
     MsgBox "Por favor abreviar Carimbo"
     Exit Sub
  Else
     TxtCarimbo.Text = UCase(TxtCarimbo.Text)
  End If
   
   
   SQL = ""
   SQL = "Select max(cnf_Sequencia) as Sequencia From CarimboNotaFiscal Where CNF_NumeroPed = " & txtPedido.Text
   rsCarimbo.CursorLocation = adUseClient
   rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   If IsNull(rsCarimbo("Sequencia")) Then
      auxCarimbo = 2
   Else
      auxCarimbo = rsCarimbo("Sequencia") + 1
   End If
   rsCarimbo.Close
   
   
   SQL = ""
   SQL = "Select LojaOrigem,Serie,nf,numeroped From NFCapa Where NumeroPed = " & txtPedido.Text & " and " _
         & "TipoNota = 'PD'"
   rsCarimbo.CursorLocation = adUseClient
   rsCarimbo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       

    
    SQL = ""
    SQL = "Insert into CarimboNotaFiscal(CNF_NumeroPed,CNF_Loja,CNF_serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo)" _
        & "Values ( " & txtPedido.Text & ",'" & rsCarimbo("LojaOrigem") & _
        "','',0," & auxCarimbo & ",'" & TxtCarimbo.Text & "','I')"
     rsCarimbo2.CursorLocation = adUseClient
     rsCarimbo2.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

rsCarimbo.Close
'rsCarimbo2.Close
 Call CarregaCarimbos
 
End If

 If KeyAscii = 27 Then
    Unload Me
    frmPedido.txtPesquisar.SetFocus
End If

End Sub
