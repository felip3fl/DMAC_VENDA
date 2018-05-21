VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmMovimentacaoEstoque 
   BackColor       =   &H00404040&
   Caption         =   "Movimentação Estoque"
   ClientHeight    =   7605
   ClientLeft      =   345
   ClientTop       =   3180
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   13830
   Begin VB.Frame fraGrupoDocumento 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   3870
      TabIndex        =   5
      Top             =   105
      Width           =   9780
      Begin VB.ComboBox cmbLoja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   990
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdCadastraFornecedor 
         Height          =   510
         Left            =   1380
         TabIndex        =   9
         Top             =   180
         Width           =   6870
         _cx             =   12118
         _cy             =   900
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   3158064
         ForeColor       =   12632256
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16423203
         ForeColorSel    =   8388608
         BackColorBkg    =   5263440
         BackColorAlternate=   3947580
         GridColor       =   5263440
         GridColorFixed  =   8421504
         TreeColor       =   3947580
         FloodColor      =   5263440
         SheetBorder     =   3947580
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMovimentacaoEstoque.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   5
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
         BackColorFrozen =   5263440
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin VB.Label lblLoja 
         BackColor       =   &H00505050&
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   315
         Left            =   135
         TabIndex        =   7
         Top             =   105
         Width           =   660
      End
   End
   Begin VB.Frame fraTipoPesquisa 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FA9923&
      Height          =   780
      Left            =   45
      TabIndex        =   1
      Top             =   105
      Width           =   3780
      Begin VB.TextBox txtPesquisa 
         Height          =   315
         Left            =   2265
         TabIndex        =   11
         Top             =   360
         Width           =   1440
      End
      Begin VB.OptionButton optDV 
         BackColor       =   &H00505050&
         Caption         =   "DV"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1275
         TabIndex        =   10
         Top             =   510
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optTodos 
         BackColor       =   &H00505050&
         Caption         =   "Todos"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1275
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
      Begin VB.OptionButton optFornecedor 
         BackColor       =   &H00505050&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   510
         Width           =   1140
      End
      Begin VB.OptionButton optReferencia 
         BackColor       =   &H00505050&
         Caption         =   "Referencia"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   30
         TabIndex        =   2
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00505050&
         Caption         =   "Pesquisa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   210
         Left            =   2265
         TabIndex        =   12
         Top             =   105
         Width           =   1515
      End
      Begin VB.Label lblTipoDocumento 
         BackColor       =   &H00505050&
         Caption         =   "Tipo Pesquisa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   210
         Left            =   15
         TabIndex        =   4
         Top             =   0
         Width           =   1515
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdMovimentacaoEstoque 
      Height          =   5760
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   13620
      _cx             =   24024
      _cy             =   10160
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   3158064
      ForeColor       =   12632256
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16423203
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   8421504
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   14
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMovimentacaoEstoque.frx":00D0
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmMovimentacaoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MesesPassados As Variant
Dim adoloja As New ADODB.Recordset
Dim adoGrid As New ADODB.Recordset
Dim Col, i As Long
Dim sql1, sql2, Mes As String

Private Sub Form_Load()
    
    
  grdMovimentacaoEstoque.MergeRow(0) = True
  grdMovimentacaoEstoque.MergeRow(1) = True
  grdMovimentacaoEstoque.MergeCol(0) = True
  grdMovimentacaoEstoque.MergeCol(1) = True
  grdMovimentacaoEstoque.MergeCol(2) = True
  grdMovimentacaoEstoque.MergeCol(3) = True
  grdMovimentacaoEstoque.MergeCol(4) = True
  grdMovimentacaoEstoque.MergeCol(5) = True
  grdMovimentacaoEstoque.MergeCol(6) = True
  grdMovimentacaoEstoque.MergeCol(7) = True
  grdMovimentacaoEstoque.MergeCol(8) = True
  grdMovimentacaoEstoque.MergeCol(9) = True
  grdMovimentacaoEstoque.MergeCol(10) = True
  grdMovimentacaoEstoque.MergeCol(11) = True
  grdMovimentacaoEstoque.MergeCol(12) = True
  grdMovimentacaoEstoque.MergeCol(13) = True

 MesesPassados = MesesTendencia
  grdCadastraFornecedor.TextMatrix(0, 0) = MesesPassados(0)
  grdCadastraFornecedor.TextMatrix(0, 1) = MesesPassados(1)
  grdCadastraFornecedor.TextMatrix(0, 2) = MesesPassados(2)
  grdCadastraFornecedor.TextMatrix(0, 3) = MesesPassados(3)
  grdCadastraFornecedor.TextMatrix(0, 4) = MesesPassados(4)
  grdCadastraFornecedor.TextMatrix(0, 5) = MesesPassados(5)
  grdCadastraFornecedor.TextMatrix(0, 6) = MesesPassados(6)
  
  cmbLoja.Clear
  cmbLoja.AddItem ""
        SQL = "Select * from Loja where lo_loja<>'CONSO'"
               adoloja.CursorLocation = adUseClient
            adoloja.Open SQL, ado_cn_dmac, adOpenForwardOnly, adLockPessimistic

          If Not adoloja.EOF Then
        Do While Not adoloja.EOF
            cmbLoja.AddItem adoloja("LO_Loja")
            adoloja.MoveNext
        Loop
            adoloja.Close
        End If
    cmbLoja.ListIndex = 0

End Sub

Private Sub grdTitulosDia_Click()

End Sub


Private Sub grdCadastraFornecedor_Click()
  
      If (grdCadastraFornecedor.CellChecked = 1) Then
    grdCadastraFornecedor.CellChecked = 0
    Mes = ""
  Else
   grdCadastraFornecedor.CellChecked = 1
   Mes = TraduzMesNu(grdCadastraFornecedor.TextMatrix(0, grdCadastraFornecedor.Col))
   End If
   
   For i = 0 To grdCadastraFornecedor.Cols - 1
   If i <> grdCadastraFornecedor.Col Then
        grdCadastraFornecedor.TextMatrix(1, i) = 0
   End If
   Next i
   
Call CarregaGrid
End Sub



Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)

 Select Case KeyAscii
     Case vbKeyDelete
     Case vbKeyBack
     Case 48 To 57
     Case Else
             Beep
             KeyAscii = 0
End Select

    

End Sub

Private Sub CarregaGrid()
 Screen.MousePointer = 11
   If (optReferencia.Value = True) And (txtPesquisa.Text <> "") Then
        sql1 = " and MVE_REFERENCIA='" & Trim(txtPesquisa.Text) & "' "
    
    ElseIf (optFornecedor.Value = True) And (txtPesquisa.Text <> "") Then
        sql1 = " and  PR_CodigoFornecedor='" & Trim(txtPesquisa.Text) & "' "
    
    ElseIf optTodos.Value = True Then
        sql1 = ""
    
    ElseIf optDV.Value = True Then
    
    End If
    
    If cmbLoja.Text <> "" Then
        If cmbLoja.Text <> 999 Then
            sql2 = "  and  MVE_Loja = '" & Trim(cmbLoja.Text) & "'"
        Else
            sql2 = ""
        End If
    Else
            MsgBox ("Loja Não Selecionada")
            sql2 = " and  MVE_Loja = '' "
          
    End If
            


grdMovimentacaoEstoque.Rows = 2


   SQL = "Select MVE_Referencia,pr_descricao,MVE_EstoqueInicial, MVE_Venda,MVE_transfSaida,MVE_outrasSaida,MVE_AjusteSaida,MVE_DevolucaoCompras" _
    & ",MVE_EntradaCompras,MVE_TransfEntrada,MVE_AjusteEntrada,MVE_DevolucaoVenda,MVE_OutrasEntrada,MVE_EstoqueFinal  from Est_movimentacaoestoque " _
    & "inner join produto on pr_referencia=mve_referencia " _
    & "WHERE MVE_MES='" & Mes & "' " _
    & sql1 _
    & sql2
    
    
               adoGrid.CursorLocation = adUseClient
            adoGrid.Open SQL, ado_cn_demeo, adOpenForwardOnly, adLockPessimistic
            If Not adoGrid.EOF Then
        Do While Not adoGrid.EOF
         grdMovimentacaoEstoque.AddItem adoGrid("MVE_Referencia") & Chr(vbKeyTab) _
                                            & adoGrid("pr_descricao") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_EstoqueInicial") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_Venda") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_transfSaida") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_outrasSaida") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_AjusteSaida") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_DevolucaoCompras") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_EntradaCompras") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_TransfEntrada") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_AjusteEntrada") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_DevolucaoVenda") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_OutrasEntrada") & Chr(vbKeyTab) _
                                            & adoGrid("MVE_EstoqueFinal")
                
        
        
        adoGrid.MoveNext
    Loop
    End If
    adoGrid.Close
 Screen.MousePointer = 1
End Sub

