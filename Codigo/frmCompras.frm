VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.OCX"
Begin VB.Form frmCompras 
   BackColor       =   &H00E0E0E0&
   Caption         =   "C o m p r a s"
   ClientHeight    =   3975
   ClientLeft      =   1335
   ClientTop       =   5400
   ClientWidth     =   10125
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   10125
   Begin VB.TextBox TxtPedido 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3675
      Width           =   1395
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   -15
      OleObjectBlob   =   "frmCompras.frx":0000
      Top             =   3750
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensProduto 
      Height          =   3690
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   10110
      _cx             =   17833
      _cy             =   6509
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   0
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
      BackColor       =   16761024
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12582912
      ForeColorSel    =   65535
      BackColorBkg    =   16761024
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCompras.frx":0234
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
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2

  Skin1.LoadSkin "c:\WINDOWS\system\skin.skn"
  Skin1.ApplySkin Me.hwnd
  frmCompras.Height = 4215
End Sub

Function PesquisarProduto(ByVal wWhere As String)
    grdItensProduto.Rows = 1
   ' grdPrecos.Rows = 1
    
    cmdLimpar.Caption = "Pesquisando ..."
    Screen.MousePointer = 11
    SQL = "Select PR_Referencia, " _
        & "PR_Descricao,PR_PrecoVenda1,ESL_Estoque, PR_Grupo, PR_Classe, PR_Bloqueio, LI_Descricao, SC_Descricao," _
        & "GP_Descricao from Produto,Produtobarras,EstoqueLoja,Linha,Secao,GrupoProduto " _
        & "where ESL_Referencia=PR_Referencia and LI_CodigoLinha=PR_Linha and SC_CodigoLinha=PR_Linha and " _
        & "SC_CodigoSecao=PR_Secao and GP_CodigoLinha=PR_Linha and GP_CodigoSecao=PR_Secao and " _
        & "GP_CodigoGrupo=PR_Grupo and " & wWhere & " and PR_Situacao not in('E') and PRB_Referencia = PR_Referencia" _
        & " order by PR_CodigoFornecedor,PR_Descricao"
            
    rsPesquisaPed.CursorLocation = adUseClient
    rsPesquisaPed.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsPesquisaPed.EOF Then
       ' grdDadosProduto.Enabled = True
       ' grdPrecos.Enabled = True
       AuxProdutoExiste = True
       grdPrecos.Redraw = True
        txtPesquisar.Text = rsPesquisaPed("PR_Descricao")
        'txtQuantidade.Enabled = True
        grdItensProduto.Redraw = False
        ReferenciaPreco = rsPesquisaPed("PR_PrecoVenda1")
        Do While Not rsPesquisaPed.EOF
            grdItensProduto.AddItem rsPesquisaPed("PR_Referencia") & Chr(9) _
                & rsPesquisaPed("PR_Descricao") & Chr(9) _
                & rsPesquisaPed("ESL_Estoque") & Chr(9) _
                & Format(rsPesquisaPed("PR_PrecoVenda1"), "0.00") & Chr(9) _
                & Format(rsPesquisaPed("PR_AliquotaIPI"), "0.00") & Chr(9) _
                & rsPesquisaPed("PRB_CodigoBarras") & Chr(9) _
                & rsPesquisaPed("PR_Linha") & " - " & rsPesquisaPed("LI_Descricao") & Chr(9) _
                & rsPesquisaPed("PR_Secao") & " - " & rsPesquisaPed("SC_Descricao") & Chr(9) _
                & rsPesquisaPed("PR_Grupo") & " - " & rsPesquisaPed("GP_Descricao") & Chr(9) _
                & rsPesquisaPed("PR_Classe") & Chr(9) _
                & rsPesquisaPed("PR_Bloqueio") & Chr(9) _
                & Format(rsPesquisaPed("PR_IcmsSaida"), "0.00") & Chr(9) _
                & Format(rsPesquisaPed("PR_IcmPdv"), "0.00") & Chr(9)
                wValorVenda = Format(rsPesquisaPed("PR_PrecoVenda1"), "0.00")
            rsPesquisaPed.MoveNext
        Loop
        grdItensProduto.Enabled = True
        grdItensProduto.Redraw = True
        grdItensProduto.SetFocus
        grdItensProduto.Row = 1
    Else
        AuxProdutoExiste = False
        grdDadosProduto.Rows = 1
        grdDadosProduto.Rows = 2
        grdPrecos.Rows = 1
        grdPrecos.Rows = 2
        'grdDadosProduto.Enabled = False
        'grdPrecos.Enabled = False
        'grdPrecos.Redraw = False
        
        txtPesquisar.Text = (txtPesquisar.Text & "     " & "Nenhum Registro Encontrado")
        'txtPesquisar.Text = "Nenhum Registro Encontrado"
        txtQuantidade.Enabled = False
        txtPesquisar.Enabled = True
        txtPesquisar.SetFocus
        txtPesquisar.SelStart = 0
        txtPesquisar.SelLength = Len(txtPesquisar.Text)
        txtPesquisar.SetFocus
    End If
    rsPesquisaPed.Close
    Screen.MousePointer = 0
End Function

