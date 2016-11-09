VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form FrmFichaFinanceira 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   5940
   ClientLeft      =   3315
   ClientTop       =   19890
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   45
      ScaleHeight     =   45
      ScaleWidth      =   15000
      TabIndex        =   5
      Top             =   5250
      Width           =   15000
   End
   Begin VB.CommandButton CmdConsulta 
      BackColor       =   &H00404040&
      Height          =   375
      Left            =   5820
      TabIndex        =   0
      Top             =   6735
      Width           =   600
   End
   Begin Project1.chameleonButton cmdRetorna 
      Height          =   405
      Left            =   13980
      TabIndex        =   2
      Top             =   5340
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Retorna"
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
      MICON           =   "FrmFichaFinanceira.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdFichaFinanceira 
      Height          =   4875
      Left            =   45
      TabIndex        =   3
      Top             =   300
      Width           =   3810
      _cx             =   6720
      _cy             =   8599
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
      Rows            =   17
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmFichaFinanceira.frx":001C
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensProduto 
      Height          =   4890
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Se desejar excluir um item, clique duas vezes sobre o item a ser excluido."
      Top             =   300
      Width           =   11100
      _cx             =   19579
      _cy             =   8625
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmFichaFinanceira.frx":019B
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
   Begin Project1.chameleonButton cmdItensVendas 
      Height          =   405
      Left            =   12860
      TabIndex        =   6
      Top             =   5340
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Itens Venda"
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
      MICON           =   "FrmFichaFinanceira.frx":0289
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTituloJanela 
      BackColor       =   &H00000000&
      Caption         =   "Titulo Janela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15630
   End
End
Attribute VB_Name = "FrmFichaFinanceira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wCor As String
Dim GuardaCor As String
Dim wVezes As Integer
Dim DataHora1 As String
Dim rsFichaFinanc As New ADODB.Recordset
Dim rsConsultaCredito As New ADODB.Recordset
Dim rsConsultaitens As New ADODB.Recordset
Dim wNumerodeConsultas As Integer
Dim dataHora As String
Dim SQL As String
Dim WBANCO As String
Dim DATA As Date
Dim wTime As String
Dim wComparaData As String
Dim wDupliAtraso  As Integer
Dim wMaiorAtraso  As Integer
Dim MSG



Private Sub CmdDetalhe_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   'FrmConsulta.Visible = False
End If
End Sub

Private Sub CmdConsulta_Click()
wCor = &HC0FFFF
GuardaCor = &HC0FFFF


   'FrmConsulta.Visible = True
   dataHora = Format(Date, "yyyy/mm/dd") & " 00:00:00"
   SQL = "Select * from ConsultacreditoCliente " _
       & "where CCC_Codigo=" & wCodigoCliFinan & " and ccc_Data > '" & dataHora & "' and ccc_data < '" & DataHora1 & "'"
         rsConsultaCredito.CursorLocation = adUseClient
         rsConsultaCredito.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic

   If Not rsConsultaCredito.EOF Then
           'GrdConsulta.Rows = 1
           'GrdConsulta.Redraw = False
           
           wCor = &HFFC0C0
           GuardaCor = &HFFC0C0
           Do While Not rsConsultaCredito.EOF
               'GrdConsulta.AddItem rsConsultaCredito("CCC_Loja") & Chr(9) _
               & Mid(Trim(rsConsultaCredito("CCC_DAta")), 1, 10) & Chr(9) _
               & Mid(Trim(rsConsultaCredito("CCC_DAta")), 11, 9)
               If wCor = &HFFC0C0 Then
                    PintaGrid wCor
                    wCor = &HC0FFFF
                Else
                   PintaGrid wCor
                   wCor = &HFFC0C0
                End If
               rsConsultaCredito.MoveNext
           Loop
           'GrdConsulta.Redraw = True
           'GrdConsulta.SetFocus
           'GrdConsulta.Row = 1
   End If

rsConsultaCredito.Close


End Sub

Private Sub CmdConsulta_KeyUp(KeyCode As Integer, Shift As Integer)

CmdConsulta.Caption = wNumerodeConsultas
End Sub

Private Sub cmdItensVendas_Click()
CarregarGrid
End Sub

Private Sub cmdRetorna_Click()
 Unload Me
End Sub



Private Sub cmdRetorna_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   'FrmConsulta.Visible = False
End If
End Sub




Private Sub Form_Load()

'teste
lblTituloJanela.Caption = FrmFichaFinanceira.Caption

  FrmFichaFinanceira.top = 4680
  FrmFichaFinanceira.left = 90
  FrmFichaFinanceira.Width = 15180
  FrmFichaFinanceira.Height = 5790
 
  If rdoCNMatriz.State = 1 Then
    rdoCNMatriz.Close
  End If
  
  ConectaODBCMatriz
  If GLB_ConectouOK = False Then
     Exit Sub
  End If

  grdFichaFinanceira.MergeRow(0) = True
  grdFichaFinanceira.MergeRow(1) = True
  grdFichaFinanceira.MergeCol(0) = True
  grdFichaFinanceira.MergeCol(1) = True


 dataHora = Format(Date, "yyyy/mm/dd") & " 00:00:00"
 SQL = "Select count(CCC_Codigo) as NumeroConsultas from ConsultacreditoCliente " _
    & "where CCC_Codigo=" & wCodigoCliFinan & " and ccc_Data > '" & dataHora & "'"
    rsConsultaCredito.CursorLocation = adUseClient
    rsConsultaCredito.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic

     
     If rsConsultaCredito("NumeroConsultas") = 0 Then
         
         rsConsultaCredito.Close
         
        SQL = "Select max(CCC_data) as wData from ConsultacreditoCliente "
        rsConsultaCredito.CursorLocation = adUseClient
        rsConsultaCredito.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
       
        wComparaData = rsConsultaCredito("Wdata") - 4
        
    
        SQL = "delete consultacreditocliente where CCC_Data < '" _
             & Format(wComparaData, "yyyy/mm/dd") & "'"
        rdoCNMatriz.Execute (SQL)
     End If
        
        rsConsultaCredito.Close
        
     
 '-------------------------------------------------------------------------------------------------------------
 
  
 PreencheDadosCliente wCodigoCliFinan
 'AchaLojaControle
 DataHora1 = Format(Date, "mm/dd/yyyy") & " " & Format(Time, "hh:mm:ss")
 
If wConexao = "Balcao" Then
   SQL = "Insert into ConsultacreditoCliente (CCC_Data, CCC_Loja, CCC_Codigo) " _
         & "values('" & DataHora1 & "', '" & AchaLojaControle & "','" & wCodigoCliFinan & "')"
   rdoCNMatriz.Execute (SQL)
End If
 
dataHora = Format(Date, "yyyy/mm/dd") & " 00:00:00"
 
 SQL = "Select count(CCC_Codigo) as NumeroConsultas from ConsultacreditoCliente " _
    & "where CCC_Codigo=" & wCodigoCliFinan & " and ccc_Data > '" & dataHora & "' and ccc_data < '" & DataHora1 & "'"
    rsConsultaCredito.CursorLocation = adUseClient
    rsConsultaCredito.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
     
     If rsConsultaCredito.EOF = False Then
        wNumerodeConsultas = rsConsultaCredito("NumeroConsultas")
        If wNumerodeConsultas = -1 Then
           wNumerodeConsultas = 0
           rsConsultaCredito.Close
        End If
        CmdConsulta.Caption = wNumerodeConsultas

        rsConsultaCredito.Close
     End If
 

 End Sub
  Function PesquisaCliente(ByVal Cliente As String)
  
   SQL = ""
        SQL = "Select * from HistoricoClientes " _
            & "where HIC_codigoCliente  = '" & Cliente & "' "
        rsFichaFinanc.CursorLocation = adUseClient
        rsFichaFinanc.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
  
       
  End Function
  
  Function PreencheDadosCliente(ByVal Cliente As Double)

  

   '" & wCodigoCliFinan & "' "
  
       PesquisaCliente wCodigoCliFinan
       

       
    If Not rsFichaFinanc.EOF Then
        grdFichaFinanceira.TextMatrix(1, 1) = Format(rsFichaFinanc("HIC_limiteCredito"), "###,###,###,##0.00")
        grdFichaFinanceira.TextMatrix(2, 1) = Format(rsFichaFinanc("HIC_dataLimiteCredito"), "dd/mm/yyyy")
        grdFichaFinanceira.TextMatrix(3, 1) = Format(rsFichaFinanc("HIC_duplAbertas"), "###,###,###,##0.00") & " - " & rsFichaFinanc("HIC_qtdeDuplAbertas")
        grdFichaFinanceira.TextMatrix(4, 1) = rsFichaFinanc("HIC_duplAtrasado")
        grdFichaFinanceira.TextMatrix(5, 1) = Format(rsFichaFinanc("HIC_saldoCompras"), "###,###,###,##0.00")
        grdFichaFinanceira.TextMatrix(6, 1) = Format(rsFichaFinanc("HIC_ultimaCompra"), "###,###,###,##0.00")
        grdFichaFinanceira.TextMatrix(7, 1) = Format(rsFichaFinanc("HIC_dataUltimaCompra"), "dd/mm/yyyy")
        grdFichaFinanceira.TextMatrix(8, 1) = Format(rsFichaFinanc("HIC_maiorCompra"), "###,###,###,##0.00")
        grdFichaFinanceira.TextMatrix(9, 1) = Format(rsFichaFinanc("HIC_dataMaiorCompra"), "dd/mm/yyyy")
        grdFichaFinanceira.TextMatrix(10, 1) = rsFichaFinanc("HIC_quantidadeCompras")
        grdFichaFinanceira.TextMatrix(11, 1) = Format(rsFichaFinanc("HIC_ultimoPagamento"), "###,###,###,##0.00")
        grdFichaFinanceira.TextMatrix(12, 1) = Format(rsFichaFinanc("HIC_dataUltimoPagamento"), "dd/mm/yyyy")
        grdFichaFinanceira.TextMatrix(13, 1) = rsFichaFinanc("HIC_maiorAtraso") & " Dia(s)"
        grdFichaFinanceira.TextMatrix(14, 1) = Format(rsFichaFinanc("HIC_TotalCompras"), "###,###,###,##0.00")

       
        CamposGrid
        grdFichaFinanceira.Row = 1
    End If

 
       rsFichaFinanc.Close
       
End Function
 
Sub consultacredito()


SQL = "Select count(CCC_Codigo) as NumeroConsultas from ConsultacreditoCliente " _
    & "where CCC_Codigo=" & wCodigoCliFinan & "ccc_Data = " & Format(Date, "mm/dd/yyyy")
    Set rsConsultaCredito = rdoCNMatriz.OpenResultset(SQL)

     
     If rsConsultaCredito.EOF = False Then
        CmdConsulta.Caption = rsConsultaCredito("NumeroConsultas")
        wNumerodeConsultas = rsConsultaCredito("NumeroConsultas")
        grdFichaFinanceira.TextMatrix(18, 1) = wNumerodeConsultas
     End If
End Sub

Private Sub GrdConsulta_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   'FrmConsulta.Visible = False
End If
End Sub
Function PintaGrid(ByVal Cor As String)
  
  'GrdConsulta.Row = GrdConsulta.Rows - 1
  'GrdConsulta.Col = 0
  'GrdConsulta.ColSel = 2
  'GrdConsulta.FillStyle = flexFillRepeat
  'GrdConsulta.CellBackColor = Cor
  'GrdConsulta.FillStyle = flexFillSingle

End Function

Sub CamposGrid()
         If (rsFichaFinanc("HIC_JurosCartorio") <> "0" Or rsFichaFinanc("HIC_duplAtrasado") <> "0") Or rsFichaFinanc("HIC_SaldoCompras") < 0 Or rsFichaFinanc("HIC_SaldoCompras") = 0 Then

      grdFichaFinanceira.Row = 16
      grdFichaFinanceira.Col = 0
      grdFichaFinanceira.CellBackColor = &HFF&
      grdFichaFinanceira.Col = 1
      grdFichaFinanceira.CellBackColor = &HFF&

End If

grdFichaFinanceira.Row = 5
grdFichaFinanceira.Col = 1
grdFichaFinanceira.CellForeColor = &HC00000

   

             
DATA = Format(rsFichaFinanc("HIC_dataUltimaCompra"), "dd/mm/yyyy")
If (DateDiff("m", DATA, Now) > 3) Then

      grdFichaFinanceira.Row = 7
      grdFichaFinanceira.Col = 1
      grdFichaFinanceira.CellForeColor = &HFF&
End If

If rsFichaFinanc("HIC_SaldoCompras") < 0 Or rsFichaFinanc("HIC_SaldoCompras") = 0 Then

      grdFichaFinanceira.Row = 5
      grdFichaFinanceira.Col = 1
      grdFichaFinanceira.CellForeColor = &HFF&
End If
    
End Sub
Sub CarregarGrid()
grdItensProduto.Rows = 1
SQL = ""

If wCodigoCliFinan > 90000 Then
    SQL = "select PR_Referencia, PR_Descricao, VI_Quantidade, VI_PrecoUnitario, VI_ValorMercadoria, VI_NotaFiscal, VI_Serie from ItemNFVenda, CapaNFVenda, Produto " _
        & "Where VI_NotaFiscal = VC_NotaFiscal and VI_Serie = VC_Serie and VI_LojaOrigem = VC_LojaOrigem and VI_Referencia = PR_Referencia and VC_Cliente =" & wCodigoCliFinan & " and vc_lojaVenda = '" & Trim(GLB_Loja) & "'"
Else
    SQL = "select PR_Referencia, PR_Descricao, VI_Quantidade, VI_PrecoUnitario, VI_ValorMercadoria, VI_NotaFiscal, VI_Serie from ItemNFVenda, CapaNFVenda, Produto " _
        & "Where VI_NotaFiscal = VC_NotaFiscal and VI_Serie = VC_Serie and VI_LojaOrigem = VC_LojaOrigem and VI_Referencia = PR_Referencia and VC_Cliente =" & wCodigoCliFinan

End If
rsConsultaitens.CursorLocation = adUseClient
rsConsultaitens.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic

   If Not rsConsultaitens.EOF Then
   Do While Not rsConsultaitens.EOF
                grdItensProduto.AddItem rsConsultaitens("PR_Referencia") & Chr(9) _
               & rsConsultaitens("PR_Descricao") & Chr(9) _
               & rsConsultaitens("VI_Quantidade") & Chr(9) _
               & Format(rsConsultaitens("VI_PrecoUnitario"), "0.00") & Chr(9) _
               & Format(rsConsultaitens("VI_ValorMercadoria"), "0.00") & Chr(9) _
               & rsConsultaitens("VI_NotaFiscal") & Chr(9) _
               & rsConsultaitens("VI_Serie")
       
   rsConsultaitens.MoveNext
   Loop
End If
rsConsultaitens.Close
End Sub

Private Sub grdItensProduto_DblClick()
'ricardo 19/10/2016
Dim variavelChaveNfe As String
Dim nf As String
Dim variavelEmail As String
Dim rsPesquisa As New ADODB.Recordset
Dim rsPesquisaEmail As New ADODB.Recordset


   '------------------------- SELECIONA A NOTA CORRESPONDENTE AO CLICK NO GRID -----------------------------------------------------------
    SQL = ""
'    SQL = "select vc_chavenfe,vc_notafiscal,vc_serie from Capanfvenda where vc_notafiscal = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 5) & "' " _
'         & "and VC_LojaOrigem = '" & AchaLojaControle & "' and vc_serie = 'NE'"
         
         SQL = "select vc_chavenfe,vc_notafiscal,vc_serie from Capanfvenda where vc_notafiscal = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 5) & "' " _
         & "and VC_LojaOrigem = '" & AchaLojaControle & "' and vc_serie = 'NE'"
       
    rsPesquisa.CursorLocation = adUseClient
    rsPesquisa.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    If rsPesquisa.EOF Then
        MsgBox "Essa nota não é do Tipo Eletronica", vbInformation
        Exit Sub
    End If
    
    
    If Trim(rsPesquisa("vc_chavenfe") = "" Or rsPesquisa("vc_notafiscal") = "" Or rsPesquisa("vc_serie") = "") Then
        MsgBox "Não temos informação sobre essa nota favor contacta o CPD"
    Exit Sub
    End If
    
    
    If Trim(rsPesquisa("vc_serie")) <> "NE" Or IsNull(Trim(rsPesquisa("vc_serie"))) Then
        MsgBox "Nota selecionada não é Eletronica", vbInformation
        Exit Sub
    End If
    
    
    If IsNull(rsPesquisa("vc_chavenfe")) Then
        MsgBox "Não é possivel enviar a chave de acesso", vbInformation
        Exit Sub
    End If
    
    If rsPesquisa("vc_notafiscal") = "" Then
          MsgBox "Não temos informação sobre essa nota favor contacta o CPD"
          Exit Sub
    End If
    
    
     variavelChaveNfe = rsPesquisa("vc_chavenfe")
     nf = rsPesquisa("vc_notafiscal")
     

    
    '----------------------- SELECIONA CLIENTE PARA VERIFICAÇÃO SE EXISTE EMAIL CADASTRADO -----------------------------------------------
    
    SQL = " Select CE_EMail from FIN_Cliente where CE_CodigoCliente = '" & frmConsCliente.txtPesquisaCliente & "'"
    
    
      rsPesquisaEmail.CursorLocation = adUseClient
      rsPesquisaEmail.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
      
     If IsNull(rsPesquisaEmail("CE_Email")) Then
        MsgBox "Email não valido ou não cadastrado favor verificar", vbDefaultButton2, "Email"
        Exit Sub
     Else
        variavelEmail = rsPesquisaEmail("CE_EMail")
     End If
     
      
    
    
    '--------------------- INSERE ChaveNfe na Tabela NFe_ide --------------------------------------------------------------------------------
    SQL = "INSERT INTO NFe_ide(eLoja,eNF,eSerie,Situacao,cUF,cNF,natOp,indPag,mod,serie,nNF,dEmi,dSaiEnt,hSaiEnt,tpNF,cMunFG,tpImp,tpEmis, " _
         & "cDV,tpAmb,finNFe,procEmi,verProc,dhCont,xJust,ChaveAcesso,refNFe,IDDEST,INDFINAL,INDPRES) " _
         & " VALUES ('" & AchaLojaControle & "','" & nf & "','NE','D','35','" & nf & "','DANFE E XML','0','55','1','" & nf & "','" & Format(Date, "yyyy/mm/dd") & "','" & Format(Date, "yyyy/mm/dd") & "','" & Format(Date, "yyyy/mm/dd") & "','1','3550308','1','1', " _
         & " '','2','1','3','2.0.0','" & Format(Date, "yyyy/mm/dd") & "','Erro no envio da Nota Fiscal Eletronica devido a problemas com Sefaz','" & variavelChaveNfe & "','','1','1','1')"
    
           
         adoCNLoja.Execute (SQL)
    
    

rsPesquisa.Close
rsPesquisaEmail.Close

End Sub


