VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmPesquisaEstoqueCentral 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Pesquisa Estoque Central"
   ClientHeight    =   5655
   ClientLeft      =   6825
   ClientTop       =   4875
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6825
      OleObjectBlob   =   "PesquisaestoqueCentral.frx":0000
      Top             =   5295
   End
   Begin VSFlex7UCtl.VSFlexGrid grdTendencia 
      Height          =   2040
      Left            =   4125
      TabIndex        =   0
      Top             =   135
      Width           =   2205
      _cx             =   3889
      _cy             =   3598
      _ConvInfo       =   1
      Appearance      =   0
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
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"PesquisaestoqueCentral.frx":0234
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
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   0
   End
   Begin VSFlex7UCtl.VSFlexGrid grdEstLoja 
      Height          =   4530
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   3870
      _cx             =   6826
      _cy             =   7990
      _ConvInfo       =   1
      Appearance      =   0
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"PesquisaestoqueCentral.frx":02C5
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
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   0
   End
End
Attribute VB_Name = "frmPesquisaEstoqueCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wReferencia As String
Dim SQL As String

Private Sub cmdGrava_Click()
    
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Call AjustaTela(frmPesquisaEstoqueCentral)
  
  'Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
  'Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
  'Skin1.ApplySkin Me.hwnd
  grdEstLoja.Editable = flexEDNone
  Screen.MousePointer = vbHourglass

  
  If rdoCNMatriz.State = 1 Then
    rdoCNMatriz.Close
  End If
  
  ConectaODBCMatriz
  If GLB_ConectouOK = False Then
     MsgBox "Erro ao conectar-se ao Banco de Dados da Matriz", vbCritical, "Atenção"
     Exit Sub
  End If
  
  wReferencia = frmPedido.grdItensProduto.TextMatrix(frmPedido.grdItensProduto.Row, 0)
  CarregaEstoqueMatriz wReferencia
  MontaTendencia wReferencia
  
  grdEstLoja.Row = 1
  Screen.MousePointer = vbNormal

End Sub

Function CarregaEstoqueMatriz(ByVal Referencia As String)
   
'================= PESQUISA ESTOQUE LOJAS =================
   SQL = ""
   SQL = "Select ES_Loja, ES_Estoque, ES_Transito, ES_MaximoInformado, ES_MinimoInformado, ES_Romaneio From Estoque, Loja " _
         & "Where ES_Loja = LO_Loja and LO_MostraEstoque = 'S' and ES_Referencia = '" & Referencia & "' " _
         & "and ES_Loja not in ('CONSO') and LO_Situacao = 'A' order by LO_Regiao"
        
    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    If Not RsDados.EOF Then
        Do While Not RsDados.EOF

           grdEstLoja.AddItem RsDados("ES_Loja") & Chr(9) _
                & Val(RsDados("ES_Estoque")) & Chr(9) _
                & Val(RsDados("ES_Transito")) & Chr(9) _
                & Val(RsDados("ES_Romaneio")) & Chr(9) _
                & Val(RsDados("ES_MinimoInformado")) & Chr(9) _
                & Val(RsDados("ES_MaximoInformado"))
           RsDados.MoveNext
        Loop
    End If
    RsDados.Close
    
'================= PESQUISA ESTOQUE CMC =================

    SQL = ""
    SQL = "Select count(*) as ContadorCMC from loja where lo_situacao = 'A' and lo_regiao = 990"
    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
       
    If RsDados("ContadorCMC") > 0 Then
        RsDados.Close
            
        SQL = ""
        SQL = "Select Sum(ES_Estoque) as Estoque, Sum(ES_Transito) as Transito, Sum(ES_Romaneio) as Romaneio, " & _
              "Sum(ES_MaximoInformado) as MaximoInformado, Sum(ES_MinimoInformado) as MinimoInformado From Estoque, Loja " & _
              "Where ES_Loja = LO_Loja and LO_Regiao = 990 and ES_Referencia = '" & Referencia & "' And " & _
              "ES_Loja not in ('CONSO','CD') and LO_Situacao = 'A'"
    
        RsDados.CursorLocation = adUseClient
        RsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
        If Not RsDados.EOF Then

            grdEstLoja.AddItem "CMC" & Chr(9) _
                & Val(RsDados("Estoque")) & Chr(9) _
                & Val(RsDados("Transito")) & Chr(9) _
                & Val(RsDados("Romaneio")) & Chr(9) _
                & Val(RsDados("MinimoInformado")) & Chr(9) _
                & Val(RsDados("MaximoInformado"))

            grdEstLoja.Row = grdEstLoja.Rows - 1
            grdEstLoja.Col = 0
            grdEstLoja.ColSel = 5
            grdEstLoja.FillStyle = flexFillRepeat
            grdEstLoja.CellBackColor = &HF7F8D3
            grdEstLoja.FillStyle = flexFillSingle
        End If
   End If
   RsDados.Close
'================= PESQUISA ESTOQUE CD =================
    SQL = ""
    SQL = "Select (Case When ES_Loja = 'CONSO' Then 'Total' Else ES_Loja End) as ES_Loja, ES_Estoque, ES_Transito, ES_MaximoInformado, ES_MinimoInformado, ES_Romaneio From Estoque, Loja " _
        & "Where ES_Loja = LO_Loja and LO_MostraEstoque = 'S' and ES_Referencia = '" & Referencia & "' and ES_Loja in('CONSO', 'CD') " _
        & "Order By ES_Loja Desc"
        
    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    If Not RsDados.EOF Then
            grdEstLoja.AddItem RsDados("ES_Loja") & Chr(9) _
                             & Val(RsDados("ES_Estoque")) & Chr(9) _
                             & Val(RsDados("ES_Transito")) & Chr(9) _
                             & Val(RsDados("ES_Romaneio")) & Chr(9) _
                             & Val(RsDados("ES_MinimoInformado")) & Chr(9) _
                             & Val(RsDados("ES_MaximoInformado"))
        
      grdEstLoja.Row = grdEstLoja.Rows - 1
      grdEstLoja.Col = 0
      grdEstLoja.ColSel = 5
      grdEstLoja.FillStyle = flexFillRepeat
      grdEstLoja.FillStyle = flexFillSingle
   
   End If
   RsDados.Close




'   If grdEstLoja.TextMatrix(grdEstLoja.Rows - 1, 0) = "CMC" Then
'      grdEstLoja.Row = grdEstLoja.Rows - 1
'      grdEstLoja.Col = 0
'      grdEstLoja.ColSel = 5
'      grdEstLoja.FillStyle = flexFillRepeat
'      grdEstLoja.CellBackColor = &HF7F8D3
'      grdEstLoja.FillStyle = flexFillSingle
        
'   End If
'   grdEstLoja.Row = 1
'   grdEstLoja.Col = 0
'   grdEstLoja.ColSel = 5
'   grdEstLoja.FillStyle = flexFillRepeat
'   grdEstLoja.CellBackColor = &HF7F8D3     '&HFCD88B
'   grdEstLoja.FillStyle = flexFillSingle

End Function

Private Sub grdEstLoja_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
    frmPedido.grdItensProduto.SetFocus
'    frmPedido.picQuadroGeral.Width = 9975

 End If
End Sub

Private Sub MontaTendencia(Referencia As String)
Dim wLoja As String

  SQL = "Select * from ControleSistema"
  rdoControle.CursorLocation = adUseClient
  rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
  If rdoControle.EOF Then
     MsgBox "Problemas com o Sistema (Controle Sistema) entrar em contato com o CPD", vbCritical, "Atenção"
     rdoControle.Close
  Else
     wLoja = Trim(rdoControle("CTS_Loja"))
     rdoControle.Close
  End If
    
    grdTendencia.TextMatrix(1, 0) = TraduzMes(Month(Date))
    grdTendencia.TextMatrix(2, 0) = TraduzMes(Month(DateAdd("m", -1, Date)))
    grdTendencia.TextMatrix(3, 0) = TraduzMes(Month(DateAdd("m", -2, Date)))
    grdTendencia.TextMatrix(4, 0) = TraduzMes(Month(DateAdd("m", -3, Date)))
    grdTendencia.TextMatrix(5, 0) = TraduzMes(Month(DateAdd("m", -4, Date)))
    grdTendencia.TextMatrix(6, 0) = TraduzMes(Month(DateAdd("m", -5, Date)))
    
    
    
    SQL = ""
    SQL = "SELECT  ES_Estoque, ES_Venda, ES_Estoque1, ES_Venda1, " & _
          "ES_Estoque2, ES_Venda2, ES_Estoque3, ES_Venda3, " & _
          "ES_Estoque4, ES_Venda4, ES_Estoque5, ES_Venda5, " & _
          "ES_Estoque6 , ES_Venda6 From Estoque " & _
          "Where ES_Loja = '" & wLoja & "' and ES_Referencia = '" & Referencia & "'"
          
    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
          
    If Not RsDados.EOF Then
      grdTendencia.TextMatrix(1, 1) = RsDados("ES_Estoque")
      grdTendencia.TextMatrix(1, 2) = RsDados("ES_Venda")
      grdTendencia.TextMatrix(2, 1) = RsDados("ES_Estoque1")
      grdTendencia.TextMatrix(2, 2) = RsDados("ES_Venda1")
      grdTendencia.TextMatrix(3, 1) = RsDados("ES_Estoque2")
      grdTendencia.TextMatrix(3, 2) = RsDados("ES_Venda2")
      grdTendencia.TextMatrix(4, 1) = RsDados("ES_Estoque3")
      grdTendencia.TextMatrix(4, 2) = RsDados("ES_Venda3")
      grdTendencia.TextMatrix(5, 1) = RsDados("ES_Estoque4")
      grdTendencia.TextMatrix(5, 2) = RsDados("ES_Venda4")
      grdTendencia.TextMatrix(6, 1) = RsDados("ES_Estoque5")
      grdTendencia.TextMatrix(6, 2) = RsDados("ES_Venda5")
    

    End If
    RsDados.Close
          
End Sub

Function TraduzMes(ByVal Mes As Integer) As String

    Select Case Mes
        Case 1, -11: TraduzMes = "Janeiro"
        Case 2, -10: TraduzMes = "Fevereiro"
        Case 3, -9: TraduzMes = "Março"
        Case 4, -8: TraduzMes = "Abril"
        Case 5, -7: TraduzMes = "Maio"
        Case 6, -6: TraduzMes = "Junho"
        Case 7, -5: TraduzMes = "Julho"
        Case 8, -4: TraduzMes = "Agosto"
        Case 9, -3: TraduzMes = "Setembro"
        Case 10, -2: TraduzMes = "Outubro"
        Case 11, -1: TraduzMes = "Novembro"
        Case 12, 0: TraduzMes = "Dezembro"
    End Select
    
End Function

Private Sub grdTendencia_GotFocus()
    grdEstLoja.SetFocus
End Sub

Private Sub grdTendencia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
    frmPedido.grdItensProduto.SetFocus
'    frmPedido.picQuadroGeral.Width = 9975

 End If
End Sub
