VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLembrete 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Lembrete"
   ClientHeight    =   5565
   ClientLeft      =   2640
   ClientTop       =   3045
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin Project1.chameleonButton cmdRetorna 
      Height          =   0
      Left            =   5460
      TabIndex        =   5
      Top             =   4815
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "Retorna"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12582912
      MPTR            =   1
      MICON           =   "frmLembrete.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtObservacao 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3810
      Width           =   6300
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   30
      Left            =   3525
      TabIndex        =   0
      Top             =   1245
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   327682
      Appearance      =   1
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdLembrete 
      Height          =   2880
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Se desejar excluir, clique duas vezes sobre a referência a ser excluido."
      Top             =   555
      Width           =   6330
      _cx             =   11165
      _cy             =   5080
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
      BackColor       =   12648447
      ForeColor       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   591035
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   12648447
      BackColorAlternate=   12648447
      GridColor       =   8454143
      GridColorFixed  =   8454143
      TreeColor       =   8454143
      FloodColor      =   12648447
      SheetBorder     =   8454143
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
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
      FormatString    =   $"frmLembrete.frx":001C
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   -1  'True
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
      BackColorFrozen =   12648447
      ForeColorFrozen =   0
      WallPaperAlignment=   0
   End
   Begin Project1.chameleonButton cmdRetornar 
      Height          =   405
      Left            =   5235
      TabIndex        =   6
      Top             =   5055
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmLembrete.frx":00C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblObservacao 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Observação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000904BB&
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   3540
      Width           =   1125
   End
   Begin VB.Label lblLembreMe 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Lembre-me quando chegar"
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
      Height          =   330
      Left            =   -30
      TabIndex        =   2
      Top             =   135
      Width           =   6705
   End
End
Attribute VB_Name = "frmLembrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String

Private Sub chameleonButton1_Click()
    Unload Me
End Sub

Private Sub cmdRetorna_Click()

Unload Me
frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub cmdRetornar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
  Call AjustaTela(frmLembrete)
'  Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
 
'   cmdRetorna.Height = 570
'   cmdRetorna.Left = 5460
'   cmdRetorna.Top = 4815
'   cmdRetorna.Width = 1020
  
  Call CarregaLembrete
  cmdRetorna.SetFocus
  
  lblObservacao.Visible = False
  txtObservacao.Visible = False

End Sub

Private Sub grdLembrete_Click()
    txtObservacao.Visible = True
    lblObservacao.Visible = True
    txtObservacao.Text = UCase(grdLembrete.TextMatrix(grdLembrete.Row, 4))
End Sub

Private Sub grdLembrete_DblClick()

    SQL = "Update LembreMe set LEM_situacao = 'L' " & _
          "where LEM_Sequencia = " & grdLembrete.TextMatrix(grdLembrete.Row, 3)
         
         adoCNLoja.Execute (SQL)
    
    Call CarregaLembrete
      
End Sub

Private Sub CarregaLembrete()
 grdLembrete.Rows = 1

  SQL = ""
  SQL = "Select LEM_Referencia, LEM_Observacao, LEM_Data, LEM_Sequencia,PR_Descricao from LembreMe,ProdutoLoja " & _
        "where lem_situacao = 'O' and lem_vendedor = " & Mid(frmPedido.txtVendedor.Text, 1, 3) & _
        " and pr_referencia = lem_referencia Order by LEM_Data,LEM_Referencia"
        rsLembrete.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

If Not rsLembrete.EOF Then
    Do While Not rsLembrete.EOF
        frmLembrete.grdLembrete.AddItem Mid(rsLembrete("LEM_Data"), 1, 10) & Chr(9) & rsLembrete("LEM_Referencia") & _
                    Chr(9) & rsLembrete("PR_descricao") & Chr(9) & rsLembrete("LEM_Sequencia") & _
                    Chr(9) & rsLembrete("LEM_Observacao")
        rsLembrete.MoveNext
    Loop
Else
    rsLembrete.Close
    Unload Me
    frmPedido.txtPesquisar.SetFocus
    Exit Sub
End If
rsLembrete.Close


End Sub

