VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmReativacaoCliente 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Reativando Cliente"
   ClientHeight    =   5970
   ClientLeft      =   2745
   ClientTop       =   165
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   405
      Left            =   8280
      TabIndex        =   70
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Vendas"
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmDescricao 
      BackColor       =   &H00505050&
      Height          =   1815
      Left            =   3120
      TabIndex        =   65
      Top             =   1800
      Width           =   5415
      Begin VB.TextBox txtProduto 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   480
         Width           =   5025
      End
      Begin VB.TextBox txtProduto 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
         Width           =   5025
      End
      Begin VB.TextBox txtProduto 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5025
      End
      Begin VB.Label Label18 
         BackColor       =   &H00505050&
         Caption         =   "Referencia/Descrição"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   6000
      Left            =   9555
      ScaleHeight     =   6000
      ScaleWidth      =   45
      TabIndex        =   63
      Top             =   210
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14820
      TabIndex        =   62
      Top             =   5085
      Width           =   14820
   End
   Begin Project1.chameleonButton cmdAvancar 
      Height          =   405
      Left            =   10545
      TabIndex        =   59
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   ">>"
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdRetornar 
      Height          =   405
      Left            =   9420
      TabIndex        =   58
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "<<"
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdImprimir 
      Height          =   405
      Left            =   12795
      TabIndex        =   57
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Imprimir"
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdRetorna 
      Height          =   405
      Left            =   13920
      TabIndex        =   56
      Top             =   5280
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   11670
      TabIndex        =   55
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Iniciar"
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   615
      Left            =   13800
      TabIndex        =   54
      Top             =   11400
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ComboBox cmbLojas 
      BackColor       =   &H8000000A&
      Height          =   315
      Index           =   0
      Left            =   15735
      TabIndex        =   53
      Text            =   "Combo1"
      Top             =   6525
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00505050&
      Caption         =   "No Período"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4335
      TabIndex        =   47
      Top             =   4140
      Visible         =   0   'False
      Width           =   5235
      Begin MSMask.MaskEdBox txtDataFimCrm 
         Height          =   315
         Left            =   2160
         TabIndex        =   50
         Top             =   240
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483638
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDataIniCrm 
         Height          =   315
         Left            =   390
         TabIndex        =   48
         Top             =   240
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483638
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Project1.chameleonButton cmdOK 
         Height          =   375
         Left            =   3600
         TabIndex        =   60
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "OK"
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
         BCOLO           =   5263440
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReativacaoCliente.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label24 
         BackColor       =   &H00505050&
         Caption         =   "até"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1800
         TabIndex        =   49
         Top             =   285
         Width           =   270
      End
   End
   Begin VB.TextBox txtMascara 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   150
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Text1"
      ToolTipText     =   "Clique em Iniciar para ver o Telefone"
      Top             =   1800
      Width           =   2325
   End
   Begin VB.TextBox txtTelefone 
      Height          =   315
      Left            =   360
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1800
      Width           =   1245
   End
   Begin VB.CheckBox chkMarketing 
      BackColor       =   &H00505050&
      Caption         =   "E-Mail Marketing"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8085
      TabIndex        =   19
      Top             =   3450
      Width           =   1500
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdDados 
      Height          =   765
      Left            =   4320
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   5235
      _cx             =   9234
      _cy             =   1349
      _ConvInfo       =   1
      Appearance      =   0
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   16777215
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
      Cols            =   32
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReativacaoCliente.frx":00C4
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
   Begin VB.ComboBox cmbUF 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   8745
      TabIndex        =   7
      Top             =   1170
      Width           =   810
   End
   Begin VB.TextBox txtCidade 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1170
      Width           =   3375
   End
   Begin VB.TextBox txtEndereco 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   150
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1170
      Width           =   5055
   End
   Begin VB.TextBox txtGerente 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   5580
      MaxLength       =   30
      TabIndex        =   21
      Top             =   4365
      Width           =   3975
   End
   Begin MSMask.MaskEdBox txtData 
      Height          =   315
      Left            =   4350
      TabIndex        =   20
      Top             =   4365
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483638
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtNFicha 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   8235
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   4
      Top             =   540
      Width           =   1320
   End
   Begin VB.TextBox txtEMail 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   4350
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3675
      Width           =   5205
   End
   Begin VB.TextBox txtComentario 
      BackColor       =   &H8000000A&
      Height          =   1215
      Left            =   150
      MaxLength       =   254
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3675
      Width           =   4035
   End
   Begin VB.ComboBox cmbOcorrencia 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   3045
      Width           =   9405
   End
   Begin MSMask.MaskEdBox txtDataFim 
      Height          =   315
      Left            =   8355
      TabIndex        =   15
      Top             =   2430
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483638
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDataIni 
      Height          =   315
      Left            =   6630
      TabIndex        =   14
      Top             =   2430
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483638
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCompras 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5340
      MaxLength       =   6
      TabIndex        =   13
      Top             =   2430
      Width           =   1200
   End
   Begin VB.TextBox txtValorMaiorCompra 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   12
      Text            =   "999.999.999,99"
      Top             =   2430
      Width           =   1290
   End
   Begin VB.TextBox txtValorUltimaCompra 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1365
      MaxLength       =   15
      TabIndex        =   10
      Text            =   "999.999.999,99"
      Top             =   2430
      Width           =   1290
   End
   Begin MSMask.MaskEdBox txtDataUltimaCompra 
      Height          =   315
      Left            =   150
      TabIndex        =   9
      Top             =   2430
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483638
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtContato 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   2565
      MaxLength       =   30
      TabIndex        =   8
      Top             =   1800
      Width           =   6990
   End
   Begin VB.TextBox txtVendedor 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   840
      MaxLength       =   3
      TabIndex        =   1
      Top             =   540
      Width           =   800
   End
   Begin VB.TextBox txtLoja 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   150
      MaxLength       =   3
      TabIndex        =   0
      Top             =   540
      Width           =   600
   End
   Begin VB.TextBox txtCliente 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   3060
      MaxLength       =   30
      TabIndex        =   3
      Top             =   540
      Width           =   5085
   End
   Begin VB.TextBox txtCodCliente 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1725
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   2
      Top             =   540
      Width           =   1245
   End
   Begin MSMask.MaskEdBox txtDataMaiorCompra 
      Height          =   315
      Left            =   2745
      TabIndex        =   11
      Top             =   2430
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483638
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdDados1 
      Height          =   4515
      Left            =   9780
      TabIndex        =   51
      Top             =   405
      Width           =   5190
      _cx             =   9155
      _cy             =   7964
      _ConvInfo       =   1
      Appearance      =   1
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   16777215
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"frmReativacaoCliente.frx":053D
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
   Begin Project1.chameleonButton cmdGravarProduto 
      Height          =   615
      Left            =   8910
      TabIndex        =   64
      Top             =   7230
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "Carrega CRM"
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
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReativacaoCliente.frx":05EA
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
      Caption         =   "Reativação Cliente"
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
      TabIndex        =   61
      Top             =   0
      Width           =   15630
   End
   Begin VB.Label lblLoja 
      AutoSize        =   -1  'True
      BackColor       =   &H00505050&
      Caption         =   "Loja"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   17385
      TabIndex        =   52
      Top             =   5745
      Width           =   300
   End
   Begin VB.Label Label23 
      BackColor       =   &H00505050&
      Caption         =   "até"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7965
      TabIndex        =   46
      Top             =   2490
      Width           =   270
   End
   Begin VB.Label Label22 
      BackColor       =   &H00505050&
      Caption         =   "Gerente"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   5565
      TabIndex        =   45
      Top             =   4125
      Width           =   585
   End
   Begin VB.Label Label21 
      BackColor       =   &H00505050&
      Caption         =   "Data"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4350
      TabIndex        =   44
      Top             =   4125
      Width           =   360
   End
   Begin VB.Label Label20 
      BackColor       =   &H00505050&
      Caption         =   "E-Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   4350
      TabIndex        =   43
      Top             =   3435
      Width           =   450
   End
   Begin VB.Label Label19 
      BackColor       =   &H00505050&
      Caption         =   "Comentário"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   42
      Top             =   3435
      Width           =   810
   End
   Begin VB.Label Label17 
      BackColor       =   &H00505050&
      Caption         =   "Ocorrência"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   150
      TabIndex        =   41
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label Label16 
      BackColor       =   &H00505050&
      Caption         =   "No Período"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   6660
      TabIndex        =   40
      Top             =   2190
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00505050&
      Caption         =   "Nº Compras"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5415
      TabIndex        =   39
      Top             =   2190
      Width           =   840
   End
   Begin VB.Label Label14 
      BackColor       =   &H00505050&
      Caption         =   "Valor"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   4080
      TabIndex        =   38
      Top             =   2190
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00505050&
      Caption         =   "Maior Compra"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2895
      TabIndex        =   37
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00505050&
      Caption         =   "Valor"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   1365
      TabIndex        =   36
      Top             =   2190
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H00505050&
      Caption         =   "Última Compra"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   150
      TabIndex        =   35
      Top             =   2190
      Width           =   1020
   End
   Begin VB.Label Label10 
      BackColor       =   &H00505050&
      Caption         =   "Contato Sr(a)."
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2850
      TabIndex        =   34
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00505050&
      Caption         =   "Telefone"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   150
      TabIndex        =   33
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label8 
      BackColor       =   &H00505050&
      Caption         =   "UF"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   8760
      TabIndex        =   32
      Top             =   930
      Width           =   225
   End
   Begin VB.Label Label7 
      BackColor       =   &H00505050&
      Caption         =   "Cidade"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5430
      TabIndex        =   31
      Top             =   930
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00505050&
      Caption         =   "Endereço"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   30
      Top             =   930
      Width           =   690
   End
   Begin VB.Label Label5 
      BackColor       =   &H00505050&
      Caption         =   "Nº Ficha"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8250
      TabIndex        =   29
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00505050&
      Caption         =   "Nome Cliente"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3105
      TabIndex        =   28
      Top             =   300
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H00505050&
      Caption         =   "Codigo"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1830
      TabIndex        =   27
      Top             =   300
      Width           =   510
   End
   Begin VB.Label Label2 
      BackColor       =   &H00505050&
      Caption         =   "Vendedor"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   945
      TabIndex        =   26
      Top             =   300
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00505050&
      Caption         =   "Loja"
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   150
      TabIndex        =   25
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frmReativacaoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim RdoRsDados As rdoResultset
Dim RdoRsDados As New ADODB.Recordset

'Dim RdoRsCRM As rdoResultset
Dim RdoRsCRM As New ADODB.Recordset
Dim wWhere As String
Dim I As Integer
Dim wRetorna As Integer
Dim wMarketing As String
Dim wOcorrencia As String
Dim wQtdRegistros As Integer
Dim wLoja As String
Dim wSalvaTimeOut As Integer

'ricardo

Dim Linha As Integer
'Dim RdoLerCliente As rdoResultset
Dim RdoLerCliente As New ADODB.Recordset
'Dim RdoCarregaProduto As rdoResultset
Dim RdoCarregaProduto As New ADODB.Recordset
'Dim RdoRsVendedor As rdoResultset
Dim RdoRsVendedor As New ADODB.Recordset
Dim Idxp   As Integer
Dim wCliente As Double
Dim wReferencia As String
Dim wVendedor As Integer
Dim wTotalFichas As Integer
Dim wTotalFichasAbertas As Integer
Dim wTotalFichasFechadas As Integer
Dim wTotalFichasRetorno As Integer
Dim NomeUsuario As String

Private Sub chameleonButton1_Click()
    frmDescricao.Visible = True
End Sub

Private Sub chkMarketing_Click()
    chkMarketing.Refresh
End Sub

Public Sub ProximoCliente()
''
If grdDados.Rows = 0 Then
   Exit Sub
Else
    'cmdGrava.Caption = "Iniciar"
    
    'If I <> grdDados.Rows - 1 Then
       grdDados.Row = grdDados.Row + 1
       I = grdDados.Row
       
          With grdDados
              txtNFicha.Text = Format(.TextMatrix(I, 0), "000000")
              txtLoja.Text = .TextMatrix(I, 1)
              txtVendedor.Text = .TextMatrix(I, 2)
              txtCodCliente.Text = .TextMatrix(I, 4)
              txtCliente.Text = Trim(.TextMatrix(I, 5))
              txtTelefone.Text = Trim(.TextMatrix(I, 6))
              txtEndereco.Text = Trim(.TextMatrix(I, 7))
              txtCidade.Text = Trim(.TextMatrix(I, 8))
              cmbUF.Text = Trim(.TextMatrix(I, 9))
              txtContato.Text = Trim(.TextMatrix(I, 10))
              txtDataUltimaCompra.Text = Format(.TextMatrix(I, 11), "dd/mm/yyyy")
              txtValorUltimaCompra.Text = Format(.TextMatrix(I, 12), "#,##0.00")
              txtDataMaiorCompra.Text = Format(.TextMatrix(I, 13), "dd/mm/yyyy")
              txtValorMaiorCompra.Text = Format(.TextMatrix(I, 14), "#,##0.00")
              txtCompras.Text = Trim(.TextMatrix(I, 15))
              txtDataIni.Text = Format(.TextMatrix(I, 17), "dd/mm/yyyy")
              txtDataFim.Text = Format(.TextMatrix(I, 18), "dd/mm/yyyy")
              txtComentario.Text = Trim(.TextMatrix(I, 26))
              txtEMail.Text = Trim(.TextMatrix(I, 28))
              cmbOcorrencia.Text = Trim(.TextMatrix(I, 21))

              
              If Trim(.TextMatrix(I, 27)) = "N" Then
                 chkMarketing.Value = 0
              Else
                 chkMarketing.Value = 1
              End If
              
          txtProduto(0).Text = .TextMatrix(I, 22)
          txtProduto(1).Text = .TextMatrix(I, 23)
          txtProduto(2).Text = .TextMatrix(I, 24)
              
          End With
    'Else
      'MsgBox "Fim do Registro.", vbInformation, Me.Caption
        cmdGrava.Refresh
        Exit Sub
    End If
    
'End If
cmdGrava.Refresh

End Sub

Private Sub cmdAvancar_Click()

If grdDados.Rows = 0 Then
   Exit Sub
Else
    cmdGrava.Caption = "Iniciar"
    
    If I <> grdDados.Rows - 1 Then
       grdDados.Row = grdDados.Row + 1
       I = grdDados.Row
       
          With grdDados
              txtNFicha.Text = Format(.TextMatrix(I, 0), "000000")
              txtLoja.Text = .TextMatrix(I, 1)
              txtVendedor.Text = .TextMatrix(I, 2)
              txtCodCliente.Text = .TextMatrix(I, 4)
              txtCliente.Text = Trim(.TextMatrix(I, 5))
              txtTelefone.Text = Trim(.TextMatrix(I, 6))
              txtEndereco.Text = Trim(.TextMatrix(I, 7))
              txtCidade.Text = Trim(.TextMatrix(I, 8))
              cmbUF.Text = Trim(.TextMatrix(I, 9))
              txtContato.Text = Trim(.TextMatrix(I, 10))
              txtDataUltimaCompra.Text = Format(.TextMatrix(I, 11), "dd/mm/yyyy")
              txtValorUltimaCompra.Text = Format(.TextMatrix(I, 12), "#,##0.00")
              txtDataMaiorCompra.Text = Format(.TextMatrix(I, 13), "dd/mm/yyyy")
              txtValorMaiorCompra.Text = Format(.TextMatrix(I, 14), "#,##0.00")
              txtCompras.Text = Trim(.TextMatrix(I, 15))
              txtDataIni.Text = Format(.TextMatrix(I, 17), "dd/mm/yyyy")
              txtDataFim.Text = Format(.TextMatrix(I, 18), "dd/mm/yyyy")
              txtComentario.Text = Trim(.TextMatrix(I, 26))
              txtEMail.Text = Trim(.TextMatrix(I, 28))
              cmbOcorrencia.Text = Trim(.TextMatrix(I, 21))

              
              If Trim(.TextMatrix(I, 27)) = "N" Then
                 chkMarketing.Value = 0
              Else
                 chkMarketing.Value = 1
              End If
              
          txtProduto(0).Text = .TextMatrix(I, 22)
          txtProduto(1).Text = .TextMatrix(I, 23)
          txtProduto(2).Text = .TextMatrix(I, 24)
              
          End With
    Else
        MsgBox "Fim do Registro.", vbInformation, Me.Caption
        cmdGrava.Refresh
        Exit Sub
    
    End If
End If
cmdGrava.Refresh

End Sub

Private Sub cmdGrava_Click()
    
    Dim SQL As String
    
    If Trim(txtCodCliente.Text) = "" Then
       MsgBox "Nenhuma ficha selecionada.", vbInformation, Me.Caption
       txtLoja.SetFocus
       Exit Sub
    End If
    
    If cmdGrava.Caption = "Iniciar" Then
       cmdGrava.Caption = "Gravar"
       txtMascara.Visible = False
       SQL = "Update CRM_Cliente Set crm_TempoInicial = '" & Format(Time, "hh:mm:ss") & "' Where crm_Loja = '" & txtLoja.Text & "' " & _
             "and crm_Vendedor = " & txtVendedor.Text & " and crm_CodigoCliente = " & txtCodCliente.Text
    
       rdoCNMatriz.Execute (SQL)
       cmdAvancar.Enabled = False
       cmdRetornar.Enabled = False
       Call BloqueiaLiberaForm
    Else
       If Trim(cmbOcorrencia.Text) = "" Or Len(Trim(cmbOcorrencia.Text)) = 1 Then
          MsgBox "Informe uma Ocorrência.", vbCritical, Me.Caption
          cmbOcorrencia.SetFocus
          Exit Sub
       End If
        
       If Len(Trim(txtContato.Text)) = 1 Or UCase(Trim(txtContato.Text)) = "Z" Then
          MsgBox "Informe o nome do contato ou deixe o campo em branco", vbCritical, Me.Caption
          txtContato.SetFocus
          Exit Sub
       End If
       
       wRetorna = grdDados.Row
       If chkMarketing.Value = 1 Then
          If Trim(txtEMail.Text) = "" Or LCase(Trim(txtEMail.Text)) = "naotem@naotem" Then
             MsgBox "Informe um E-Mail ou desmarque a opção E-Mail Marketing", vbCritical, Me.Caption
             txtEMail.SetFocus
             Exit Sub
          Else
             wMarketing = "S"
          End If
       Else
          wMarketing = "N"
       End If
       cmdGrava.Caption = "Iniciar"
       Call BloqueiaLiberaForm
       txtMascara.Visible = True
       wOcorrencia = Mid(cmbOcorrencia.Text, 1, 2)
       
       SQL = ""
       SQL = "Update CRM_Cliente Set crm_TempoFinal = '" & Format(Time, "hh:mm:ss") & "', crm_NomeCliente = '" & txtCliente.Text & "', " & _
             "crm_Endereco = '" & txtEndereco.Text & "', crm_Cidade = '" & txtCidade.Text & "', crm_UF = '" & cmbUF.Text & "'," & _
             "crm_Telefone = '" & txtTelefone.Text & "', crm_Contato = '" & txtContato.Text & "', crm_Ocorrencia1 = '" & cmbOcorrencia.Text & "', crm_DataContato = '" & Format(txtData.Text, "yyyy/mm/dd") & "'," & _
             "crm_Comentario = '" & txtComentario.Text & "', crm_Email = '" & txtEMail.Text & "', crm_EMarketing  = '" & wMarketing & "', crm_Status = '" & Mid(cmbOcorrencia.Text, 1, 2) & "' " & _
             "Where crm_Loja = '" & txtLoja.Text & "' and crm_Vendedor = " & txtVendedor.Text & " and crm_CodigoCliente = " & txtCodCliente.Text
       
       rdoCNMatriz.Execute (SQL)
        
       grdDados.RemoveItem wRetorna
       I = grdDados.Row

         If I >= 0 Then
            With grdDados
              txtNFicha.Text = Format(.TextMatrix(I, 0), "000000")                        'CRM_NumeroFicha
              txtLoja.Text = .TextMatrix(I, 1)                                            'CRM_Loja
              txtVendedor.Text = .TextMatrix(I, 2)                                        'CRM_Vendedor
              txtCodCliente.Text = .TextMatrix(I, 4)                                      'CRM_CodigoCliente
              txtCliente.Text = Trim(.TextMatrix(I, 5))                                   'CRM_NomeCliente
              txtTelefone.Text = Trim(.TextMatrix(I, 6))                                  'CRM_Telefone
              txtEndereco.Text = Trim(.TextMatrix(I, 7))                                  'CRM_Endereco
              txtCidade.Text = Trim(.TextMatrix(I, 8))                                    'CRM_Cidade
              cmbUF.Text = Trim(.TextMatrix(I, 9))                                        'CRM_UF
              txtContato.Text = Trim(.TextMatrix(I, 10))                                  'CRM_Contato
              txtDataUltimaCompra.Text = Format(.TextMatrix(I, 11), "dd/mm/yyyy")         'CRM_DataUltimaCompra
              txtValorUltimaCompra.Text = Format(.TextMatrix(I, 12), "#,##0.00")          'CRM_ValorUltimaCompra
              txtDataMaiorCompra.Text = Format(.TextMatrix(I, 13), "dd/mm/yyyy")          'CRM_DataMaiorCompra
              txtValorMaiorCompra.Text = Format(.TextMatrix(I, 14), "#,##0.00")           'CRM_ValorMaiorCompra
              txtCompras.Text = .TextMatrix(I, 15)                                        'CRM_QtdeCompras
              txtDataIni.Text = Format(.TextMatrix(I, 17), "dd/mm/yyyy")                  'CRM_PeriodoInicial
              txtDataFim.Text = Format(.TextMatrix(I, 18), "dd/mm/yyyy")                  'CRM_PeriodoFinal
              txtComentario.Text = Trim(.TextMatrix(I, 26))                               'CRM_Comentario
              txtEMail.Text = Trim(.TextMatrix(I, 28))                                    'CRM_Email
              cmbOcorrencia.Text = Trim(.TextMatrix(I, 21))                               'CRM_Ocorrencia1
              If Trim(.TextMatrix(I, 27)) = "N" Then                                      'CRM_Emarketing
                 chkMarketing.Value = 0
              Else
                 chkMarketing.Value = 1
              End If
              wRetorna = I
              txtProduto(0).Text = .TextMatrix(I, 22)
              txtProduto(1).Text = .TextMatrix(I, 23)
              txtProduto(2).Text = .TextMatrix(I, 24)
            End With
         Else
            MsgBox "Não existem mais clientes para contato.", vbInformation, Me.Caption
            Call LimpaForm
            txtLoja.SetFocus
          End If
         
       ProximoCliente  ' MUDANÇA
       'Enter           ' MUDANÇA
       
       cmdAvancar.Enabled = True
       cmdRetornar.Enabled = True
    End If
    
     
End Sub


Private Sub cmdGravarProduto_Click()
  If Trim(txtLoja.Text) = "" Then
     MsgBox "Informe uma Loja.", vbCritical, Me.Caption
     txtLoja.SetFocus
     Exit Sub
  End If
     
  Frame1.Visible = True
  txtDataIniCrm.Enabled = True
  txtDataFimCrm.Enabled = True
  txtDataIniCrm.SetFocus
  
End Sub

Private Sub cmdImprimir_Click()
  If Trim(txtCodCliente.Text) = "" Then
    MsgBox "Nenhuma ficha selecionada.", vbInformation, Me.Caption
    txtLoja.SetFocus
  Else
    ImprimeFichaCliente
  End If
  
End Sub

Private Sub CmdOK_Click()
'Dim SQL As String
'
'  If IsDate(txtDataIniCrm.Text) = False Then
'     MsgBox "Data Inválida.", vbCritical, Me.Caption
'     txtDataIniCrm.SetFocus
'     Exit Sub
'  End If
'
'  If IsDate(txtDataFimCrm.Text) = False Then
'     MsgBox "Data Inválida.", vbCritical, Me.Caption
'     txtDataFimCrm.SetFocus
'     Exit Sub
'  End If
'
'Screen.MousePointer = vbHourglass
'cmdOk.Enabled = False
'
'  SQL = ""
'  SQL = "SELECT LO_Loja FROM Loja Where LO_Situacao = 'A' and LO_Loja = '" & Trim(txtLoja.Text) & "'"
'  Set RdoRsDados = rdoCnSup.OpenResultset(SQL)
'
'    If RdoRsDados.EOF = False Then
'       RdoRsDados.Close
'
'      If Trim(txtLoja.Text) = "271" Or Trim(txtLoja.Text) = "315" Then
'         wLoja = "315','271"
'      ElseIf Trim(txtLoja.Text) = "396" Or Trim(txtLoja.Text) = "506" Then
'         wLoja = "396','506"
'      Else
'         wLoja = txtLoja.Text
'      End If
'
'      SQL = ""
''      SQL = "Insert Into CRM_Cliente(CRM_CodigoCliente,CRM_Loja,CRM_Vendedor,CRM_NomeCliente,CRM_PeriodoInicial,CRM_PeriodoFinal," & _
'            "CRM_Telefone,CRM_Endereco,CRM_Cidade,CRM_DataUltimaCompra,CRM_EMail) Select Distinct (VC_Cliente),'" & Trim(txtLoja.Text) & "'," & _
'            "VC_CodigoVendedor,VC_NomeCliente,'" & Format(txtDataIniCrm.Text, "yyyy/mm/dd") & "','" & Format(txtDataFimCrm.Text, "yyyy/mm/dd") & "',VC_TelefoneCliente," & _
'            "VC_EnderecoCliente,VC_MunicipioCliente,VC_DataEmissao,'naotem@naotem' From CapaNFVenda " & _
'            "Where VC_DataEmissao Between '" & Format(txtDataIniCrm.Text, "yyyy/mm/dd") & "' and '" & Format(txtDataFimCrm.Text, "yyyy/mm/dd") & "' and VC_Tiponota='V' and vc_lojavenda in('" & wLoja & "') " & _
'            "and VC_Cliente <> 999999 and VC_TotalNota > 99.99 and VC_Cliente > 900000"
'
'      SQL = "INSERT INTO CRM_CLIENTE (CRM_CodigoCliente,CRM_Loja,CRM_Vendedor,CRM_NomeCliente,CRM_PeriodoInicial,CRM_PeriodoFinal," & _
'            "CRM_Telefone,CRM_Endereco,CRM_Cidade) Select VC_Cliente,'" & Trim(txtLoja.Text) & "',VC_CodigoVendedor," & _
'            "VC_NomeCliente,'" & Format(txtDataIniCrm.Text, "yyyy/mm/dd") & "','" & Format(txtDataFimCrm.Text, "yyyy/mm/dd") & "',VC_TelefoneCliente," & _
'            "VC_EnderecoCliente,VC_MunicipioCliente FROM CapanfVenda " & _
'            "WHERE VC_DataEmissao Between '" & Format(txtDataIniCrm.Text, "yyyy/mm/dd") & "' and '" & Format(txtDataFimCrm.Text, "yyyy/mm/dd") & "' and VC_TipoNota='V'" & _
'            "and VC_LojaVenda in('" & wLoja & "') and VC_Cliente <> 999999 and VC_TotalNota > 99.99 and VC_Cliente < 900000 " & _
'            "Group BY VC_Cliente,VC_CodigoVendedor,VC_NomeCliente,VC_TelefoneCliente,VC_EnderecoCliente,VC_MunicipioCliente"
'
'      rdoCnSup.Execute (SQL)
'
'      SQL = ""
'      SQL = "Update CRM_CLIENTE SET CRM_DataUltimaCompra = VC_DataEmissao,CRM_ValorUltimaCompra = VC_TotalNota," & _
'            "CRM_DataMaiorCompra=VC_DataEmissao,CRM_ValorMaiorCompra = VC_TotalNota, CRM_QtdeCompras = 1 " & _
'            "From CRM_Cliente, CapaNFVenda Where CRM_CodigoCliente = VC_Cliente and CRM_Loja = '" & Trim(txtLoja.Text) & "' and " & _
'            "VC_LojaVenda in('" & wLoja & "')"
'
'      wSalvaTimeOut = rdoCnSup.QueryTimeout
'
'      rdoCnSup.QueryTimeout = 90
'      rdoCnSup.Execute (SQL)
'
'      rdoCnSup.QueryTimeout = wSalvaTimeOut
'
'       Call AtualizaProdutoCRM
'       If wQtdRegistros = 0 Then
'          MsgBox "Atualização concluída com sucesso.", vbInformation, "Atualiza Produto CRM"
'       Else
'          MsgBox "Atualização concluída." & vbLf & _
'                 wQtdRegistros & " não foram atualizados.", vbInformation, "Atualiza Produto CRM"
'       End If
'
'    Else
'       MsgBox "Loja não cadastrada.", vbCritical, Me.Caption
'       RdoRsDados.Close
'       txtLoja.SetFocus
'    End If
'
'    Frame1.Visible = False
'
'txtDataIniCrm.Text = "__/__/____"
'txtDataFimCrm.Text = "__/__/____"
'cmdOk.Enabled = True
'txtDataIniCrm.Enabled = False
'txtDataFimCrm.Enabled = False
'Screen.MousePointer = vbNormal
    
End Sub
Public Sub montaComboLoja(comboLojas As ComboBox)
'ricardo


'        SQL = "SELECT GLB_lOJA FROM CONEXAOSISTEMA"
'
'
'        RdoRsloja.CursorLocation = adUseClient
'        RdoRsloja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'        Do While Not .EOF
'
'            comboLoja.AddItem Trim(RdoRsloja("GLB_lOJA"))
'            RdoRsloja.MoveNext
'        Loop
'        RdoRsloja.Close

On Error GoTo trataerro
    Dim RdoRsloja As New ADODB.Recordset
    Dim SQL As String
    With adoCNLoja
    
    
      SQL = ""
      SQL = "select CTS_lOJA from ControleSistema"
        
        RdoRsloja.CursorLocation = adUseClient
        RdoRsloja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
        Do While Not RdoRsloja.EOF
            
            'comboLoja.AddItem Trim(RdoRsloja("CTS_lOJA"))
            RdoRsloja.MoveNext
        Loop
        RdoRsloja.Close
    End With

    Exit Sub
trataerro:
    Select Case Err.Number
        Case Else
            'mensagemErroDesconhecido Err, "Erro na leitura de lista de lojas"
    End Select
End Sub

Private Sub cmbLoja_KeyPress(KeyAscii As Integer)
'ricardo
'If KeyAscii = 13 Then
'    cmdPesquisa_Click
'End If

End Sub

Private Sub PreencheComboLojas(comboLoja As ComboBox)

Dim RdoRslojas As New ADODB.Recordset

Dim SQL As String
'ricardo
'  SQL = ""
'  SQL = "select lo_loja from loja where lo_situacao = 'A' and " & _
'         "lo_regiao < 402 order by lo_regiao"
'
'
'        RdoRslojas.CursorLocation = adUseClient
'        RdoRslojas.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
                                       
                                       
    SQL = ""
    SQL = "select CTS_lOJA from ControleSistema"
        
        RdoRslojas.CursorLocation = adUseClient
        RdoRslojas.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
        Do While Not RdoRslojas.EOF
            
            comboLoja.AddItem Trim(RdoRslojas("CTS_lOJA"))
            RdoRslojas.MoveNext
        Loop
        RdoRslojas.Close
        'cmbLoja.ListIndex = 0
        
'     'cmbLojas.AddItem "181"
'     Do While Not RdoRslojas.EOF
'        cmbLojas.AddItem Trim(RdoRslojas("lo_Loja"))
'        RdoRslojas.MoveNext
'     Loop
'        RdoRslojas.Close
'        cmbLojas.AddItem ""
'
'        cmbLojas.ListIndex = 0
         
End Sub
Private Sub PreencheTXTLoja(txtLoja As TextBox)

Dim RdoRstxtloja As New ADODB.Recordset

Dim SQL As String
'ricardo

                                                                     
    SQL = ""
    SQL = "select CTS_lOJA from ControleSistema"
        
        RdoRstxtloja.CursorLocation = adUseClient
        RdoRstxtloja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
        Do While Not RdoRstxtloja.EOF
            
            txtLoja.Text = Trim(RdoRstxtloja("CTS_lOJA"))
            RdoRstxtloja.MoveNext
        Loop
        RdoRstxtloja.Close
       
        

         
End Sub
Private Sub cmdPesquisa_Click()
'ricardo
Dim SQL As String
Dim rdoLoja As New ADODB.Recordset

'    If Trim(cmbLoja.Text) = "" Then
'    MsgBox "Selecione uma Loja.", vbCritical, Me.Caption
'    cmbLoja.SetFocus
'    Exit Sub
'End If
  
    SQL = ""
      SQL = "select CTS_lOJA from ControleSistema"
        
        rdoLoja.CursorLocation = adUseClient
        rdoLoja.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
        Do While Not rdoLoja.EOF
            
            txtLoja.Text = Trim(rdoLoja("CTS_lOJA"))
            rdoLoja.MoveNext
        Loop
        rdoLoja.Close
  
  
  grdDados1.Rows = 1
  wTotalFichas = 0
  wTotalFichasAbertas = 0
  wTotalFichasFechadas = 0
  wTotalFichasRetorno = 0
      
  Screen.MousePointer = vbHourglass
    
      SQL = ""
      SQL = "Select Distinct CRM_Vendedor, Count(*) as TotalRegistros From CRM_Cliente " & _
            "Where CRM_Loja = '" & txtLoja.Text & "' Group By CRM_Vendedor Order By CRM_Vendedor"
            
      'Set RdoRsVendedor = rdoCNMatriz.OpenResultset(SQL)
      RdoRsVendedor.CursorLocation = adUseClient
      RdoRsVendedor.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
      
      
      
      If RdoRsVendedor.EOF = False Then
         Do While Not RdoRsVendedor.EOF
            '************ Total de Fichas ************
            wTotalFichas = RdoRsVendedor("TotalRegistros")
            '************ Total de Fichas Abertas ************
            SQL = ""
            SQL = "Select Count(*) as Abertas From CRM_Cliente " & _
                  "Where crm_Loja = '" & txtLoja.Text & "' and crm_Vendedor = " & RdoRsVendedor("CRM_Vendedor") & " and crm_Status = 'A'"
                       
            'Set RdoRsDados = rdoCNMatriz.OpenResultset(SQL)
             RdoRsDados.CursorLocation = adUseClient
             RdoRsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
            
            If RdoRsDados.EOF = False Then
               wTotalFichasAbertas = RdoRsDados("Abertas")
            End If
            RdoRsDados.Close
            
            '************ Total de Fichas Finalizadas ************
            SQL = ""
            SQL = "Select Count(*) as Fechadas From CRM_Cliente " & _
                  "Where crm_Loja = '" & txtLoja.Text & "' and crm_Vendedor = " & RdoRsVendedor("CRM_Vendedor") & " and crm_Status not in('A','VL','CR')"
                       
            'Set RdoRsDados = rdoCNMatriz.OpenResultset(SQL)
            RdoRsDados.CursorLocation = adUseClient
            RdoRsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
            
            If RdoRsDados.EOF = False Then
               wTotalFichasFechadas = RdoRsDados("Fechadas")
            End If
            RdoRsDados.Close
            
            '************ Total de Fichas Aguardando Retorno ************
            SQL = ""
            SQL = "Select Count(*) as Retorna From CRM_Cliente " & _
                  "Where crm_Loja = '" & txtLoja.Text & "' and crm_Vendedor = " & RdoRsVendedor("CRM_Vendedor") & " and crm_Status in('VL','CR')"
                       
            'Set RdoRsDados = rdoCNMatriz.OpenResultset(SQL)
            RdoRsDados.CursorLocation = adUseClient
            RdoRsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
            
            If RdoRsDados.EOF = False Then
               wTotalFichasRetorno = RdoRsDados("Retorna")
            End If
            RdoRsDados.Close
            
            grdDados1.AddItem RdoRsVendedor("CRM_Vendedor") & Chr(9) & wTotalFichas & Chr(9) & wTotalFichasAbertas & _
                             Chr(9) & wTotalFichasFechadas & Chr(9) & wTotalFichasRetorno
            
            RdoRsVendedor.MoveNext
         Loop
         
      Else
         MsgBox "Nenhum Registro Encontrado.", vbInformation, Me.Caption
         RdoRsVendedor.Close
         'cmbLoja.SetFocus
         Screen.MousePointer = vbNormal
         Exit Sub
      End If
            
        
        RdoRsVendedor.Close
        'cmdImprimir.Enabled = True
        Screen.MousePointer = vbNormal
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub cmdRetornar_Click()
If grdDados.Rows = 0 Then
   Exit Sub
Else
   cmdGrava.Caption = "Iniciar"
    I = grdDados.Row
    If 0 <> I Then
       grdDados.Row = grdDados.Row - 1
       I = grdDados.Row
          
          With grdDados
              txtNFicha.Text = Format(.TextMatrix(I, 0), "000000")
              txtLoja.Text = Trim(.TextMatrix(I, 1))
              txtVendedor.Text = Trim(.TextMatrix(I, 2))
              txtCodCliente.Text = Trim(.TextMatrix(I, 4))
              txtCliente.Text = Trim(.TextMatrix(I, 5))
              txtTelefone.Text = Trim(.TextMatrix(I, 6))
              txtEndereco.Text = Trim(.TextMatrix(I, 7))
              txtCidade.Text = Trim(.TextMatrix(I, 8))
              cmbUF.Text = Trim(.TextMatrix(I, 9))
              txtContato.Text = Trim(.TextMatrix(I, 10))
              txtDataUltimaCompra.Text = Format(.TextMatrix(I, 11), "dd/mm/yyyy")
              txtValorUltimaCompra.Text = Format(.TextMatrix(I, 12), "#,##0.00")
              txtDataMaiorCompra.Text = Format(.TextMatrix(I, 13), "dd/mm/yyyy")
              txtValorMaiorCompra.Text = Format(.TextMatrix(I, 14), "#,##0.00")
              txtCompras.Text = .TextMatrix(I, 15)
              txtDataIni.Text = Format(.TextMatrix(I, 17), "dd/mm/yyyy")
              txtDataFim.Text = Format(.TextMatrix(I, 18), "dd/mm/yyyy")
              txtComentario.Text = Trim(.TextMatrix(I, 26))
              txtEMail.Text = Trim(.TextMatrix(I, 28))
              cmbOcorrencia.Text = Trim(.TextMatrix(I, 21))
              
              If UCase(Trim(.TextMatrix(I, 27))) = "N" Then
                 chkMarketing.Value = 0
              Else
                 chkMarketing.Value = 1
              End If
              
          txtProduto(0).Text = .TextMatrix(I, 22)
          txtProduto(1).Text = .TextMatrix(I, 23)
          txtProduto(2).Text = .TextMatrix(I, 24)
              
          End With
    Else
        MsgBox "Inicio do Registro.", vbInformation, Me.Caption
        cmdGrava.Refresh

        Exit Sub
    
    End If
End If
cmdGrava.Refresh

End Sub



Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
    cmdPesquisa_Click
End Sub

Private Sub Form_Click()
    frmDescricao.Visible = False
End Sub

Private Sub Form_Load()
Dim SQL As String
'Dim adoCNLoja As ADODB.Connection
'Dim rdoCnSup As New adodb.Connection

ConectaODBCMatriz
  Me.top = 4680
  Me.left = 90
  Me.Width = 15180
  Me.Height = 5790
 With cmbOcorrencia
    .AddItem ""
    .AddItem "NL - Não Localizado"
    .AddItem "VL - Voltar a Ligar"
    .AddItem "CR - Cliente Retorna"
    .AddItem "OO - Outras Ocorrências"
    .AddItem "NA - Não Atendeu s/ Retorno"
    .AddItem "CO - Contatado OK"
 End With

Call LimpaForm
Call PreencheComboUF
Call BloqueiaLiberaForm

SQL = ""
SQL = "Select * From Usuario Where US_Grupo = 13  and US_Nome = 'gerentes'"

     RdoRsDados.CursorLocation = adUseClient
     RdoRsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic


    'Set RdoRsDados = rdoCNMatriz.OpenResultset(SQL)
    If RdoRsDados.EOF = False Then
        cmdGravarProduto.Visible = True
    Else
        cmdGravarProduto.Visible = False
    End If
    RdoRsDados.Close
    
    'ricardo
    
    'cmbLoja.Enabled = True
    
    'PreencheComboLojas cmbLoja
    'PreencheTXTLoja txtLoja
    'montaComboLoja cmbLoja
    
    frmReativacaoCliente.txtVendedor.Text = frmPedido.txtVendedor.Text
    'PreencheCampos (wWhere)
    txtVendedor_KeyPress 13
    frmDescricao.Visible = False
    
    cmdPesquisa_Click
    
End Sub

Private Sub PreencheComboUF()
Dim SQL As String
    cmbUF.AddItem ""
    
    SQL = ""
    SQL = "select UF_Estado from Estados Order By UF_Estado"
    
     RdoRsDados.CursorLocation = adUseClient
     RdoRsDados.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    'Set RdoRsDados = adoCNLoja.OpenResultset("Select UF_Estado From Estados Order By UF_Estado")
     If RdoRsDados.EOF = False Then
        Do While Not RdoRsDados.EOF
           cmbUF.AddItem RdoRsDados("UF_Estado")
           RdoRsDados.MoveNext
        Loop
     End If
     RdoRsDados.Close
     cmbUF.Text = "SP"
     
End Sub

Private Sub frmDescricao_DblClick()
    frmDescricao.Visible = False
End Sub

Private Sub txtDataFimCrm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtLoja.Text) <> "" Then
       cmdGravarProduto.SetFocus
    End If
End Sub

Private Sub txtDataIniCrm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtLoja.Text) <> "" Then
       txtDataFimCrm.SetFocus
    End If
End Sub

Private Sub txtLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtLoja.Text) <> "" Then
       txtVendedor.SetFocus
    End If

End Sub

Private Sub AtualizaProdutoCRM()
Dim wProdutoCRM1 As String
Dim wProdutoCRM2 As String
Dim wProdutoCRM3 As String
Dim wCodigoCliente As Long
Dim wCodigoVendedor As Integer
Dim SQL As String


On Error GoTo ErroRegistro

    SQL = ""
    SQL = "Select CRM_Vendedor, CRM_CodigoCliente, CRM_PeriodoInicial, CRM_PeriodoFinal From CRM_Cliente " & _
          "Where CRM_Loja = '" & txtLoja.Text & "'"
          
    Set RdoRsDados = adoCNLoja.OpenResultset(SQL)
      
    Do While Not RdoRsDados.EOF
        wCodigoCliente = RdoRsDados("CRM_CodigoCliente")
        wCodigoVendedor = RdoRsDados("CRM_Vendedor")
    
        SQL = ""
        SQL = "Select Top 3 PR_Referencia, PR_Descricao, VI_PrecoLista " & _
              "From CapanfVenda, ItemNfVenda, Produto " & _
              "Where vi_referencia = pr_referencia And vc_serie = vi_serie And vc_tiponota = vi_tiponota " & _
              "and vc_notafiscal = vi_notafiscal and vc_dataemissao = vi_dataemissao " & _
              "and vc_lojaorigem = vi_lojaorigem " & _
              "and vc_dataemissao between '" & Format(RdoRsDados("CRM_PeriodoInicial"), "yyyy/mm/dd") & "' and '" & Format(RdoRsDados("CRM_PeriodoFinal"), "yyyy/mm/dd") & "' " & _
              "and vc_vendedorlojavenda = " & wCodigoVendedor & " and vc_cliente = " & wCodigoCliente
        
        
        Set RdoRsCRM = adoCNLoja.OpenResultset(SQL)
        wProdutoCRM1 = ""
        wProdutoCRM2 = ""
        wProdutoCRM3 = ""
        
        If Not RdoRsCRM.EOF Then
           wProdutoCRM1 = RdoRsCRM("PR_Referencia") & "  " & RdoRsCRM("PR_Descricao") & "  " & Format(RdoRsCRM("VI_PrecoLista"), "#,##0.00")
           RdoRsCRM.MoveNext
        End If
        
        If Not RdoRsCRM.EOF Then
           wProdutoCRM2 = RdoRsCRM("PR_Referencia") & "  " & RdoRsCRM("PR_Descricao") & "  " & Format(RdoRsCRM("VI_PrecoLista"), "#,##0.00")
           RdoRsCRM.MoveNext
        End If
            
        If Not RdoRsCRM.EOF Then
           wProdutoCRM3 = RdoRsCRM("PR_Referencia") & "  " & RdoRsCRM("PR_Descricao") & "  " & Format(RdoRsCRM("VI_PrecoLista"), "#,##0.00")
           RdoRsCRM.MoveNext
        End If
        
        RdoRsCRM.Close
        SQL = ""
        SQL = "UPDATE CRM_Cliente SET CRM_Ocorrencia2 = '" & wProdutoCRM1 & "', CRM_Ocorrencia3 = '" & wProdutoCRM2 & "', CRM_Ocorrencia4 = '" & wProdutoCRM3 & "' WHERE " & _
              "CRM_CodigoCliente = " & RdoRsDados("CRM_CodigoCliente") & " and CRM_Vendedor = " & RdoRsDados("CRM_Vendedor") & " and " & _
              "CRM_Loja = '" & Trim(txtLoja.Text) & "'"
        
        adoCNLoja.Execute (SQL)
        
        RdoRsDados.MoveNext
        
        DoEvents
    Loop
          
    RdoRsDados.Close
    Exit Sub

ErroRegistro:
    wQtdRegistros = wQtdRegistros + 1
    Resume Next
    
End Sub

Private Sub txtMascara_GotFocus()
    txtContato.SetFocus
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
Dim rdoLojas As New ADODB.Recordset
Dim SQL As String


    Screen.MousePointer = 11
        
'    If KeyAscii = vbKeyReturn Then
'       If Trim(txtLoja.Text) = "" Then
'          MsgBox "Informe uma Loja.", vbCritical, Me.Caption
'          'txtLoja.SetFocus
'          Exit Sub
'       End If

      SQL = ""
      SQL = "select CTS_lOJA from ControleSistema"
        
        rdoLojas.CursorLocation = adUseClient
        rdoLojas.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
        Do While Not rdoLojas.EOF
            
            txtLoja.Text = Trim(rdoLojas("CTS_lOJA"))
            rdoLojas.MoveNext
        Loop
        rdoLojas.Close



       
       If Trim(txtVendedor.Text) = "" Or IsNumeric(txtVendedor.Text) = False Then
          MsgBox "Informe o Código do Vendedor.", vbCritical, Me.Caption
          txtVendedor.SetFocus
          Exit Sub
       End If
       wWhere = "Select * From CRM_Cliente Where crm_Status in('A','VL','CR')"
       
       If txtLoja.Text <> "999" And txtVendedor.Text <> "999" Then
          wWhere = wWhere & " and crm_Vendedor = " & txtVendedor.Text & " and crm_Loja = '" & txtLoja.Text & "'"
       ElseIf txtLoja.Text <> "999" And txtVendedor.Text = "999" Then
          wWhere = wWhere & " and crm_Loja = '" & txtLoja.Text & "'"
       ElseIf txtLoja = "999" And txtVendedor.Text <> "999" Then
          wWhere = wWhere & " and crm_Vendedor = " & txtVendedor.Text
       End If
       
       PreencheCampos (wWhere)
    'End If
    Screen.MousePointer = 0

End Sub


Private Sub PreencheCampos(wWhere As String)
grdDados.Rows = 0
    
     RdoRsDados.CursorLocation = adUseClient
     RdoRsDados.Open wWhere, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
          
     If RdoRsDados.EOF = False Then
          
          I = 0
          Do While Not RdoRsDados.EOF
             I = I + 1
             grdDados.AddItem RdoRsDados("CRM_NumeroFicha") & Chr(9) & RdoRsDados("CRM_Loja") & Chr(9) & _
                      RdoRsDados("CRM_Vendedor") & Chr(9) & RdoRsDados("CRM_ItemVendedor") & Chr(9) & RdoRsDados("CRM_CodigoCliente") & Chr(9) & _
                      RdoRsDados("CRM_NomeCliente") & Chr(9) & RdoRsDados("CRM_Telefone") & Chr(9) & RdoRsDados("CRM_Endereco") & Chr(9) & RdoRsDados("CRM_Cidade") & Chr(9) & _
                      RdoRsDados("CRM_UF") & Chr(9) & RdoRsDados("CRM_Contato") & Chr(9) & RdoRsDados("CRM_DataUltimaCompra") & Chr(9) & RdoRsDados("CRM_ValorUltimaCompra") & Chr(9) & RdoRsDados("CRM_DataMaiorCompra") & Chr(9) & _
                      RdoRsDados("CRM_ValorMaiorCompra") & Chr(9) & RdoRsDados("CRM_QtdeCompras") & Chr(9) & RdoRsDados("CRM_ValorTotalCompras") & Chr(9) & RdoRsDados("CRM_PeriodoInicial") & Chr(9) & RdoRsDados("CRM_PeriodoFinal") & Chr(9) & _
                      RdoRsDados("CRM_TempoInicial") & Chr(9) & RdoRsDados("CRM_TempoFinal") & Chr(9) & RdoRsDados("CRM_Ocorrencia1") & Chr(9) & RdoRsDados("CRM_Ocorrencia2") & Chr(9) & RdoRsDados("CRM_Ocorrencia3") & Chr(9) & RdoRsDados("CRM_Ocorrencia4") & Chr(9) & _
                      Format(RdoRsDados("CRM_DataContato"), "dd/mm/yyyy") & Chr(9) & RdoRsDados("CRM_Comentario") & Chr(9) & RdoRsDados("CRM_Emarketing") & Chr(9) & RdoRsDados("CRM_Email") & Chr(9) & _
                      RdoRsDados("CRM_Status") & Chr(9) & RdoRsDados("CRM_Situacao") & Chr(9) & I
          
             RdoRsDados.MoveNext
          Loop
          RdoRsDados.Close
                
          grdDados.Select 0, 0
          I = grdDados.Row
          
          With grdDados
              txtNFicha.Text = Format(.TextMatrix(I, 0), "000000")                        'CRM_NumeroFicha
              txtLoja.Text = .TextMatrix(I, 1)                                            'CRM_Loja
              txtVendedor.Text = .TextMatrix(I, 2)                                        'CRM_Vendedor
              txtCodCliente.Text = .TextMatrix(I, 4)                                      'CRM_CodigoCliente
              txtCliente.Text = Trim(.TextMatrix(I, 5))                                   'CRM_NomeCliente
              txtTelefone.Text = Trim(.TextMatrix(I, 6))                                  'CRM_Telefone
              txtEndereco.Text = Trim(.TextMatrix(I, 7))                                  'CRM_Endereco
              txtCidade.Text = Trim(.TextMatrix(I, 8))                                    'CRM_Cidade
              cmbUF.Text = Trim(.TextMatrix(I, 9))                                        'CRM_UF
              txtContato.Text = Trim(.TextMatrix(I, 10))                                  'CRM_Contato
              txtDataUltimaCompra.Text = Format(.TextMatrix(I, 11), "dd/mm/yyyy")         'CRM_DataUltimaCompra
              txtValorUltimaCompra.Text = Format(.TextMatrix(I, 12), "#,##0.00")          'CRM_ValorUltimaCompra
              txtDataMaiorCompra.Text = Format(.TextMatrix(I, 13), "dd/mm/yyyy")          'CRM_DataMaiorCompra
              txtValorMaiorCompra.Text = Format(.TextMatrix(I, 14), "#,##0.00")           'CRM_ValorMaiorCompra
              txtCompras.Text = .TextMatrix(I, 15)                                        'CRM_QtdeCompras
              txtDataIni.Text = Format(.TextMatrix(I, 17), "dd/mm/yyyy")                  'CRM_PeriodoInicial
              txtDataFim.Text = Format(.TextMatrix(I, 18), "dd/mm/yyyy")                  'CRM_PeriodoFinal
              txtComentario.Text = Trim(.TextMatrix(I, 26))                               'CRM_Comentario
              txtEMail.Text = Trim(.TextMatrix(I, 28))                                    'CRM_Email
              cmbOcorrencia.Text = Trim(.TextMatrix(I, 21))                               'CRM_Ocorrencia1
              
              If Trim(.TextMatrix(I, 27)) = "N" Then                                      'CRM_Emarketing
                 chkMarketing.Value = 0
              Else
                 chkMarketing.Value = 1
              End If
              wRetorna = I
          
          txtProduto(0).Text = .TextMatrix(I, 22)
          txtProduto(1).Text = .TextMatrix(I, 23)
          txtProduto(2).Text = .TextMatrix(I, 24)
          End With
          Call PintaText
          txtData.Text = Format(Date, "dd/mm/yyyy")
          txtMascara.Visible = True
          
      Else
        MsgBox "Nenhum Registro Encontrado.", vbInformation, Me.Caption
        Call LimpaForm
        'txtLoja.SetFocus
        Exit Sub
      End If

      
End Sub

Private Sub PintaText()
'      txtNFicha.BackColor = vbYellow
'      txtDataIni.BackColor = vbYellow
'      txtDataFim.BackColor = vbYellow
'      txtCompras.BackColor = vbYellow
'      txtDataUltimaCompra.BackColor = vbYellow
'      txtValorUltimaCompra.BackColor = vbYellow
'      txtDataMaiorCompra.BackColor = vbYellow
'      txtValorMaiorCompra.BackColor = vbYellow
'      txtCodCliente.BackColor = vbYellow
'      txtProduto(0).BackColor = vbYellow
'      txtProduto(1).BackColor = vbYellow
'      txtProduto(2).BackColor = vbYellow
'
'      txtData.BackColor = vbYellow
      txtData.Enabled = False
      
End Sub

Private Sub LimpaForm()
Dim cObjeto As Control
  
  For Each cObjeto In Me.Controls
      
      If (TypeOf cObjeto Is TextBox) Then
        cObjeto.Text = ""
        cObjeto.BackColor = &H80000005
      End If
      
      If (TypeOf cObjeto Is CheckBox) Then
        cObjeto.Value = 0
      End If
      
      If (TypeOf cObjeto Is ComboBox) Then
        cObjeto.Text = ""
      End If
      
      If (TypeOf cObjeto Is MaskEdBox) Then
        cObjeto.Text = "__/__/____"
      End If
  
  Next
    
    cmbUF.Text = "SP"
    grdDados.Rows = 0
    grdDados.Visible = False
    
    
End Sub

Private Sub ImprimeFichaCliente()
  

    Printer.FontName = "ARIAL"
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 5
    Printer.CurrentX = 0
    Printer.CurrentY = 0


Printer.Print
Printer.Print Tab(14); "Nº Ficha: " & txtNFicha.Text


    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.FontName = "COURIER NEW"
    Printer.Print

Printer.Print Tab(10); "Loja: " & txtLoja.Text; Tab(22); "Vendedor: " & txtVendedor.Text; Tab(37)
Printer.Print
Printer.Print Tab(10); "Cod.Cliente: " & txtCodCliente.Text
Printer.Print Tab(10); "Nome Cliente: " & txtCliente.Text

Printer.Print Tab(10); "Endereço : " & txtEndereco.Text; Tab(54);
Printer.Print Tab(10); "Cidade: " & txtCidade.Text; Tab(45); "UF: " & cmbUF.Text
Printer.Print Tab(10); "Telefone : " & txtTelefone.Text; Tab(35); "Contato: " & txtContato.Text
Printer.Print
Printer.Print Tab(10); "Última Compra: " & Format(txtDataUltimaCompra.Text, "dd/mm/yyyy")
Printer.Print Tab(10); "Valor: " & Format(txtValorUltimaCompra.Text, "###,##0.00"); Tab(32); "Maior Compra: " & Format(txtValorMaiorCompra.Text, "###,##0.00"); Tab(58); "Nº Compras: " & txtCompras.Text

Printer.Print Tab(10); "Período: " & Format(txtDataIni.Text, "dd/mm/yyyy") & " Até " & Format(txtDataFim.Text, "dd/mm/yyyy") '; Tab(50); "Ocorrência: " & cmbOcorrencia.Text

Printer.Print Tab(10); "Ocorrência: " & cmbOcorrencia.Text
Printer.Print
Printer.Print Tab(10); "Referencia/Descricao"
Printer.Print Tab(10); "Desativado(  )  " & "Mudou(  )  " & "Sem Interesse(  )  " & "Re Ativado(  )"
Printer.Print Tab(10); txtProduto(0).Text
Printer.Print Tab(10); txtProduto(1).Text
Printer.Print Tab(10); txtProduto(2).Text

Printer.Print
Printer.Print Tab(10); "Comentário: " & Mid(txtComentario.Text, 1, 250)
Printer.Print Tab(22); Mid(txtComentario.Text, 51, 100)
Printer.Print Tab(22); Mid(txtComentario.Text, 101, 150)
Printer.Print Tab(22); Mid(txtComentario.Text, 151, 200)
Printer.Print
Printer.Print Tab(10); "       _________________________________________________________________________________"
Printer.Print
Printer.Print Tab(10); "       _________________________________________________________________________________"
Printer.Print

If chkMarketing.Value = 0 Then
   Printer.Print Tab(10); "E-Mail: " & txtEMail.Text; Tab(50); "E-Mail Marketing --> Não"
Else
   Printer.Print Tab(10); "E-Mail: " & txtEMail.Text; Tab(50); "E-Mail Marketing --> Sim "
End If
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); "       ___________              _________________________________________"
Printer.Print Tab(10); "           Data " & "                             Gerente "
'Printer.Print Tab(1); "__________________________________________________________________________________________"


Printer.EndDoc

End Sub

Private Sub BloqueiaLiberaForm()
Dim cObjeto As Control
Dim wStatus As Boolean
  
  If cmdGrava.Caption = "Iniciar" Then
    wStatus = False
  Else
    wStatus = True
  End If

  For Each cObjeto In Me.Controls
     If (TypeOf cObjeto Is TextBox) Then
       cObjeto.Enabled = wStatus
     End If
      
     If (TypeOf cObjeto Is CheckBox) Then
       cObjeto.Enabled = wStatus
     End If
      
     If (TypeOf cObjeto Is ComboBox) Then
       cObjeto.Enabled = wStatus
     End If
      
     If (TypeOf cObjeto Is MaskEdBox) Then
       cObjeto.Enabled = wStatus
     End If
  Next
     
  txtLoja.Enabled = True
  txtVendedor.Enabled = True
  txtNFicha.Enabled = False
  txtDataIni.Enabled = False
  txtDataFim.Enabled = False
  txtCompras.Enabled = False
  txtDataUltimaCompra.Enabled = False
  txtValorUltimaCompra.Enabled = False
  txtDataMaiorCompra.Enabled = False
  txtValorMaiorCompra.Enabled = False
  txtCodCliente.Enabled = False
  txtProduto(0).Enabled = False
  txtProduto(1).Enabled = False
  txtProduto(2).Enabled = False
  txtData.Enabled = False
End Sub

