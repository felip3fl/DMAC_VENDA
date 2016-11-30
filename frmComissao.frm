VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmComissao 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Comissão Vendedor"
   ClientHeight    =   6375
   ClientLeft      =   5100
   ClientTop       =   3705
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPesquisaCliente 
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
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "[Insert] para inserir um novo cliente  "
      Top             =   840
      Width           =   6180
   End
   Begin VB.TextBox txtSenhaComissao1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   31
      Top             =   11160
      Width           =   1200
   End
   Begin VB.TextBox txtSenhaComissao 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   22
      Top             =   10560
      Width           =   1200
   End
   Begin VB.Frame FrmOpcao 
      BackColor       =   &H00505050&
      Caption         =   "Opção/Senha"
      ForeColor       =   &H00FFFFFF&
      Height          =   2730
      Left            =   1095
      TabIndex        =   19
      Top             =   1275
      Width           =   4620
      Begin VB.OptionButton chkReativacao 
         BackColor       =   &H00505050&
         Caption         =   "Reativação de Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   42
         Top             =   1800
         Width           =   4305
      End
      Begin VB.TextBox txtSenhaComissao2 
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
         IMEMode         =   3  'DISABLE
         Left            =   2265
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2205
      End
      Begin VB.OptionButton chkAlteraLojaVenda 
         BackColor       =   &H00505050&
         Caption         =   "Alterar Loja Venda"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   40
         Top             =   2200
         Width           =   4305
      End
      Begin VB.OptionButton chkPorCliente 
         BackColor       =   &H00505050&
         Caption         =   "Por Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   21
         Top             =   1400
         Width           =   4305
      End
      Begin VB.OptionButton chkPorNotaFiscal 
         BackColor       =   &H00505050&
         Caption         =   "Comissão"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Top             =   1000
         Width           =   4305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informe sua senha:"
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
         Height          =   240
         Left            =   150
         TabIndex        =   41
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.Frame frmNf 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   615
      Left            =   15
      TabIndex        =   10
      Top             =   5565
      Visible         =   0   'False
      Width           =   6420
      Begin VB.Label lblValor 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   2235
         TabIndex        =   18
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
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
         Left            =   1560
         TabIndex        =   17
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   1560
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblSerie 
         BackStyle       =   0  'Transparent
         Caption         =   "DDD"
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
         Left            =   795
         TabIndex        =   14
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Serie:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblNf 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   645
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N.F.:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   15
      ScaleHeight     =   45
      ScaleWidth      =   6360
      TabIndex        =   1
      Top             =   5640
      Width           =   6360
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensNF 
      Height          =   1515
      Left            =   6600
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   6180
      _cx             =   10901
      _cy             =   2672
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmComissao.frx":0000
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdNotaFiscal 
      Height          =   1335
      Left            =   6645
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   6180
      _cx             =   10901
      _cy             =   2355
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmComissao.frx":0085
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
   Begin MSMask.MaskEdBox mskDataInicial 
      Height          =   315
      Left            =   1080
      TabIndex        =   23
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataFinal 
      Height          =   315
      Left            =   2640
      TabIndex        =   24
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   0
      HideSelection   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdComissao 
      Height          =   3225
      Left            =   120
      TabIndex        =   25
      Top             =   1320
      Width           =   6180
      _cx             =   10901
      _cy             =   5689
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmComissao.frx":0102
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdNotaCliente 
      Height          =   3105
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   6180
      _cx             =   10901
      _cy             =   5477
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmComissao.frx":019A
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensCliente 
      Height          =   3195
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   6180
      _cx             =   10901
      _cy             =   5636
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmComissao.frx":0243
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
   Begin MSMask.MaskEdBox mskDataInicial1 
      Height          =   315
      Left            =   1080
      TabIndex        =   32
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataFinal1 
      Height          =   315
      Left            =   2655
      TabIndex        =   33
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   0
      HideSelection   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdCliente1 
      Height          =   3135
      Left            =   120
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   6180
      _cx             =   10901
      _cy             =   5530
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmComissao.frx":02C8
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome / Codigo / CPF / CGC"
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
      Left            =   120
      TabIndex        =   38
      Top             =   600
      Width           =   6165
   End
   Begin VB.Label LabelSenha1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Height          =   240
      Left            =   360
      TabIndex        =   36
      Top             =   11280
      Width           =   735
   End
   Begin VB.Label lblA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "à"
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
      Height          =   240
      Left            =   2400
      TabIndex        =   35
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblPeriodo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
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
      Height          =   240
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   825
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Height          =   240
      Left            =   360
      TabIndex        =   28
      Top             =   10680
      Width           =   735
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "à"
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
      Height          =   240
      Left            =   2400
      TabIndex        =   27
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
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
      Height          =   240
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   825
   End
   Begin VB.Label lblTotalDevolucao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Devolução:"
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
      Height          =   240
      Left            =   150
      TabIndex        =   7
      Top             =   5190
      Width           =   1785
   End
   Begin VB.Label lblValortotalDevolucao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
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
      Height          =   240
      Left            =   2070
      TabIndex        =   6
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Label lblTotalComissao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Comissão:"
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
      Height          =   240
      Left            =   3570
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblValorTotalComissao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5370
      TabIndex        =   4
      Top             =   4920
      Width           =   960
   End
   Begin VB.Label lblTotalVendas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vendas:"
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
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   4920
      Width           =   1470
   End
   Begin VB.Label lblValorTotalVendas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
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
      Height          =   240
      Left            =   2070
      TabIndex        =   2
      Top             =   4920
      Width           =   1170
   End
End
Attribute VB_Name = "frmComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wVendedor As String
Dim GLB_Senha As String
Dim rsComissao As New ADODB.Recordset
Dim rsComissaoCapa As New ADODB.Recordset

Dim SQL As String
Dim wTotalVenda As Double
Dim wTotalDevolucao As Double
Dim wTotalComissaoVenda As Double
Dim wLinhagrd As Double
Dim wNomeCliente As String

Dim rsCodigoCliente As New ADODB.Recordset
Dim rsNotaCliente As New ADODB.Recordset
Dim notafisc As String

Dim rsLabelValorNota As New ADODB.Recordset
Dim resultadoNota As String
Dim rsNomeCliente As New ADODB.Recordset
Dim NomeUsuario As String

Private Sub chkAlteraLojaVenda_Click()
    
    frmAlteraLojaVenda.Show 1
    frmAlteraLojaVenda.ZOrder
    Unload Me
    
End Sub

Private Sub chkReativacao_Click()
            'ricardo
            frmReativacaoCliente.Show 1
            frmReativacaoCliente.ZOrder
            Unload Me
        
End Sub

Private Sub chkReativacao_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 27 Then
        FrmOpcao.Visible = True
        txtSenhaComissao2.Visible = True
        Label6.Visible = True
        chkPorNotaFiscal.enable = False
        chkPorNotaFiscal.enable = False
        chkReativacao.enable = False
        chkAlteraLojaVenda.Enabled = False
        Label6.Enabled = True
        txtSenhaComissao2.Enabled = True
     End If

End Sub
Private Sub chkPorCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        FrmOpcao.Visible = True
        txtSenhaComissao2.Visible = True
        Label6.Visible = True
        chkPorNotaFiscal.enable = False
        chkPorNotaFiscal.enable = False
        chkReativacao.enable = False
        chkAlteraLojaVenda.Enabled = False
        Label6.Enabled = True
        txtSenhaComissao2.Enabled = True
    End If
End Sub
Private Sub chkPorNotaFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        FrmOpcao.Visible = True
        txtSenhaComissao2.Visible = True
        Label6.Visible = True
        chkPorNotaFiscal.enable = False
        chkPorNotaFiscal.enable = False
        chkReativacao.enable = False
        chkAlteraLojaVenda.Enabled = False
        Label6.Enabled = True
        txtSenhaComissao2.Enabled = True
    End If
End Sub

Private Sub carregaSenhaVendedor()

    Dim rsSenha As New ADODB.Recordset

   wVendedor = Mid(frmPedido.txtVendedor.Text, 1, 3)
  
   SQL = "select ve_Nome,ve_senha from vende where ve_codigo = '" & wVendedor & "'"
   rsSenha.CursorLocation = adUseClient
   rsSenha.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   
   'ricardo
   GLB_Senha = Trim(rsSenha("ve_Senha"))
   
   rsSenha.Close

End Sub

Private Sub Form_Activate()
    carregaSenhaVendedor
End Sub

Private Sub Form_Load()

    chkPorNotaFiscal.Enabled = False
    chkPorCliente.Enabled = False
    chkReativacao.Enabled = False
    chkAlteraLojaVenda.Enabled = False

   frmNf.top = lblTotalVendas.top
   Call AjustaTela(frmComissao)
   wVendedor = ""
   GLB_Senha = ""
   grdComissao.Row = -1
   grdNotaFiscal.top = grdComissao.top
   grdNotaFiscal.left = grdComissao.left
   grdNotaFiscal.Height = grdComissao.Height
   
   grdItensNF.top = grdComissao.top
   grdItensNF.left = grdComissao.left
   grdItensNF.Height = grdComissao.Height
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''
'    grdNotaCliente.top = grdCliente1.top
'    grdNotaCliente.left = grdCliente1.left
'    grdNotaCliente.Height = grdCliente1.Height
'
'
'    grdItensCliente.top = grdNotaCliente.top
'    grdItensCliente.left = grdNotaCliente.left
'    grdItensCliente.Height = grdNotaCliente.Height
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''
 


   
   mskDataInicial.Text = "__/__/____"
   mskDataFinal.Text = "__/__/____"
   mskDataInicial.Enabled = True
   mskDataFinal.Enabled = True
   
  wTotalVenda = 0
  wTotalDevolucao = 0
  wTotalComissaoVenda = 0
  lblValorTotalComissao.Caption = "0,00"
  lblValorTotalVendas.Caption = "0,00"
  lblValortotalDevolucao.Caption = "0,00"
  grdComissao.Row = -1
   
   
    FrmOpcao.Visible = True
    frmComissao.FrmOpcao.Visible = True
    
    lblSenha.Visible = False
    txtSenhaComissao.Visible = False
    lblPeriodo.Visible = False
    mskDataInicial.Visible = False
    lblA.Visible = False
    mskDataFinal.Visible = False
    Label4.Visible = False
    txtPesquisaCliente.Visible = False
    grdComissao.Visible = False
    grdNotaCliente.Visible = False
    grdItensCliente.Visible = False
    lblTotalVendas.Visible = False
    lblValorTotalVendas.Visible = False
    lblTotalComissao.Visible = False
    lblValorTotalComissao.Visible = False
    lblTotalDevolucao.Visible = False
    lblValortotalDevolucao.Visible = False
    Label1.Visible = False
    lblNf.Visible = False
    Label3.Visible = False
    lblCliente.Visible = False
    Label2.Visible = False
    lblSerie.Visible = False
    Label5.Visible = False
    lblValor.Visible = False
    'Frame1.Visible = False
    LabelSenha1.Visible = False
    txtSenhaComissao1.Visible = False
    lblPeriodo1.Visible = False
    mskDataInicial1.Visible = False
    lblA1.Visible = False
    mskDataFinal1.Visible = False
    
    chkPorNotaFiscal.Enabled = False
    chkPorCliente.Enabled = False
    chkReativacao.Enabled = False
    chkAlteraLojaVenda.Enabled = False
            
            
End Sub

Private Sub FrmOpcao_DblClick()
    frmAlteraLojaVenda.Show 1
End Sub

Private Sub grdCliente1_Click()
    txtPesquisaCliente.Text = grdCliente1.TextMatrix(grdCliente1.Row, 1)
End Sub

Private Sub grdComissao_DblClick()
wLinhagrd = grdComissao.Row
  grdComissao.Visible = False
  If wLinhagrd > 0 Then
    grdNotaFiscal.Visible = True
     grdItensNF.Visible = False
     grdNotaFiscal.SetFocus
     CarregagrdNotaFiscal
Else
grdComissao.Visible = True
End If
End Sub

Private Sub grdComissao_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
        frmPedido.txtPesquisar.SetFocus
    If rdoCNMatriz <> "" Then
     rdoCNMatriz.Close
     End If
End If
End Sub

Private Sub grdComissao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
wLinhagrd = grdComissao.Row
  grdComissao.Visible = False
  If wLinhagrd <> 0 Then
    grdNotaFiscal.Visible = True
     grdItensNF.Visible = False
     grdNotaFiscal.SetFocus
    CarregagrdNotaFiscal
Else
grdComissao.Visible = True
End If

End If

End Sub

Private Sub grdItensNF_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    grdItensNF.Visible = False
 grdNotaFiscal.Visible = True
 grdNotaFiscal.SetFocus
 frmNf.Visible = False
End If
End Sub



Private Sub grdNotaFiscal_DblClick()
        wLinhagrd = grdNotaFiscal.Row
          grdNotaFiscal.Visible = False
          If wLinhagrd > 0 Then
            grdItensNF.Visible = True
            grdItensNF.SetFocus
            frmNf.Visible = True
            grdItensNF.SetFocus
            CarregagrditensNf
        Else
        grdComissao.Visible = True
        End If


End Sub

Private Sub grdNotaFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
   grdNotaFiscal.Visible = False
   grdComissao.Visible = True
   grdComissao.SetFocus
   
End If
End Sub

Private Sub grdNotaFiscal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
wLinhagrd = grdNotaFiscal.Row
  grdNotaFiscal.Visible = False
  If wLinhagrd <> 0 Then
    grdItensNF.Visible = True
    frmNf.Visible = True
    grdItensNF.SetFocus
    CarregagrditensNf
Else
grdNotaFiscal.Visible = True
End If

End If


End Sub

Private Sub mskDataFinal_GotFocus()
          mskDataFinal.SelStart = 0
          mskDataFinal.SelLength = Len(mskDataFinal.Text)
          mskDataFinal.SetFocus
          wTotalVenda = 0
          wTotalDevolucao = 0
          wTotalComissaoVenda = 0
          lblValorTotalComissao.Caption = "0,00"
          lblValorTotalVendas.Caption = "0,00"
          lblValortotalDevolucao.Caption = "0,00"
          grdComissao.Rows = 1
End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
        frmPedido.txtPesquisar.SetFocus
    If rdoCNMatriz <> "" Then
     rdoCNMatriz.Close
     End If
End If
End Sub

Private Sub mskDataFinal_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Not IsDate(mskDataInicial.Text) Then
       MsgBox "Data inicial Inválida", vbCritical, "Atenção"
       mskDataInicial.SelLength = Len(mskDataInicial.Text)
       mskDataInicial.SelStart = 0
       mskDataInicial.SetFocus
       Exit Sub
   End If

   If Not IsDate(mskDataFinal.Text) Then
       MsgBox "Data Final Inválida", vbCritical, "Atenção"
       mskDataFinal.SelLength = Len(mskDataFinal.Text)
       mskDataFinal.SelStart = 0
       mskDataFinal.SetFocus
       Exit Sub
   End If

   If Format(mskDataInicial.Text, "yyyy/mm/dd") > Format(mskDataFinal.Text, "yyyy/mm/dd") Then
       MsgBox "Data Inicial não pode ser maior que a Data Final", vbCritical, "Atenção"
       mskDataInicial.SelLength = Len(mskDataInicial.Text)
       mskDataInicial.SelStart = 0
       mskDataInicial.SetFocus
       Exit Sub
   End If
   
  Screen.MousePointer = 11

   ConectaODBCMatriz
   If GLB_ConectouOK = False Then
       MsgBox "Erro ao conectar-se ao Banco de Dados da Matriz", vbCritical, "Atenção"
       Exit Sub
   End If

'ricardo
grdNotaCliente.Visible = False
grdComissao.Visible = True


    SQL = "select vi_dataemissao,sum(vi_valormercadoria) as vi_valormercadoria, " & _
          "sum(vi_valorComissao) as vi_valorComissao ,sum(vi_vendacomissao) as vi_vendacomissao " & _
          "From itemnfvenda, capanfvenda " & _
          "where vi_lojaorigem = vc_lojaorigem and vi_notafiscal = vc_notafiscal and vi_serie = vc_serie and " & _
          "vc_vendedorlojavenda = '" & wVendedor & "' and vi_tiponota = 'v' and " & _
          "vi_dataemissao between '" & Format(mskDataInicial.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal.Text, "yyyy/mm/dd") & "' " & _
          "group by vi_dataemissao"
   rsComissao.CursorLocation = adUseClient
   rsComissao.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
   

   If rsComissao.EOF = False Then
      Do While rsComissao.EOF = False
           
           SQL = "select sum(vc_totalnota) as vc_totalnota From capanfvenda " & _
                 "where vc_tiponota = 'v' and vc_dataemissao = '" & Format(rsComissao("vi_dataemissao"), "yyyy/mm/dd") & "' and " & _
                 "vc_vendedorlojavenda = '" & wVendedor & "' "
           rsComissaoCapa.CursorLocation = adUseClient
           rsComissaoCapa.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
           
           wTotalVenda = wTotalVenda + rsComissaoCapa("vc_totalnota")
           wTotalComissaoVenda = wTotalComissaoVenda + rsComissao("vi_valorComissao")


           If rsComissaoCapa.EOF = False Then
                  grdComissao.AddItem Format(rsComissao("vi_dataemissao"), "dd/mm/yyyy") & Chr(9) & _
                                      Format(rsComissaoCapa("vc_totalnota"), "###,###,###,##0.00") & Chr(9) & _
                                      Format(rsComissao("vi_vendacomissao"), "###,###,###,##0.00") & Chr(9) & _
                                      Format(rsComissao("vi_valorComissao"), "###,###,###,##0.00")
           Else
                  MsgBox "Erro na Capa da Comissão. Informe o Departamento de TI."
           End If
           
           rsComissao.MoveNext
           rsComissaoCapa.Close
       Loop
    Else
       MsgBox "Não há comissão nesse período"
    End If


    SQL = "select sum(vc_totalnota) as vc_totalnota From capanfvenda " & _
          "where vc_tiponota = 'E' and vc_vendedorlojavenda = '" & wVendedor & "' and " & _
          "vc_dataemissao between '" & Format(mskDataInicial.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal.Text, "yyyy/mm/dd") & "' "
    rsComissaoCapa.CursorLocation = adUseClient
    rsComissaoCapa.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic

    
    lblValortotalDevolucao.Caption = IIf(IsNull(rsComissaoCapa("vc_totalnota")), "0,00", Format(rsComissaoCapa("vc_totalnota"), "###,###,###,##0.00"))
    rsComissaoCapa.Close
    
    SQL = "select sum(vi_valorComissao) as vi_valorComissao " & _
          "From itemnfvenda, capanfvenda " & _
          "where vi_lojaorigem = vc_lojaorigem and vi_notafiscal = vc_notafiscal and vi_serie = vc_serie and " & _
          "vc_vendedorlojavenda = '" & wVendedor & "' and vi_tiponota = 'E' and " & _
          "vi_dataemissao between '" & Format(mskDataInicial.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal.Text, "yyyy/mm/dd") & "' "
          rsComissaoCapa.CursorLocation = adUseClient
          rsComissaoCapa.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    lblValorTotalComissao.Caption = Format(wTotalComissaoVenda - IIf(IsNull(rsComissaoCapa("vi_ValorComissao")), 0, rsComissaoCapa("vi_ValorComissao")), "###,###,###,##0.00")
    lblValorTotalVendas.Caption = Format(wTotalVenda, "###,###,###,##0.00")
    
    
    rsComissaoCapa.Close
    rsComissao.Close
    Screen.MousePointer = 0
    
End If
   
End Sub

Private Sub mskDataInicial_GotFocus()
          mskDataInicial.SelStart = 0
          mskDataInicial.SelLength = Len(mskDataInicial.Text)
          mskDataInicial.SetFocus
          wTotalVenda = 0
          wTotalDevolucao = 0
          wTotalComissaoVenda = 0
          lblValorTotalComissao.Caption = "0,00"
          lblValorTotalVendas.Caption = "0,00"
          lblValortotalDevolucao.Caption = "0,00"
          grdComissao.Rows = 1
End Sub

Private Sub mskDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
        frmComissao.FrmOpcao.Visible = True
        'frmPedido.txtPesquisar.SetFocus
    If rdoCNMatriz <> "" Then
     rdoCNMatriz.Close
     End If
End If
End Sub
Private Sub mskDataInicial1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
        frmPedido.txtPesquisar.SetFocus
        frmComissao.FrmOpcao.Visible = True
    If rdoCNMatriz <> "" Then
     rdoCNMatriz.Close
     End If
End If
End Sub

Private Sub mskDataInicial_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     mskDataFinal.SetFocus
   End If
End Sub

Private Sub mskDataFinal1_KeyPress(KeyAscii As Integer)
'ricardo

End Sub

Private Sub txtPesquisaCliente_Click()
    'grdComissao.Rows = 1
End Sub

Private Sub txtSenhaComissao_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
        frmPedido.txtPesquisar.SetFocus
    If rdoCNMatriz <> "" Then
     rdoCNMatriz.Close
     End If
End If
End Sub

Private Sub txtSenhaComissao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       If txtSenhaComissao.Text = "" Then
       
          MsgBox "Informar senha do vendedor", vbExclamation, Me.Caption
          
       ElseIf txtSenhaComissao.Text = "'" Or Trim(txtSenhaComissao.Text) <> GLB_Senha Then
       
          MsgBox "Senha incorreta", vbExclamation, Me.Caption
          txtSenhaComissao.SelStart = 0
          txtSenhaComissao.SelLength = Len(txtSenhaComissao.Text)
          txtSenhaComissao.SetFocus
          
       Else
          mskDataInicial.Enabled = True
          mskDataFinal.Enabled = True
          mskDataInicial.SetFocus
       End If
Else
    txtSenhaComissao.SetFocus
       
    End If

End Sub

Private Sub CarregagrdNotaFiscal()
wNomeCliente = ""
grdNotaFiscal.Rows = 1
    SQL = "select VC_NotaFiscal,VC_Serie,VC_TotalNota, VC_Cliente,VC_NomeCliente " & _
          "From  capanfvenda " & _
          "where vc_vendedorlojavenda= '" & wVendedor & "' and  vc_dataemissao = '" & Format(grdComissao.TextMatrix(wLinhagrd, 0), "yyyy/mm/dd") & "'" & _
          " and vc_tiponota = 'v' order by VC_NotaFiscal"
   rsComissaoCapa.CursorLocation = adUseClient
   rsComissaoCapa.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
       Do While rsComissaoCapa.EOF = False
        
        '------- Cliente Loja-------
        sql1 = "select Ce_Razao  from  fin_cliente where ce_codigocliente=" & rsComissaoCapa("VC_Cliente")
        rsComissao.CursorLocation = adUseClient
        rsComissao.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        If rsComissao.EOF = False Then
       wNomeCliente = rsComissao("Ce_Razao")
       Else
       wNomeCliente = ""
       End If
       rsComissao.Close
        '-------------
        
        
        grdNotaFiscal.AddItem rsComissaoCapa("VC_NotaFiscal") & Chr(9) & _
                              rsComissaoCapa("VC_Serie") & Chr(9) & _
                             Format(rsComissaoCapa("VC_TotalNota"), "###,###,###,##0.00") & Chr(9) & _
                              rsComissaoCapa("VC_Cliente") & " - " & wNomeCliente
       rsComissaoCapa.MoveNext
       Loop
       
   rsComissaoCapa.Close

End Sub
Private Sub CarregagrditensNf()
grdItensNF.Rows = 1
    SQL = "select VI_Referencia,VI_Quantidade,VI_ValorMercadoria,PR_Descricao from  itemnfvenda,Produto " & _
          " where VI_NotaFiscal=" & grdNotaFiscal.TextMatrix(wLinhagrd, 0) & "   and  VI_Serie ='" & grdNotaFiscal.TextMatrix(wLinhagrd, 1) & "'" & _
          "  and PR_Referencia=VI_Referencia and vi_tiponota = 'V' and vi_lojaorigem='" & wLoja & "' order by vi_referencia"
   rsComissaoCapa.CursorLocation = adUseClient
   rsComissaoCapa.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
       Do While rsComissaoCapa.EOF = False
         grdItensNF.AddItem rsComissaoCapa("VI_Referencia") & Chr(9) & _
                              rsComissaoCapa("VI_Quantidade") & Chr(9) & _
                              Format(rsComissaoCapa("VI_ValorMercadoria"), "###,###,###,##0.00") & Chr(9) & _
                              rsComissaoCapa("PR_Descricao")
       rsComissaoCapa.MoveNext
       Loop
       
   rsComissaoCapa.Close
   'grdItensNF.RowSel = 1
   
   
   '---------FrmdadosNota
   
   lblCliente.Caption = grdNotaFiscal.TextMatrix(wLinhagrd, 3)
   lblNf.Caption = grdNotaFiscal.TextMatrix(wLinhagrd, 0)
   lblSerie.Caption = grdNotaFiscal.TextMatrix(wLinhagrd, 1)
   lblValor.Caption = grdNotaFiscal.TextMatrix(wLinhagrd, 2)


End Sub

'#####################################################################################################
'#####################################################################################################

''Function PesquisaCliente(ByVal tipoPesquisa As Integer, ByVal Cliente As String) As Boolean
''
''    '-------------------------------Pesquisa Pelo Codigo Cliente (1) ---------------------------------
''
''
''    If tipoPesquisa = 1 Then
''
''        SQL = ""
''        SQL = "select NF.cliente as Codigo_Cliente, NF.Nomcli as Nome_do_Cliente " & _
''        "from nfcapa as NF, fin_cliente as FIN " & _
''        "where NF.vendedor = '" & Mid(frmPedido.txtVendedor, 1, 2) & "' and DataEmi between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' and NF.Cliente <> '999999' and NF.cliente = FIN.ce_codigoCliente " & _
''        "and NF.nf <> 0  and FIN.ce_codigoCliente =  '" & txtPesquisaCliente.Text & "' "
''
''        rsCodigoCliente.CursorLocation = adUseClient
''        rsCodigoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''
''
''     '-------------------------------Pesquisa Pelo Nome Cliente (2) ---------------------------------
''
''    ElseIf tipoPesquisa = 2 Then
''        SQL = ""
''        SQL = "select NF.cliente as Codigo_Cliente, NF.Nomcli as Nome_do_Cliente " & _
''        "from nfcapa as NF, fin_cliente as FIN " & _
''        "where NF.vendedor = '" & Mid(frmPedido.txtVendedor.Text, 1, 2) & "' " & _
''        "and DataEmi between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and " & _
''        "'" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' and NF.Cliente <> '999999' and NF.cliente = FIN.ce_codigoCliente " & _
''        "and NF.nf <> 0 and NF.NomeCli like '" & txtPesquisaCliente.Text & "%' order by NF.nf "
''
''        rsNomeCliente.CursorLocation = adUseClient
''        rsNomeCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''
''
''    Else
''        Exit Function
''    End If
''
''End Function

Private Sub txtPesquisaCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        If Trim(txtPesquisaCliente.Text) <> "" Then
            txtPesquisaCliente.Text = UCase(txtPesquisaCliente.Text)
            If IsNumeric(txtPesquisaCliente.Text) = True Then
                If Len(txtPesquisaCliente.Text) < 11 Then
                    '************Pesquisa Pelo codigo do Cliente (1)**********************
                    If PesquisaCliente(1, txtPesquisaCliente.Text) = True Then
                    grdCliente1.Rows = 1
                        Do While Not rsCodigoCliente.EOF
                          grdCliente1.AddItem rsCodigoCliente("Codigo_Cliente") & Chr(9) & rsCodigoCliente("Nome_do_Cliente")
                          rsCodigoCliente.MoveNext
                        Loop
                        rsCodigoCliente.Close
                        
                         
                    End If
                    
            
                ElseIf Len(txtPesquisaCliente.Text) >= 11 Then
                    '***************Pesquisa por CGC ou Cpf (2)***************************
                    If PesquisaCliente(2, txtPesquisaCliente.Text) = True Then
                        grdCliente1.Rows = 1
                        Do While Not rsCodigoCliente.EOF
                           grdCliente1.AddItem rsCodigoCliente("Codigo_Cliente") & Chr(9) & rsCodigoCliente("Nome_do_Cliente")
                           rsCodigoCliente.MoveNext
                        Loop
                        rsCodigoCliente.Close
                       

                    End If
                Else
                End If
                
            ElseIf IsNumeric(txtPesquisaCliente.Text) = False Then
                '************Pesquisa Pelo Nome Cliente (3)******************************
                If PesquisaCliente(3, txtPesquisaCliente.Text) = True Then
                    grdCliente1.Rows = 1
                    Do While Not rsCodigoCliente.EOF
                     grdCliente1.AddItem rsCodigoCliente("Codigo_Cliente") & Chr(9) & rsCodigoCliente("Nome_do_Cliente")
                     rsCodigoCliente.MoveNext
                    Loop
                    rsCodigoCliente.Close
                     
                    
               End If
               End If
               End If
               End If
             
               
End Sub

Function PesquisaCliente(ByVal tipoPesquisa As Integer, ByVal Cliente As String) As Boolean

'
'--------------------------------Pesquisa Pelo Codigo do Cliente (1)-------------------------

    If tipoPesquisa = 1 Then
        SQL = ""
        SQL = "select NF.cliente as Codigo_Cliente, NF.Nomcli as Nome_do_Cliente " & _
        "from nfcapa as NF, fin_cliente as FIN " & _
        "where NF.vendedor = '" & Mid(frmPedido.txtVendedor, 1, 2) & "' and DataEmi between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' and NF.Cliente <> '999999' and NF.cliente = FIN.ce_codigoCliente " & _
        "and NF.nf <> 0  and FIN.ce_codigoCliente =  '" & txtPesquisaCliente.Text & "' "

        rsCodigoCliente.CursorLocation = adUseClient
        rsCodigoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
''
'''
'''-------------------------------Pesquisa por cgc ou cpf (2) ---------------------------------
'''
    ElseIf tipoPesquisa = 2 Then
        SQL = ""
        SQL = ""
        
        SQL = "SELECT CAP.VC_CLIENTE as Codigo_Cliente, NF.Nomcli as Nome_do_Cliente  FROM CAPANFVENDA as CAP, FIN_CLIENTE as FIN, NFCAPA as NF " & _
        "Where VC_CLIENTE = Nf.Cliente " & _
        "AND Vc_VendedorLojaVenda = '" & Mid(frmPedido.txtVendedor.Text, 1, 2) & "' " & _
        "AND CAP.VC_DATAEMISSAO BETWEEN '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' " & _
        "AND NF.CLIENTE <> '999999' " & _
        "AND CAP.VC_CLIENTE = FIN.CE_CODIGOCLIENTE AND CAP.VC_NOTAFISCAL <> 0 " & _
        "AND FIN.CE_CGC = '" & txtPesquisaCliente.Text & "'"


        ConectaODBCMatriz
        rsCodigoCliente.CursorLocation = adUseClient
        rsCodigoCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic


'''
'''-------------------------------Pesquisa Pelo Nome Cliente (3) ---------------------------------
'''
    ElseIf tipoPesquisa = 3 Then
        SQL = ""
        SQL = "select NF.cliente as Codigo_Cliente, NF.Nomcli as Nome_do_Cliente " & _
        "from nfcapa as NF, fin_cliente as FIN " & _
        "where NF.vendedor = '" & Mid(frmPedido.txtVendedor.Text, 1, 3) & "' " & _
        "and DataEmi between '" & Format(mskDataInicial.Text, "yyyy/mm/dd") & "' and " & _
        "'" & Format(mskDataFinal.Text, "yyyy/mm/dd") & "' and NF.Cliente <> '999999' and NF.cliente = FIN.ce_codigoCliente " & _
        "and NF.nf <> 0 and NF.NomCli like '" & txtPesquisaCliente.Text & "%' order by NF.nf "
        
        rsCodigoCliente.CursorLocation = adUseClient
        rsCodigoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    Else
        Exit Function
    End If
    If Not rsCodigoCliente.EOF Then
        PesquisaCliente = True
    Else
        PesquisaCliente = False
        rsCodigoCliente.Close
    End If
    DescricaoOperacao "Pronto"

End Function

Private Sub chkPorNotaFiscal_KeyPress(KeyAscii As Integer)
'ricardo
    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub
Private Sub chkPorNotaFiscal_Click()
'ricardo
'MsgBox "ok"
            
            'frmComissao.LabelSenha1.Visible = False
            'frmComissao.txtSenhaComissao1.Visible = False
            frmComissao.lblPeriodo1.Visible = True
            frmComissao.mskDataInicial1.Visible = True
            frmComissao.lblA1.Visible = True
            frmComissao.mskDataFinal1.Visible = True
            frmComissao.Label4.Visible = False
            frmComissao.txtPesquisaCliente.Visible = False
            frmComissao.grdCliente1.Visible = False
            frmComissao.lblTotalVendas.Visible = False
            frmComissao.lblValorTotalVendas.Visible = False
            
            'frmComissao.lblSenha.Visible = True
            'frmComissao.txtSenhaComissao.Visible = True
            frmComissao.lblPeriodo.Visible = True
            frmComissao.mskDataInicial.Visible = True
            frmComissao.lblA.Visible = True
            frmComissao.mskDataFinal.Visible = True
            frmComissao.Label4.Visible = False
            frmComissao.txtPesquisaCliente.Visible = False
            frmComissao.grdComissao.Visible = True
            frmComissao.lblTotalVendas.Visible = True
            frmComissao.lblValorTotalVendas.Visible = True
            frmComissao.lblTotalComissao.Visible = True
            frmComissao.lblValorTotalComissao.Visible = True
            frmComissao.lblTotalDevolucao.Visible = True
            frmComissao.lblValortotalDevolucao.Visible = True
            frmComissao.Label1.Visible = True
            frmComissao.lblNf.Visible = True
            frmComissao.Label3.Visible = True
            frmComissao.lblCliente.Visible = True
            frmComissao.Label2.Visible = True
            frmComissao.lblSerie.Visible = True
            frmComissao.Label5.Visible = True
            frmComissao.lblValor.Visible = True
            
            frmComissao.FrmOpcao.Visible = False
            'frmComissao.txtSenhaComissao.SetFocus
            
            grdComissao.Visible = False
            grdNotaCliente.Visible = True
            Label4.Visible = True
            txtPesquisaCliente.Visible = True

End Sub
Private Sub chkPorCliente_Click()
'ricardo
            frmComissao.FrmOpcao.Visible = False
            
            
            frmComissao.lblSenha.Visible = False
            frmComissao.txtSenhaComissao.Visible = False
            frmComissao.lblPeriodo.Visible = True
            frmComissao.mskDataInicial.Visible = True
            frmComissao.lblA.Visible = True
            frmComissao.mskDataFinal.Visible = True
            frmComissao.Label4.Visible = False
            frmComissao.txtPesquisaCliente.Visible = False
            frmComissao.grdComissao.Visible = False
            frmComissao.lblTotalVendas.Visible = False
            frmComissao.lblValorTotalVendas.Visible = False
            frmComissao.lblTotalComissao.Visible = False
            frmComissao.lblValorTotalComissao.Visible = False
            frmComissao.lblTotalDevolucao.Visible = False
            frmComissao.lblValortotalDevolucao.Visible = False
            frmComissao.Label1.Visible = False
            frmComissao.lblNf.Visible = False
            frmComissao.Label3.Visible = False
            frmComissao.lblCliente.Visible = False
            frmComissao.Label2.Visible = False
            frmComissao.lblSerie.Visible = False
            frmComissao.Label5.Visible = False
            frmComissao.lblValor.Visible = False
            frmComissao.grdNotaCliente.Visible = False

           
            frmComissao.LabelSenha1.Visible = True
            frmComissao.txtSenhaComissao1.Visible = True
            frmComissao.lblPeriodo1.Visible = True
            frmComissao.mskDataInicial1.Visible = True
            frmComissao.lblA1.Visible = True
            frmComissao.mskDataFinal1.Visible = True
            frmComissao.Label4.Visible = True
            frmComissao.txtPesquisaCliente.Visible = True
            frmComissao.grdCliente1.Visible = True
            frmComissao.lblTotalVendas.Visible = True
            frmComissao.lblValorTotalVendas.Visible = True

            
End Sub

Private Sub txtSenhaComissao1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtSenhaComissao2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtSenhaComissao2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       
       If frmComissao.txtSenhaComissao2.Text = "" Then
          MsgBox "Informar senha do vendedor", vbExclamation, Me.Caption
       ElseIf frmComissao.txtSenhaComissao2.Text = "'" Or Trim(frmComissao.txtSenhaComissao2.Text) <> GLB_Senha Then
       
          MsgBox "Senha incorreta", vbExclamation, Me.Caption
          frmComissao.txtSenhaComissao2.SelStart = 0
          frmComissao.txtSenhaComissao2.SelLength = Len(frmComissao.txtSenhaComissao2.Text)
          frmComissao.txtSenhaComissao2.SetFocus
          
       Else
          'ricardo
'            frmComissao.Label6.Visible = False
'            frmComissao.txtSenhaComissao2.Visible = False
            frmComissao.Label6.Enabled = False
            frmComissao.txtSenhaComissao2.Enabled = False
            chkPorNotaFiscal.Enabled = True
            chkPorCliente.Enabled = True
            chkReativacao.Enabled = True
            chkAlteraLojaVenda.Enabled = True
            frmComissao.txtSenhaComissao2.Text = ""
            
            
       End If
       
    End If

End Sub
Private Sub grdCliente1_DblClick()
Dim rsNotaCliente As New ADODB.Recordset
Dim notafisc As String
Dim rsLabelValorNota As New ADODB.Recordset
Dim resultadoNota As String

    
    grdCliente1.Visible = False
    grdNotaCliente.Visible = True
   
    
    
    '--------------------------------Pesquisa Pela Nota do Cliente -------------------------
    
    SQL = ""
    SQL = ""
    SQL = ""
    SQL = "select  vc_NotaFiscal,vc_serie,vc_dataemissao,vc_TotalNota, vc_Desconto from capanfvenda, itemnfvenda " & _
    "Where vc_NotaFiscal = VI_NotaFiscal And vc_serie = vi_serie And VI_LojaOrigem = VC_LojaOrigem " & _
    "and vc_DataEmissao between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' " & _
    "and VC_VendedorLojaVenda = '" & Mid(frmPedido.txtVendedor, 1, 2) & "' " & _
    "and vc_cliente = '" & grdCliente1.TextMatrix(grdCliente1.Row, 0) & "' " & _
    "group by vc_NotaFiscal,vc_serie,vc_dataemissao,vc_TotalNota, vc_Desconto "


    'CAP.VC_CLIENTE = '57437'
        
         
    ConectaODBCMatriz
    rsNotaCliente.CursorLocation = adUseClient
    rsNotaCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    Do While Not rsNotaCliente.EOF
        grdNotaCliente.AddItem rsNotaCliente("vc_NotaFiscal") & Chr(9) & _
                               rsNotaCliente("vc_serie") & Chr(9) & _
                               rsNotaCliente("vc_dataemissao") & Chr(9) & _
                               rsNotaCliente("vc_TotalNota") & Chr(9) & _
                               rsNotaCliente("vc_Desconto")
        rsNotaCliente.MoveNext
    Loop
    
    SomaNota

    rsNotaCliente.Close
    
    

End Sub

Public Sub SomaNota()
Dim rsLabelValorNota As New ADODB.Recordset
Dim resultado As String

    SQL = "select sum(vc_totalNota) as ValorNota from capanfvenda " & _
    "where vc_DataEmissao between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' " & _
    "and VC_VendedorLojaVenda = '" & Mid(frmPedido.txtVendedor.Text, 1, 2) & "' " & _
    "and vc_cliente = '" & grdCliente1.TextMatrix(grdCliente1.Row, 0) & "' "
    
    'ConectaODBCMatriz
    rsLabelValorNota.CursorLocation = adUseClient
    rsLabelValorNota.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
     resultado = rsLabelValorNota("ValorNota")
     lblValorTotalVendas.Caption = resultado
     
    rsLabelValorNota.Close

End Sub

Private Sub grdNotaCliente_DblClick()
Dim rsItensCliente As New ADODB.Recordset

    grdNotaCliente.Visible = False
    grdItensCliente.Visible = True
   
   
    
     '--------------------------------Pesquisa Pela Itens do Cliente -------------------------
    SQL = ""
    SQL = ""
''    SQL = "select  NFI.Referencia as Referencia, NFI.QTDE as Qtda, NFI.VLTOTITEM as Valor ,pr_Descricao " & _
''    "from nfitens as NFI, fin_cliente as FIN, ProdutoLoja " & _
''    "where NFI.nf = '" & grdNotaCliente.TextMatrix(grdNotaCliente.Row, 0) & "' and NFI.referencia = pr_referencia " & _
''    "and DataEmi between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "'" & _
''    "and NFI.vendedor = '" & Mid(frmPedido.txtVendedor, 1, 2) & "'" & _
''    "and NFI.cliente = '" & grdCliente1.TextMatrix(grdCliente1.Row, 0) & "' and NFI.cliente = FIN.ce_codigoCliente"

    SQL = "select vi_referencia,vi_quantidade,vi_valorMercadoria,pr_Descricao  from itemnfvenda , capanfvenda, produto " & _
    "Where VC_NotaFiscal = VI_NotaFiscal And vc_serie = vi_serie And VI_LojaOrigem = VC_LojaOrigem And pr_referencia = vi_referencia " & _
    "and vi_DataEmissao between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "'" & _
    "and Vc_VendedorLojaVenda = '" & Mid(frmPedido.txtVendedor, 1, 2) & "'" & _
    "and vc_cliente = '" & grdCliente1.TextMatrix(grdCliente1.Row, 0) & "' " & _
    "and vc_cliente = '" & grdCliente1.TextMatrix(grdCliente1.Row, 0) & "' order by pr_Descricao "
    
   

   'ConectaODBCMatriz
    rsItensCliente.CursorLocation = adUseClient
    rsItensCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic


    Do While Not rsItensCliente.EOF
        grdItensCliente.AddItem rsItensCliente("vi_referencia") & Chr(9) & _
                                rsItensCliente("vi_quantidade") & Chr(9) & _
                                rsItensCliente("vi_valorMercadoria") & Chr(9) & _
                                rsItensCliente("pr_Descricao")
                                
        rsItensCliente.MoveNext
    Loop
        SomaItens
        rsItensCliente.Close
        
    
End Sub

Public Sub SomaItens()

    Dim rsLabelValor As New ADODB.Recordset
    Dim resultado As String
    
    lblValorTotalVendas.Caption = ""
    
   SQL = ""
   SQL = "select sum(vi_valorMercadoria) as ValorNota from itemnfvenda,capanfvenda, produto " & _
   "Where VC_NotaFiscal = VI_NotaFiscal And vc_serie = vi_serie And VI_LojaOrigem = VC_LojaOrigem And pr_referencia = vi_referencia " & _
   "and vi_DataEmissao between '" & Format(mskDataInicial1.Text, "yyyy/mm/dd") & "' and '" & Format(mskDataFinal1.Text, "yyyy/mm/dd") & "' " & _
   "and Vc_VendedorLojaVenda = '" & Mid(frmPedido.txtVendedor.Text, 1, 2) & "'" & _
   "and vc_cliente = '" & grdCliente1.TextMatrix(grdCliente1.Row, 0) & "' "
        
    'ConectaODBCMatriz
    rsLabelValor.CursorLocation = adUseClient
    rsLabelValor.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
    
    resultado = rsLabelValor("ValorNota")
    lblValorTotalVendas.Caption = resultado
    
    rsLabelValor.Close

End Sub

Private Sub grdCliente1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        grdCliente1.Rows = 1
        txtSenhaComissao1.Text = ""
        mskDataInicial1.Text = "__/__/____"
        mskDataFinal1.Text = "__/__/____"
        txtPesquisaCliente.Text = ""
        txtSenhaComissao1.SetFocus
        lblValorTotalVendas.Caption = ""
        
    End If
    
End Sub

Private Sub grdNotaCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        grdNotaCliente.Rows = 1
        grdNotaCliente.Visible = False
        grdCliente1.Visible = True
        SomaItens
        
    End If

End Sub

Private Sub grdItensCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        grdItensCliente.Rows = 1
        grdItensCliente.Visible = False
        grdNotaCliente.Visible = True
        SomaNota
        
    End If
    
End Sub































