VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7u.ocx"
Begin VB.Form frmPedido 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "DMAC Venda"
   ClientHeight    =   10260
   ClientLeft      =   3510
   ClientTop       =   165
   ClientWidth     =   15120
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   Icon            =   "frmPedido.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPedido.frx":23FA
   Picture         =   "frmPedido.frx":2CC4
   ScaleHeight     =   10260
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox lblBloqueio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   6075
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "BLOQUEIO"
      Top             =   4980
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Timer Timer4 
      Left            =   15660
      Top             =   3525
   End
   Begin VB.Timer timerDescricaoBotoes 
      Left            =   4635
      Top             =   2160
   End
   Begin VB.Frame frmRelogio 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5000
      Left            =   9585
      TabIndex        =   45
      Top             =   5070
      Width           =   5000
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000000&
         Height          =   3645
         Left            =   650
         Picture         =   "frmPedido.frx":C737
         ScaleHeight     =   3645
         ScaleWidth      =   3750
         TabIndex        =   46
         Top             =   420
         Width           =   3750
         Begin VB.Timer Timer2 
            Interval        =   1
            Left            =   3045
            Top             =   3015
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   -1665
            TabIndex        =   47
            Top             =   30
            Width           =   2460
         End
         Begin VB.Line LS 
            BorderColor     =   &H00404040&
            BorderWidth     =   2
            X1              =   1485
            X2              =   1875
            Y1              =   1170
            Y2              =   720
         End
         Begin VB.Line LM 
            BorderColor     =   &H00FFFFFF&
            X1              =   1400
            X2              =   840
            Y1              =   1400
            Y2              =   810
         End
         Begin VB.Line LH 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   1020
            X2              =   510
            Y1              =   1545
            Y2              =   1335
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   105
         Top             =   765
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   49
         Top             =   4170
         Width           =   4995
      End
      Begin VB.Label lblDataInicial 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Quinta, 17 de julho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   0
         TabIndex        =   48
         Top             =   4710
         Width           =   4995
      End
   End
   Begin VB.PictureBox picLimitadorBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   90
      ScaleHeight     =   1995
      ScaleWidth      =   15150
      TabIndex        =   42
      Top             =   2625
      Width           =   15180
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   12000
         Left            =   -190
         TabIndex        =   43
         Top             =   -265
         Width           =   30000
         ExtentX         =   52917
         ExtentY         =   21167
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
         Location        =   "http:///"
      End
   End
   Begin VB.Frame fraCondicao 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   16515
      TabIndex        =   40
      Top             =   2190
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtCondicaoFaturado 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   16380
      TabIndex        =   38
      Top             =   1620
      Visible         =   0   'False
      Width           =   2190
   End
   Begin Project1.chameleonButton cmbPedido 
      Height          =   810
      Left            =   12120
      TabIndex        =   35
      Top             =   1650
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1429
      BTYPE           =   2
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":E3FB
      PICN            =   "frmPedido.frx":E417
      PICH            =   "frmPedido.frx":FA26
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdTR 
      Caption         =   "TR"
      Height          =   210
      Left            =   5265
      TabIndex        =   29
      Top             =   10860
      Visible         =   0   'False
      Width           =   345
   End
   Begin Project1.chameleonButton cmdLimpar 
      Height          =   0
      Left            =   5505
      TabIndex        =   22
      Top             =   10545
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "cmdLimpar"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":11035
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   3
      Left            =   2790
      TabIndex        =   14
      ToolTipText     =   "Agenda"
      Top             =   10815
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   7
      TX              =   "Agenda"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":11051
      PICN            =   "frmPedido.frx":1106D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   1
      Left            =   4390
      TabIndex        =   13
      ToolTipText     =   "Cliente"
      Top             =   10815
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   7
      TX              =   "Clientes"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":1178E
      PICN            =   "frmPedido.frx":117AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmbPedido2 
      Height          =   0
      Left            =   11790
      TabIndex        =   10
      Top             =   1185
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":11D8B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox PicBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H004E2A12&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   8790
      ScaleHeight     =   5565
      ScaleWidth      =   6480
      TabIndex        =   9
      Top             =   4905
      Visible         =   0   'False
      Width           =   6480
      Begin SHDocVwCtl.WebBrowser wbFichaTecnica 
         Height          =   6380
         Left            =   -30
         TabIndex        =   51
         Top             =   -250
         Width           =   6500
         ExtentX         =   11465
         ExtentY         =   11254
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
         Location        =   "http:///"
      End
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   65535
      Left            =   510
      Top             =   11940
   End
   Begin VB.PictureBox picQuadroGeral 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5790
      Left            =   90
      ScaleHeight     =   5790
      ScaleWidth      =   15165
      TabIndex        =   5
      Tag             =   "&H00AE7411&"
      Top             =   4680
      Width           =   15165
      Begin VB.Frame Frame2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3825
         Left            =   0
         TabIndex        =   41
         Top             =   675
         Width           =   8670
         Begin VSFlex7DAOCtl.VSFlexGrid grdItensProduto 
            Height          =   3885
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   8955
            _cx             =   15796
            _cy             =   6853
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
            FloodColor      =   3421236
            SheetBorder     =   8421504
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   14
            Cols            =   21
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPedido.frx":11DA7
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   7815
         MaxLength       =   4
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPesquisar 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   0
         TabIndex        =   0
         Top             =   240
         Width           =   7815
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   0
         ScaleHeight     =   885
         ScaleWidth      =   8685
         TabIndex        =   6
         Top             =   5130
         Width           =   8685
         Begin VB.PictureBox fradados 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   -15
            ScaleHeight     =   405
            ScaleWidth      =   8685
            TabIndex        =   7
            Top             =   240
            Width           =   8685
            Begin VB.TextBox txtVendedor 
               BackColor       =   &H00C0C0C0&
               Enabled         =   0   'False
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
               Left            =   1365
               MaxLength       =   3
               TabIndex        =   3
               Top             =   0
               Width           =   1320
            End
            Begin VB.TextBox txtPedido 
               Alignment       =   1  'Right Justify
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
               Left            =   15
               TabIndex        =   2
               Top             =   0
               Width           =   1320
            End
         End
         Begin VB.Label Label2 
            BackColor       =   &H00B63C18&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor"
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
            Left            =   1380
            TabIndex        =   50
            Top             =   15
            Width           =   2355
         End
         Begin VB.Label lblPedidoVendedor 
            BackColor       =   &H00B63C18&
            BackStyle       =   0  'Transparent
            Caption         =   " Pedido"
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
            Left            =   -60
            TabIndex        =   11
            Top             =   15
            Width           =   2355
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdDadosProduto 
         Height          =   570
         Left            =   0
         TabIndex        =   8
         Top             =   4545
         Width           =   8670
         _cx             =   15293
         _cy             =   1005
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
         BackColorBkg    =   5263440
         BackColorAlternate=   14737632
         GridColor       =   14737632
         GridColorFixed  =   8421504
         TreeColor       =   8421504
         FloodColor      =   16777215
         SheetBorder     =   8421504
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPedido.frx":12027
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
         BackColorFrozen =   -268435441
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   330
         OleObjectBlob   =   "frmPedido.frx":1210C
         Top             =   5310
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00B63C18&
         BackStyle       =   0  'Transparent
         Caption         =   " Referência/Fornecedor/Descrição/Codigo de Barras"
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
         Left            =   -75
         TabIndex        =   12
         Top             =   15
         Width           =   4845
      End
   End
   Begin Project1.chameleonButton cmdFechaPedido 
      Height          =   465
      Left            =   -200
      TabIndex        =   15
      ToolTipText     =   "Fecha Pedido"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "&Fechar (F1)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":12340
      PICN            =   "frmPedido.frx":1235C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   2
      Left            =   1185
      TabIndex        =   16
      ToolTipText     =   "Consulta"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "&Consulta (F12)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":12A94
      PICN            =   "frmPedido.frx":12AB0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   12
      Left            =   14500
      TabIndex        =   17
      ToolTipText     =   "Cotação"
      Top             =   10425
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "C&otação (F11)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":130F8
      PICN            =   "frmPedido.frx":13114
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   6
      Left            =   10500
      TabIndex        =   18
      ToolTipText     =   "Venda Distancia"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "&Venda Distância (F6)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":138AE
      PICN            =   "frmPedido.frx":138CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   9
      Left            =   13695
      TabIndex        =   19
      ToolTipText     =   "Desconto"
      Top             =   10425
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "&Desconto (F2)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":13F31
      PICN            =   "frmPedido.frx":13F4D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   10
      Left            =   12900
      TabIndex        =   20
      ToolTipText     =   "Frete"
      Top             =   10425
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "F&rete (F4)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":14697
      PICN            =   "frmPedido.frx":146B3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   7
      Left            =   11300
      TabIndex        =   21
      ToolTipText     =   "Carimbo"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "Car&imbo (F9)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":14EA6
      PICN            =   "frmPedido.frx":14EC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   0
      Index           =   0
      Left            =   420
      TabIndex        =   23
      ToolTipText     =   "F1"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "&Fechar (F1)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   15258293
      FCOLO           =   15258293
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":15695
      PICN            =   "frmPedido.frx":156B1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   4
      Left            =   12900
      TabIndex        =   24
      ToolTipText     =   "Lembre-me"
      Top             =   10815
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "Lembre-me"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":15DE9
      PICN            =   "frmPedido.frx":15E05
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   8
      Left            =   12100
      TabIndex        =   25
      ToolTipText     =   "Transferencia"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "Transferencia"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":16583
      PICN            =   "frmPedido.frx":1659F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdVersao 
      Height          =   0
      Left            =   6750
      TabIndex        =   26
      Top             =   10710
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":16CF6
      PICN            =   "frmPedido.frx":16D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   5
      Left            =   1995
      TabIndex        =   27
      ToolTipText     =   "E-mail Market"
      Top             =   10815
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   7
      TX              =   "Email Market"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":18874
      PICN            =   "frmPedido.frx":18890
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   11
      Left            =   14500
      TabIndex        =   28
      ToolTipText     =   "Comissão"
      Top             =   10815
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":19084
      PICN            =   "frmPedido.frx":190A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdTotalPedidoGE2 
      Height          =   0
      Left            =   11640
      TabIndex        =   31
      Top             =   780
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "0,00             "
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":19639
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   13
      Left            =   13695
      TabIndex        =   36
      ToolTipText     =   "Venda Distancia"
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "NEWS"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":19655
      PICN            =   "frmPedido.frx":19671
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox mskDatafaturado 
      Height          =   285
      Left            =   16350
      TabIndex        =   39
      Top             =   1170
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   12632256
      ForeColor       =   4210752
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
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Project1.chameleonButton cmdBotoes 
      Height          =   465
      Index           =   14
      Left            =   3590
      TabIndex        =   44
      ToolTipText     =   "Desconto"
      Top             =   10815
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   820
      BTYPE           =   2
      TX              =   "Calculadora"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":19D79
      PICN            =   "frmPedido.frx":19D95
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgIconBandeja 
      Height          =   240
      Left            =   375
      Picture         =   "frmPedido.frx":1A60B
      Top             =   255
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblQtde 
      BackColor       =   &H00B63C18&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde."
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
      Height          =   225
      Left            =   16515
      TabIndex        =   37
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label cmdQtdeItens 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9420
      TabIndex        =   34
      Top             =   1635
      Width           =   2565
   End
   Begin VB.Label cmdTotalPedidoGE 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   9420
      TabIndex        =   33
      Top             =   2250
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label cmdTotalPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   9420
      TabIndex        =   32
      Top             =   1905
      Width           =   2565
   End
   Begin VB.Label lblDescricaoBotao 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimento Caixa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10935
      TabIndex        =   30
      Top             =   11230
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image webInternet3 
      Height          =   2475
      Left            =   90
      Picture         =   "frmPedido.frx":1A995
      Stretch         =   -1  'True
      Top             =   100
      Width           =   15180
   End
   Begin VB.Image webInternet1 
      Appearance      =   0  'Flat
      Height          =   2025
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   2625
      Width           =   5670
   End
   Begin VB.Image webInternet2 
      Appearance      =   0  'Flat
      Height          =   2025
      Left            =   90
      Stretch         =   -1  'True
      Top             =   2625
      Width           =   9525
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' RELOGIO ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Dim tempoRestante As String
Dim wBanner As String
Dim wBanner2 As String
Dim cSize As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Const PI = 3.14159
Const LB_FINDSTRING = &H18F

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim auxVEN_CodigoVendedor As Integer
Dim auxItens As Double
Dim auxQtdeItens As Double
Dim wGuardaLinha As Long
Dim NroItens As Long
Dim Index As Integer
Dim wPosicaoLoja As String
Dim wVendedor As String
Dim wProtocolo As Integer
Dim SQL As String
Dim GuardaCor As String
Dim ReferenciaPreco As String
Dim NomeColuna As String
Dim wReferencia As String
Dim wQtde As Double
Dim wCodigo As Integer
Dim wPesquisaCodigo As Integer
Dim wSequencia As Integer
Dim wPesquisaSequencia As Integer
Dim wValorDados As String

Dim wFichaTec As String

Dim wSerie As Integer
Dim wIcms As Double
Dim wDesconto As Double
Dim wPreco As Double
Dim wPLISTA As Double
Dim wLinha As Integer
Dim wSecao As Integer
Dim wIcmPdv As Double
Dim wCodBarra As String
Dim wAliqIPI As Double
Dim wPrecoUnitAlternativa As Double
Dim wValorMercadoriaAlternativa As Double
Dim wReferenciaAlternativa As String
Dim wDescricaoAlternativa As String
Dim wWhere As String
Dim wValorVenda As Double
Dim wPrecoCalculado As Double
Dim wValorTotalCalculado As Double
Dim NomeParcela As String
Dim NomeCoefic As String
Dim QtdeParcelas As Integer
Dim wTotalPedido As Double
Dim auxPedido As String
Dim AuxProdutoExiste As Boolean
Dim wProdutoNaoExiste As Boolean
Dim auxVendedordoPedido As Integer
Dim wQuantidade As Integer
Dim wProdutoClasseP As Boolean
Dim wClasseProduto As String



'Variaveis do TR
Dim DescricaoAlternativaEmBranco As Boolean
Dim wValorTR As Double
Dim wValorSN As Double

'**************************************************
'Variaveis para tratar erro de script no WEBBROWSER
Dim WithEvents objDoc As MSHTMLCtl.HTMLDocument
Attribute objDoc.VB_VarHelpID = -1
Dim WithEvents objWind As MSHTMLCtl.HTMLWindow2
Attribute objWind.VB_VarHelpID = -1
Dim objEvent As CEventObj
'**************************************************

Dim wBannerReferencia As Boolean
Dim wBannerLinha As Boolean

Dim tempoMouseParado As Double

Dim Imagem As String
Dim wIndicePreco As String * 1

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal Hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Private Sub chameleonButton1_Click()
 If txtPesquisar.Enabled = True Then
    frmComissao.Show 1
    frmComissao.ZOrder
 End If
End Sub


Private Sub popupNomeBotao(nomeBotao As String, posicaoBotaoY)
    timerDescricaoBotoes.Enabled = True
    timerDescricaoBotoes.Interval = 500
    'lblDescricaoBotao.Visible = True
    lblDescricaoBotao.Caption = nomeBotao
    lblDescricaoBotao.left = posicaoBotaoY - 465
'   lblDescricaoBotao.left = nomeBotao.
End Sub

Private Sub cmbCliente_Click()

End Sub

Private Sub cmdBotoes_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    popupNomeBotao cmdBotoes(Index).Caption, cmdBotoes(Index).left
End Sub

Private Sub cmdBotoes_Click(Index As Integer)
    Select Case Index
    Case 0
        primeiroF1
    Case 1
        If txtPesquisar.Enabled = True Then
            wClienteTelaAdicionais = False
            frmConsCliente.ZOrder
            frmConsCliente.Show 1
        End If
        
    Case 2
        On Error GoTo ErronaDelecao
        
        adoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        SQL = "Delete NFItens Where NumeroPed = " & txtPedido.Text & " and TipoNota = 'PD'"
        adoCNLoja.Execute SQL
        SQL = "Delete CarimboNotaFiscal where cnf_NumeroPed = " & txtPedido.Text
        adoCNLoja.Execute SQL
        SQL = "Delete NFCapa Where TipoNota = 'PD' and NumeroPed = " & txtPedido.Text
        adoCNLoja.Execute SQL
        Screen.MousePointer = vbNormal
        adoCNLoja.CommitTrans
        Call LimpaForm
        Exit Sub
        
ErronaDelecao:
        MsgBox "Erro na deleção do pedido " & Err.description, vbCritical, "Aviso"
        adoCNLoja.RollbackTrans
        Screen.MousePointer = vbNormal
        
    Case 3
        If txtPesquisar.Enabled = True Then
            frmAgenda.Show 1
            frmAgenda.ZOrder
        End If
        
    Case 4
        frmLembreMe.Show 1
        frmLembreMe.ZOrder
    Case 5
        If txtPesquisar.Enabled = True Then
            frmEmarketing.Show 1
            frmEmarketing.ZOrder
        End If
    
    Case 6
        frmVendaDistancia.Show 1
        frmVendaDistancia.ZOrder
        
    Case 7
        frmCarimbos.Show 1
        frmCarimbos.ZOrder
        
    Case 8
        frmTransferencia.Show 1
        frmTransferencia.ZOrder
        
    Case 9
        frmDesconto.Show 1
        frmDesconto.ZOrder
        
    Case 10
        frmFrete.Show 1
        frmFrete.ZOrder
        
    Case 11
        If txtPesquisar.Enabled = True Then
            frmComissao.Show 1
            frmComissao.ZOrder
        End If

    Case 12
        FrmCotacao.Show 1
        FrmCotacao.ZOrder
        
    Case 13
        FrmNews.Show 1
        FrmNews.ZOrder
        
    Case 14
        frmCalculadora.Show 1
        frmCalculadora.ZOrder

    End Select
End Sub

Private Sub cmdBotoes_MouseOut(Index As Integer)
    timerDescricaoBotoes.Enabled = False
    lblDescricaoBotao.Visible = False
End Sub

Private Sub CmdDesfaz_Click()
  FrmDesfazProcesso.txtPedido = txtPedido.Text
  FrmDesfazProcesso.Show 1
  FrmDesfazProcesso.ZOrder
End Sub

Private Sub cmdFechaPedido_Click()
    Call FechaPedido
 '    frmFinalizaPedido.Show 1
 '    frmFinalizaPedido.ZOrder
End Sub


Private Sub cmdTotalPedido_Change()

    If cmdTotalPedido.Caption <> "" Then
        wValor = cmdTotalPedido.Caption
    End If
    
End Sub

Private Sub cmdTR_Click()
  frmTR.Show 1
  frmTR.ZOrder
End Sub

Private Sub cmdVersao_Click()
    MsgBox "Versão " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Sobre"
End Sub

Private Sub carregaDataInicial()
    Dim semana As String

    semana = Format(Date, "dddd")
    If UCase(semana) = "SÁBADO" Or UCase(semana) = "DOMINGO" Then
        semana = UCase(left(semana, 1)) & Mid(semana, 2, Len(semana))
    Else
        semana = UCase(left(semana, 1)) & Mid(semana, 2, InStr(1, semana, "-", 1) - 2)
    End If
    'semana =
    
    lblDataInicial.Caption = semana & ", " & Format(Date, "d") & " de " & Format(Date, "mmmm")
    
End Sub

Private Sub Command1_Click()
    frmRelogio.left = (PicBanner.left - (PicBanner.left \ 2))
    frmRelogio.top = (PicBanner.top - (PicBanner.top \ 2))
End Sub

Private Sub Form_Activate()
    wbFichaTecnica.Height = 6050
    WebBrowser1.Navigate (wBanner2)

End Sub

Private Sub Form_Click()
    If txtPedido.Enabled And txtPedido.Visible Then
        txtPedido.SetFocus
    'ElseIf txtPesquisar.Enabled And txtPesquisar.Visible Then
        'txtPesquisar.SetFocus
    End If
End Sub

Private Sub Form_Load()

Timer4.Enabled = True

cmdVersao.Height = 615
cmdBotoes(0).Height = 460
tempoRestante = "00:00:10"

picLimitadorBanner.ZOrder 0
picLimitadorBanner.Height = 2025



'Call criaIconeBarra(TrayAdd, Me.Hwnd, Me.Caption, imgIconBandeja.Picture)

resolucaoOriginal.Colunas = resolucaoTela.Colunas
resolucaoOriginal.Linhas = resolucaoTela.Linhas
Call AlterarResolucao(1024, 768)

carregaDataInicial

PicBanner.Visible = False
'PicBanner.BackColor = vbBlack
frmRelogio.BackColor = vbBlack

frmRelogio.BackColor = vbBlack

left = (Screen.Width - Width) / 2
top = (Screen.Height - Height) / 2
cmdBotoes(12).top = cmdBotoes(0).top
cmdBotoes(10).top = cmdBotoes(0).top
cmdBotoes(9).top = cmdBotoes(0).top

'frmPedido.txtVendedor.Width = 915
'frmPedido.fradados.Width = 1830

  'cmdVersao.Caption = ""
  
 On Error GoTo erro
 Call LerControleSistema
 'Call LerControleCaixa''''''''''''''''''''AQUI
 Call VerificaInternet
 
 wbFichaTecnica.Visible = True
 'PicBanner
 'wbFichaTecnica.Navigate "C:\Sistemas\DMAC Venda\desc.HTML"
 'frmPedido.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\frmpedido1024768hd")
 webInternet3.Picture = LoadPicture(endIMG("topo1024768hd"))
 'webBannnerChamada.Navigate ("C:\Sistemas\DMAC Venda\Imagens\BannerChamada\configBannerChamada")
 WebBrowser1.Navigate (wBanner2)
 'webInternet1.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerTopo2\BannerTopo2.swf")
 'webInternet2.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerTopo1\BannerTopo1a.swf")
 NroBanner = 1

 grdItensProduto.Enabled = False
 'grdPrecos.Enabled = False
 grdDadosProduto.Enabled = False
 txtQuantidade.Enabled = False
 wClienteTelaAdicionais = False
 GBL_Frete = 0
 PicBanner.ZOrder
 RemoveMenus
  
  
 Dim OMes As String

' RELOGIO

'OMes = Format(Date, "mmmm")
    'If App.PrevInstance Then End
    'List1.ListIndex = SendMessage(List1.hWnd, LB_FINDSTRING, -1, ByVal CStr(OMes))
    'Image1.ToolTipText = Format(Date, " dddd" & ", " & "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy ")
    'Label2.Caption = Format(Date, "dd") & vbCrLf & List1.List(List1.ListIndex) & vbCrLf & Format(Date, "yyyy ")
    'Dim hr&, dl&
    'Dim usew&, useh&
    'usew& = Me.Width / Screen.TwipsPerPixelX - 2
    'useh& = Me.Height / Screen.TwipsPerPixelY - 2
'    X = 0
    'Deixa o Form ficar em forma círcular
    'hr& = CreateEllipticRgn(4, 4, usew, useh)
    'dl& = SetWindowRgn(Me.hWnd, hr, True)
    
    'cSize = CreateEllipticRgn(4, 23, 270, 290)
    'SetWindowRgn Me.Hwnd, cSize, True
    Call Timer1_Timer
    'NoTopoSim Me.Hwnd
  
    'picLimitadorBanner.Height = 7850
    'WebBrowser1.Navigate ("C:\Sistemas\DMAC Venda\Imagens\BannerTopo1\BannerTopo1b.GIF")
    WebBrowser1.SetFocus
    

  
erro:
    Exit Sub
End Sub

Private Sub grdItensProduto_EnterCell()
 On Error GoTo trata_erro
 Dim codigoHTML As String
 Dim SQL As String

    Screen.MousePointer = 11

    If Mid(grdItensProduto.TextMatrix(grdItensProduto.Row, 9), 1, 1) = "P" Then
        grdDadosProduto.BackColor = &HFF&
        grdDadosProduto.ForeColor = vbWhite
    ElseIf Mid(grdItensProduto.TextMatrix(grdItensProduto.Row, 20), 1, 3) = "180" Then
        grdDadosProduto.BackColor = &H800080
        grdDadosProduto.ForeColor = vbWhite
    ElseIf Mid(grdItensProduto.TextMatrix(grdItensProduto.Row, 19), 1, 3) = "180" Then
        grdDadosProduto.BackColor = &H80FF&
        grdDadosProduto.ForeColor = vbWhite
    Else
        grdDadosProduto.BackColor = &HE0E0E0
        grdDadosProduto.ForeColor = vbBlack
        'grdItensProduto.BackColorSel = &H343434
    End If
 
  'enderecoImagem = "C:\Sistemas\DMAC Venda\Imagens\BannerAcessorios\ET" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0)
  'PicBanner.Picture = LoadPicture(enderecoImagem)
  'wbFichaTecnica.Navigate "C:\Sistemas\DMAC Venda\desc.HTML"
    
    wbFichaTecnica.Visible = True
    SQL = "select top 1 PRO_DESCR_LONGA as Descricao, " & vbNewLine & _
    "PRO_ITENS_INCLUSOS as DescricaoItens, " & vbNewLine & _
    "PRO_ESPECIFICACAO_SITE as DescricaoEspecificacao " & vbNewLine & _
    "from produtodescricao " & vbNewLine & _
    "where pro_referencia = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "'"
    
    rdoDescricao.CursorLocation = adUseClient
    rdoDescricao.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
    If Not rdoDescricao.EOF Then
        codigoHTML = "<html>" & vbNewLine & _
        "<body bgcolor=" & Chr(34) & "#C0C0C0" & Chr(34) & "><font face=" & Chr(34) & "Verdana, Arial, Helvetica, sans-serif" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & " color=" & Chr(34) & "#000000" & Chr(34) & "><br>"
        
        If Trim(rdoDescricao("DescricaoEspecificacao")) <> "" Then
            codigoHTML = codigoHTML & "<font face=" & Chr(34) & "Verdana size=" & Chr(34) & "4" & Chr(34) & "" & Chr(34) & "><b><i>Especificação</b></font></i><br><p align=" & Chr(34) & "Justify" & Chr(34) & ">" & vbNewLine & _
            rdoDescricao("DescricaoEspecificacao") & "</p>"
        End If
        
        If Trim(rdoDescricao("Descricao")) <> "" Then
            codigoHTML = codigoHTML & "<br><br><p align=" & Chr(34) & "Justify" & Chr(34) & ">" & vbNewLine & _
            rdoDescricao("Descricao") & "</p>"
        End If
        
        If Trim(rdoDescricao("DescricaoItens")) <> "" Then
            codigoHTML = codigoHTML & "<br><br><font face=" & Chr(34) & "Verdana size=" & Chr(34) & "4" & Chr(34) & "" & Chr(34) & "><b><i>Itens</i></b></font><br><p align=" & Chr(34) & "Justify" & Chr(34) & ">" & vbNewLine & _
            rdoDescricao("DescricaoItens") & "</p>"
        End If
        
        
        codigoHTML = codigoHTML & "<br><br></p></font></body>" & vbNewLine & "</html>"
        
        PicBanner.Visible = True
    Else
        PicBanner.Visible = False
    End If
    
    rdoDescricao.Close
  
    Open "C:\Sistemas\DMAC Venda\desc.HTML" For Output As #1
    Print #1, codigoHTML
    Close #1
  
  wbFichaTecnica.Navigate "C:\Sistemas\DMAC Venda\desc.HTML"
  
  Screen.MousePointer = 0
  grdItensProduto.SetFocus
  
  'wbFichaTecnica.Navigate
  
trata_erro:
    
    Screen.MousePointer = 0
  If Err.Number = 53 Then
    PicBanner.Visible = False
  'Else
     'MsgBox "Ocorreu o erro : " & Err.Number & " - " & Err.description
  End If
End Sub

Private Sub Timer1_Timer()
    Dim H As Single, m As Single, s As Single
    Dim TotHours As Single
    Dim PtCentro As Integer
    Dim PtCentro2 As Integer
    Label3.Caption = Format(Time, "hh:nn:ss")
    PtCentro = Picture3.Width / 2
    PtCentro2 = Picture3.Height / 2
    LH.X1 = PtCentro
    LH.Y1 = PtCentro2
    LM.X1 = PtCentro
    LM.Y1 = PtCentro2
    LS.X1 = PtCentro
    LS.Y1 = PtCentro2
    
    H = Hour(Time)
    m = Minute(Time)
    s = Second(Time)
    TotHours = H + m / 60
    
    LH.X2 = 800 * Cos(PI / 180 * (30 * TotHours - 90)) + LH.X1
    LH.Y2 = 800 * Sin(PI / 180 * (30 * TotHours - 90)) + LH.Y1
    
    LM.X2 = 1100 * Cos(PI / 180 * (6 * m - 90)) + LH.X1
    LM.Y2 = 1100 * Sin(PI / 180 * (6 * m - 90)) + LH.Y1
    
    LS.X2 = 800 * Cos(PI / 180 * (6 * s - 90)) + LH.X1
    LS.Y2 = 800 * Sin(PI / 180 * (6 * s - 90)) + LH.Y1

End Sub

Private Sub grdItensProduto_DblClick()
      
 On Error GoTo trata_erro
 Dim enderecoImagem As String
 
  enderecoImagem = "C:\Sistemas\DMAC Venda\Imagens\EspecificacaoTecnica\ET" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0)
  PicBanner.Picture = LoadPicture(enderecoImagem)
  PicBanner.Visible = True
  

  
  
trata_erro:
  

  If Err.Number = 53 Then
    PicBanner.Visible = False
  'Else
     'MsgBox "Ocorreu o erro : " & Err.Number & " - " & Err.description
  End If

      
'      wbFichaTecnica.Navigate2 wFichaTec & grdItensProduto.TextMatrix(grdItensProduto.Row, 0)
'      wbFichaTecnica.Visible = True
'     wbFichaTecnica.ZOrder

'    PicBanner.Visible = False

'navegador.Visible = False

End Sub



Private Sub grdItensProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       txtPesquisar.Enabled = True
       txtPesquisar.SetFocus
       txtPesquisar.SelStart = 0
       txtPesquisar.SelLength = Len(txtPesquisar.Text)
    ElseIf KeyCode = 13 Then
'       grdItensProduto.BackColorSel = &HC0C000
       
       'If grdPrecos.Enabled = True Then
          'grdPrecos.SetFocus
          'grdPrecos.Row = 1
       If lblBloqueio.Text = "BLOQUEADO" And lblBloqueio.Visible = True Then
          MsgBox "Você não pode vende produto com o Bloqueio 9", vbExclamation, "DMAC Venda"
       Else
          txtQuantidade.Enabled = True
          txtQuantidade.SetFocus
       End If
    ElseIf KeyCode = vbKeyF4 Then
        frmPesquisaEstoqueCentral.ZOrder
        frmPesquisaEstoqueCentral.Show 1
    End If
End Sub

Private Sub grdItensProduto_LostFocus()
 'PicBanner.Visible = False
End Sub

Private Sub grdItensProduto_RowColChange()

tempoRestante = "00:00:10"

 txtPesquisar.Text = grdItensProduto.TextMatrix(grdItensProduto.Row, 1)
 grdDadosProduto.TextMatrix(1, 0) = grdItensProduto.TextMatrix(grdItensProduto.Row, 16)  'Fornecedor
 grdDadosProduto.TextMatrix(1, 1) = grdItensProduto.TextMatrix(grdItensProduto.Row, 7)  'Linha]
 grdDadosProduto.TextMatrix(1, 2) = grdItensProduto.TextMatrix(grdItensProduto.Row, 18)  'Garantia
 grdDadosProduto.TextMatrix(1, 3) = grdItensProduto.TextMatrix(grdItensProduto.Row, 10)  'Bloqueio
 grdDadosProduto.TextMatrix(1, 4) = grdItensProduto.TextMatrix(grdItensProduto.Row, 9)  'Classe
 grdDadosProduto.TextMatrix(1, 5) = grdItensProduto.TextMatrix(grdItensProduto.Row, 11) 'ICMS
 grdDadosProduto.TextMatrix(1, 6) = grdItensProduto.TextMatrix(grdItensProduto.Row, 17) 'ST
 grdDadosProduto.TextMatrix(1, 7) = grdItensProduto.TextMatrix(grdItensProduto.Row, 15) 'NCM
 
 wIndicePreco = grdItensProduto.TextMatrix(grdItensProduto.Row, 14)                     'indicePreco
 wValorVenda = Format(grdItensProduto.TextMatrix(grdItensProduto.Row, 2), "0.00")

lblBloqueio.Visible = True
 Select Case grdDadosProduto.TextMatrix(1, 3)
    Case "1"
        lblBloqueio.Text = "Encomenda"
    Case "2"
        lblBloqueio.Text = "Fora de Linha"
    Case "4"
        lblBloqueio.Text = "Especial"
    Case "9"
        lblBloqueio.Text = "BLOQUEADO"
    Case Else
        lblBloqueio.Visible = False
 End Select
 
 'If grdPrecos.TextMatrix(0, 0) = "Financiado" Then
    'Call MontaPrecos("FI", wIndicePreco)
 'ElseIf grdPrecos.TextMatrix(0, 0) = "Faturado" Then
 '   Call MontaPrecos("FA", wIndicePreco)
 'ElseIf grdPrecos.TextMatrix(0, 0) = "A Vista" Then
 '   Call MontaPrecos("AV", wIndicePreco)
 'ElseIf grdPrecos.TextMatrix(0, 0) = "Cartão" Then
 '   Call MontaPrecos("CC", wIndicePreco)
 'End If
 
 'navegador.Visible = True
 
End Sub

Private Sub grdPrecos_DblClick()
''    If grdPrecos.Col = 0 And cmdQtdeItens.Caption = 0 Then
''       grdPrecos.Enabled = True
''
''       If grdPrecos.TextMatrix(0, 0) = "Financiado" Then
''          grdPrecos.TextMatrix(0, 0) = "Cartão"
''          Call MontaPrecos("CC", wIndicePreco)
''       ElseIf grdPrecos.TextMatrix(0, 0) = "Faturado" Then
''          grdPrecos.TextMatrix(0, 0) = "A Vista"
''          Call MontaPrecos("AV", wIndicePreco)
''       ElseIf grdPrecos.TextMatrix(0, 0) = "A Vista" Then
''          grdPrecos.TextMatrix(0, 0) = "Financiado"
''          Call MontaPrecos("FI", wIndicePreco)
''       ElseIf grdPrecos.TextMatrix(0, 0) = "Cartão" Then
''          grdPrecos.TextMatrix(0, 0) = "Faturado"
''          Call MontaPrecos("FA", wIndicePreco)
''       End If
''    Else
''       grdPrecos.Enabled = False
''    End If
End Sub

Private Sub grdPrecos_EnterCell()

''        txtCondicaoFaturado.Visible = True
''        txtCondicaoFaturado.Text = grdPrecos.TextMatrix(grdPrecos.Row, 2)
''        mskDatafaturado.Text = "__/__/____"

End Sub

Private Sub grdPrecos_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode = 27 Then
''
''          grdItensProduto.SetFocus
''          txtPesquisar.Enabled = True
''          txtPesquisar.SetFocus
''
''    ElseIf KeyCode = 13 Then
''
''        If grdPrecos.TextMatrix(grdPrecos.Row, 0) = "85" Then
''          fraCondicao.Enabled = True
'''         txtCondicaoFaturado.SetFocus
''          txtCondicaoFaturado.Visible = False
''          txtCondicaoFaturado.Text = ""
''          mskDatafaturado.SetFocus
''
''         Exit Sub
''        End If
''
'' '          grdPrecos.BackColorSel = &HC00000
''           txtQuantidade.Enabled = True
''           txtQuantidade.SetFocus
''           grdPrecos.Enabled = False
''
''    End If

End Sub

Private Sub cmbPedido_Click()
If cmdQtdeItens.Caption <> 0 Then
    frmConsultaItensdoPedido.Show 1
    frmConsultaItensdoPedido.ZOrder
ElseIf frmPedido.txtPedido.Text = "" Then
    frmConsultaPedido.Show 1
    frmConsultaPedido.ZOrder
End If
 
End Sub

Private Sub cmbPedido_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbCustom
End Sub


Private Sub Text2_Change()

End Sub

Private Sub lblPesquisa1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub mskDatafaturado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If IsDate(mskDatafaturado.Text) = False Then
        MsgBox "Data inválida.", vbCritical, Me.Caption
        mskDatafaturado.SetFocus
        Exit Sub
       End If
       txtQuantidade.Enabled = True
       txtQuantidade.SetFocus
       wGuardaPagamento = mskDatafaturado.Text
  End If

End Sub

Private Sub mskDatafaturado_LostFocus()
    If fraCondicao.Enabled = True Then
        fraCondicao.Enabled = False
        'If mskDatafaturado.Text = "" Then
            'grdPrecos.SetFocus
        'Else
            'wGuardaLinha = grdPrecos.Row
        'End If
    End If
End Sub


Private Sub s_Change()

End Sub



Private Sub Timer3_Timer()

End Sub

Private Sub Timer4_Timer()
    If tempoRestante <> "00:00:00" Then
        tempoRestante = DateAdd("s", -1, tempoRestante)
    Else
'        MsgBox "OI"
'        WebBrowser1.Refresh
        WebBrowser1.SetFocus
    End If
End Sub

Private Sub timerDescricaoBotoes_Timer()
    tempoMouseParado = tempoMouseParado + 1
    If tempoMouseParado >= 2 Then
        lblDescricaoBotao.Visible = True
        timerDescricaoBotoes.Enabled = False
    End If
End Sub

Private Sub tmrRefresh_Timer()

    Call LerControleSistema
    Call VerificaInternet

End Sub


Private Sub tmrTroca_Timer()

End Sub

'Public Sub ExecutarAcao()
    
'    Temporizador.Enabled = False
             
'End Sub
 
 
'Private Sub tmrTroca_Timer()
'    Call TrocaBannerTopo1
'End Sub



Private Sub txtPedido_Change()
    
 If IsNumeric(txtPedido.Text) = False Then
   txtPedido.Text = ""
End If

End Sub

Private Sub txtPedido_GotFocus()

   'Timer4.Enabled = True
   ShellExecute Hwnd, "open", ("C:\Sistemas\DMAC Venda\TrocaVersao.exe"), "", "", 1
   WebBrowser1.Navigate (wBanner2)
   
   cmdBotoes(4).Visible = False
   cmdBotoes(1).Visible = False
   cmdBotoes(5).Visible = False
   cmdBotoes(3).Visible = False
   cmdBotoes(0).Visible = False
   
   cmdBotoes(11).Visible = False
   cmdBotoes(13).Visible = False
   cmdBotoes(14).Visible = False
   
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   txtPedido.Text = ""
   sairDoSistema
   End
End If

If KeyAscii = 46 Then
      txtPedido.Text = 0
      txtPedido.SelStart = 0
      txtPedido.SelLength = Len(txtPedido.Text)
      txtPedido.SetFocus
      Exit Sub
   End If
   
   If KeyAscii = 44 Then
      txtPedido.Text = 0
      txtPedido.SelStart = 0
      txtPedido.SelLength = Len(txtPedido.Text)
      txtPedido.SetFocus
      Exit Sub
   End If

If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
   If txtPedido.Text <> "" Then
      If IsNumeric(txtPedido.Text) = False Then
         txtPedido.Text = ""
         Exit Sub
      End If
   End If
   
   If txtPedido.Text = "" Then
      SQL = "Select * from ControleSistema"
           rsPegaNumeroPedido.CursorLocation = adUseClient
           rsPegaNumeroPedido.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
 
           On Error GoTo erronoUpdate
        
           If Not rsPegaNumeroPedido.EOF Then
              adoCNLoja.BeginTrans
              Screen.MousePointer = vbHourglass
              SQL = ""
              SQL = "Update ControleSistema set CTS_NumeroPedido=(CTS_NumeroPedido + 1)"
                    adoCNLoja.Execute SQL
                    Screen.MousePointer = vbNormal
                    adoCNLoja.CommitTrans
            
              txtPedido.Text = (rsPegaNumeroPedido("CTS_NumeroPedido"))
              auxPedido = (rsPegaNumeroPedido("CTS_NumeroPedido"))
              txtPedido.Enabled = False
              txtVendedor.Enabled = True
              'txtPesquisar.Enabled = True
              txtVendedor.SetFocus
              rsPegaNumeroPedido.Close
              Exit Sub
          Else
              MsgBox "Erro no Controle do Sistema avise o CPD"
              rsPegaNumeroPedido.Close
          End If
 
   Else
      Call VerificaItensVendas
      SomaItensVenda
      If cmdLimpar.Caption = "Pedido não cadastrado ou encerrado" Then
         txtPedido.SetFocus
         txtVendedor.Enabled = False
         txtPesquisar.Enabled = False
         txtQuantidade.Enabled = False
         Exit Sub
      End If
      
'''      SQL = ""
'''      SQL = "Select vendedor From nfcapa Where numeroped = " & txtPedido.Text
'''      rsVendedor.CursorLocation = adUseClient
'''      rsVendedor.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'''
'''      auxVendedordoPedido = rsVendedor("Vendedor")
'''      rsVendedor.Close
      
'      Call LerVendedordoPedido
      
      fradados.Enabled = True
      txtVendedor.Enabled = True
      txtPedido.Enabled = False
      cmbPedido.Visible = True
      cmdBotoes(1).Visible = True
      cmdBotoes(2).Visible = True
      cmdBotoes(4).Visible = True
      cmdBotoes(11).Visible = True
      cmdBotoes(13).Visible = True
 cmdBotoes(14).Visible = True
      txtPesquisar.Enabled = True
      txtVendedor.SetFocus
      
'      rsPegaNumeroPedido.Close

      wPesquisaCodigo = 1
      inibebotoes (frmPedido.txtPedido)
     
      Exit Sub
   End If
  
End If
Exit Sub
erronoUpdate:
MsgBox "Erro na atualização do número do pedido " & Err.description, vbCritical, "Aviso"
adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
rsPegaNumeroPedido.Close

End Sub

Private Sub txtPedido_LostFocus()

Timer4.Enabled = False

If frmPedido.txtPedido.Text = "" And frmPedido.txtPedido.Enabled = False Then
   frmPedido.txtPedido.SetFocus
End If

End Sub

Private Sub txtPesquisar_Change()

tempoRestante = "00:00:10"

 If txtPesquisar.Text = "'" Then
     MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
     txtPesquisar.Text = ""
     txtPesquisar.SetFocus
     Exit Sub
 End If
 If cmdFechaPedido.Visible = True Then
    
    If cmdQtdeItens.Caption > 0 Then
        cmdBotoes(0).Visible = True
        
    Else
        cmdBotoes(0).Visible = False
        
    End If
    cmdFechaPedido.Visible = False
    cmdBotoes(2).Visible = False
    cmdBotoes(12).Visible = False
    cmdBotoes(6).Visible = False
    cmdBotoes(9).Visible = False
    cmdBotoes(8).Visible = False
    cmdBotoes(10).Visible = False
    
    cmdTR.Visible = False
    cmdBotoes(7).Visible = False
    'PicBanner.Visible = True
    cmbPedido.Visible = True
    cmdBotoes(1).Visible = True
    cmdBotoes(2).Visible = True
    cmdBotoes(4).Visible = True
    cmdBotoes(11).Visible = True
    cmdBotoes(13).Visible = True
 cmdBotoes(14).Visible = True
    
    cmdTotalPedidoGE.Visible = False
    
    GBL_Frete = 0
    
    SQL = "Select sum(vltotitem) as vltotitem From Nfitens Where NumeroPed = " & frmPedido.txtPedido.Text
    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    frmPedido.cmdTotalPedido.Caption = Format(rsComplementoVenda("vltotitem") + GBL_Frete, "###,###,###,##0.00")
     rsComplementoVenda.Close
     
    'Do While Len(frmPedido.cmdTotalPedido.Caption) <= 12
      'frmPedido.cmdTotalPedido.Caption = frmPedido.cmdTotalPedido.Caption '+ " "
    'Loop
    
 End If
 
End Sub

Private Sub txtPesquisar_GotFocus()
   lblBloqueio.Visible = False
   txtPesquisar.SelStart = 0
   txtPesquisar.SelLength = Len(txtPesquisar.Text)
   'PicBanner.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerChamada\BannerChamada.jpg")
   'webBannnerChamada.Navigate ("C:\Sistemas\DMAC Venda\Imagens\BannerChamada\configBannerChamada")
'   If wbFichaTecnica.Visible = True Then
'      wbFichaTecnica.Visible = False
'      PicBanner.Visible = True
'   End If
End Sub

Private Sub primeiroF1()
   Call VerificaItensVendas
   If auxQtdeItens <> 0 Then
        'If Mid(wVendedor, 1, 3) = "790" Then
        '    Call FechaPedido
        '    Exit Sub
        'End If
     If cmdFechaPedido.Visible = True Then
            Call FechaPedido
     Else
        'If itemComGarantiaEstendida(txtpedido.Text) = True Then
            'frmGarantiaEstendida.Show 1
            'frmGarantiaEstendida.ZOrder
        'End If
            frmTrocaModalidadeVenda.Show 1
            frmTrocaModalidadeVenda.ZOrder
            
            If wGravaModalidade = True Then
                carregaProdutoGarantia
                frmAdicionais.Show 1
                frmAdicionais.ZOrder
            End If
     
     End If
  End If
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    primeiroF1
End If

If KeyCode = vbKeyF2 Then
      If cmdFechaPedido.Visible = True Then
         If cmdBotoes(9).Visible = False Then
            Exit Sub
         End If
         frmDesconto.Show 1
         frmDesconto.ZOrder
         frmDesconto.txtTotalPedido.Text = Trim(frmPedido.cmdTotalPedido.Caption)
         frmDesconto.txtPedido = frmPedido.txtPedido.Text
      Else
         txtPesquisar.SetFocus
      End If
End If

If KeyCode = vbKeyF4 Then
    If cmdFechaPedido.Visible = True Then
         If cmdBotoes(10).Visible = False Then
            Exit Sub
         End If
         frmFrete.txtTotalPedido.Text = Trim(cmdTotalPedido.Caption)
         frmFrete.txtPedido = txtPedido.Text
         frmFrete.Show 1
         frmFrete.ZOrder
    Else
         txtPesquisar.SetFocus
    End If
End If


If KeyCode = vbKeyF6 Then
      If cmdFechaPedido.Visible = True Then
         If cmdBotoes(6).Visible = False Then
            Exit Sub
         End If
         frmVendaDistancia.Show 1
         frmVendaDistancia.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
 End If
 
 If KeyCode = vbKeyF7 Then
      If cmdFechaPedido.Visible = True Then
         If cmdTR.Visible = False Then
            Exit Sub
         End If
         frmTR.Show 1
         frmTR.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
 End If
 
 If KeyCode = vbKeyF9 Then
     If cmdFechaPedido.Visible = True Then
         If cmdBotoes(7).Visible = False Then
            Exit Sub
         End If
         
         frmCarimbos.txtPedido = txtPedido.Text
         frmCarimbos.Show 1
         frmCarimbos.ZOrder
     Else
         txtPesquisar.SetFocus
     End If
End If

If KeyCode = vbKeyF11 Then
    If cmdFechaPedido.Visible = True Then
        FrmCotacao.Show 1
        FrmCotacao.ZOrder
    Else
        txtPesquisar.SetFocus
    End If
End If
  
If KeyCode = vbKeyF12 Then
'***********************
'      Deleta Capa e Itens na consulta
       On Error GoTo ErronaDelecao
          adoCNLoja.BeginTrans
          Screen.MousePointer = vbHourglass
          SQL = "Delete NFItens Where NumeroPed = " & txtPedido.Text & " and TipoNota = 'PD'"
          adoCNLoja.Execute SQL
          
          SQL = "Delete NFCapa Where TipoNota = 'PD' and NumeroPed = " & txtPedido.Text
          adoCNLoja.Execute SQL
          
          SQL = "Delete CarimboNotaFiscal Where CNF_NumeroPed = " & txtPedido.Text
          adoCNLoja.Execute SQL
          
          Screen.MousePointer = vbNormal
          adoCNLoja.CommitTrans
          Call LimpaForm
          
End If
'***********************
  If KeyCode = vbKeyTab Then
         txtPesquisar.SetFocus
  End If

Exit Sub

ErronaDelecao:
''MsgBox "Erro na deleção do pedido " & Err.description, vbCritical, "Aviso"
''adoCNLoja.RollbackTrans
''Screen.MousePointer = vbNormal

End Sub

Private Sub txtPesquisar_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  If cmdFechaPedido.Visible = True Then

    If cmdQtdeItens.Caption > 0 Then
        cmdBotoes(0).Visible = True
        
    Else
        cmdBotoes(0).Visible = False
        
    End If
    
    'cmdBotoes(0).Visible = True
    cmdBotoes(2).Visible = True
    cmdFechaPedido.Visible = False
    cmdBotoes(2).Visible = False
    cmdBotoes(12).Visible = False
    cmdBotoes(8).Visible = False
    cmdBotoes(6).Visible = False
    cmdBotoes(9).Visible = False
    cmdBotoes(10).Visible = False
    
    cmdTR.Visible = False
    cmdBotoes(7).Visible = False

    'PicBanner.Visible = True
  End If

    If txtPesquisar.Text <> "" Then
        If IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) = 3 Then
            wWhere = "PRB_Tipocodigo = 'D' and PR_CodigoFornecedor =" & Trim(txtPesquisar.Text) & " "
            PesquisarProduto wWhere ' Pesquisa por fonecedor
        ElseIf IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) = 7 Then
            wWhere = "PRB_Tipocodigo = 'D' and PR_Referencia ='" & Trim(txtPesquisar.Text) & "' "
            PesquisarProduto wWhere ' Pesquisa por referencia
        ElseIf IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) > 3 Then
            wWhere = "PRB_CodigoBarras = '" & Trim(txtPesquisar.Text) & "' "
            PesquisarProduto wWhere ' Pesquisa por codigo de barras
        ElseIf IsNumeric(Trim(txtPesquisar.Text)) = False Then
            If IsNumeric(Mid(txtPesquisar.Text, 1, 3)) = True And Trim(Mid(txtPesquisar.Text, 4, 1)) = "" Then
                 wWhere = "PRB_Tipocodigo = 'D' and PR_Descricao Like '" & Trim(UCase(Mid(Trim(txtPesquisar.Text), 4, _
                 Len(Trim(Trim(txtPesquisar.Text)))))) & "%' and PR_CodigoFornecedor = " _
                 & Mid(txtPesquisar, 1, 3)
                PesquisarProduto wWhere  ' Pesquisa por Fornecedor e Descrição
            Else
                wWhere = "PRB_Tipocodigo = 'D' and PR_Descricao Like '" & UCase(Trim(txtPesquisar.Text)) & "%' "
                PesquisarProduto wWhere  ' Pesquisa por descrição
            End If
        Else
             cmdLimpar.Caption = "Pesquisa Inválida"
             Exit Sub
        End If

            grdItensProduto.Enabled = True
            
            If cmdQtdeItens.Caption > 0 Then
                SQL = ""
                SQL = "Select ModalidadeVenda, Parcelas From NFCapa Where Numeroped = " & txtPedido.Text
           
                rdoControle.CursorLocation = adUseClient
                rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

                If rdoControle.EOF = False Then
                    'grdPrecos.TextMatrix(0, 0) = Trim(rdoControle("ModalidadeVenda"))
         
                    Select Case UCase(Trim(rdoControle("ModalidadeVenda")))
                         Case "FINANCIADO"
                              Call MontaPrecos("FI", "")
                         Case "FATURADO"
                              Call MontaPrecos("FA", "")
                         Case "A VISTA"
                              Call MontaPrecos("AV", "")
                         Case "CARTÃO"
                              Call MontaPrecos("CC", "")
                    End Select
         
                    'For Index = 1 To grdPrecos.Rows - 1
                        'If Trim(Mid(grdPrecos.TextMatrix(Index, 0), 1, 2)) = Trim(rdoControle("Parcelas")) Then
                            'grdPrecos.Row = Index
                            'grdPrecos.Enabled = False
                            'Exit For
                       ' End If
                    'Next Index
         
                End If
                rdoControle.Close
         
            'ElseIf cmdQtdeItens.Caption = 0 Then
               'grdPrecos.Enabled = True
            End If
    Else
            txtPesquisar.SetFocus
    End If
    
ElseIf KeyAscii = 27 Then

    If cmdQtdeItens.Caption > 0 Then
        cmdBotoes(0).Visible = True
        
    Else
        cmdBotoes(0).Visible = False
        
    End If

    cmdFechaPedido.Visible = False
    cmdBotoes(2).Visible = False
    cmdBotoes(12).Visible = False
    cmdBotoes(8).Visible = False
    cmdBotoes(6).Visible = False
    cmdBotoes(9).Visible = False
    cmdBotoes(10).Visible = False
    
    cmdTR.Visible = False
    cmdBotoes(7).Visible = False
    PicBanner.Visible = True
    cmbPedido.Visible = True

     If txtPesquisar.Text <> "" Then
        'cmdBotoes(0).Visible = True
        cmdBotoes(2).Visible = True
        cmdBotoes(1).Visible = True
        cmdBotoes(4).Visible = True
        cmdBotoes(11).Visible = True
        cmdBotoes(13).Visible = True
        cmdBotoes(14).Visible = True
        txtPesquisar.SelStart = 0
        txtPesquisar.SelLength = Len(txtPesquisar.Text)
        txtPesquisar.SetFocus
     Else
        Call VerificaItensVendas
        
On Error GoTo ErroDeletaNFCapa
         If auxQtdeItens = 0 Then
            adoCNLoja.BeginTrans
            SQL = "Delete NFCapa Where TipoNota = 'PD' and NumeroPed = " & txtPedido.Text
            adoCNLoja.Execute SQL
            adoCNLoja.CommitTrans
            
            sairDoSistema
            End
        Else
           cmdBotoes(2).Visible = True
           cmbPedido.Visible = True
           cmdBotoes(1).Visible = True
           cmdBotoes(4).Visible = True
           cmdBotoes(11).Visible = True
           cmdBotoes(13).Visible = True
 cmdBotoes(14).Visible = True
           txtPesquisar.SetFocus
        End If
    End If
ElseIf (KeyAscii <> 27) Or (KeyAscii <> 13) Then
     txtPesquisar.SetFocus
End If
Exit Sub
 
ErroDeletaNFCapa:
    adoCNLoja.RollbackTrans
    Exit Sub
    
End Sub

Function PesquisarProduto(ByVal wWhere As String)
    grdItensProduto.Rows = 1
      
    cmdLimpar.Caption = "Pesquisando ..."
    Screen.MousePointer = 11
            
    
    'If auxQtdeItens <= 0 Then
        'grdPrecos.TextMatrix(0, 0) = "A Vista"
    'End If
                       
SQL = ""

If Trim(GLB_Loja) = "184" Then
    SQL = "Select '' as IcmsSaida," & _
          "'' as IcmsPdv," & _
          "PRB_CodigoBarras,PR_Referencia,PR_Descricao,PR_PrecoVenda AS PR_PrecoVenda1,ES_Estoque as EL_Estoque,'B' as PR_Classe,'' as pr_CodigoProdutoNoFornecedor," & _
          "PR_Bloqueio,PR_SubstituicaoTributaria,'' as LPR_Linha,'PRODUTO SITE' as LPR_Descricao,'1' as pr_indicePreco, pr_classeFiscal,'60' as PR_ST,'N' AS PR_GarantiaEstendida ,'SITE' as FO_NOMEFANTASIA " & _
          "From ProdutoLoja, Produtobarras, svdmac.dmac.dbo.Estoque " & _
          "Where ES_Referencia=PR_Referencia and " & wWhere & " and PR_Situacao not in('E') and PRB_Referencia = PR_Referencia and ES_Loja = '184'" & _
          " " & _
          "Order By PR_CodigoFornecedor,PR_Descricao"
Else
    SQL = "Select (CASE WHEN PR_SubstituicaoTributaria = 'N' THEN PR_ICMSSaida ELSE PR_ICMSSaidaIva End) as IcmsSaida," & _
          "(CASE WHEN PR_SubstituicaoTributaria = 'N' THEN PR_IcmPdv ELSE PR_ICMSPDVSaidaIva End) as IcmsPdv," & _
          "PRB_CodigoBarras,PR_Referencia,PR_Descricao,PR_PrecoVenda1,EL_Estoque,PR_Classe,pr_CodigoProdutoNoFornecedor," & _
          "PR_Bloqueio,PR_SubstituicaoTributaria,LPR_Linha,LPR_Descricao,pr_indicePreco, pr_classeFiscal,PR_ST,PR_GarantiaEstendida ,FO_NOMEFANTASIA, " & _
          "EL_NaoComercializado, EL_NaoComercializadoCONSO " & _
          "From ProdutoLoja, Produtobarras, EstoqueLoja, LinhaProduto,fornecedor " & _
          "Where EL_Referencia=PR_Referencia and " & wWhere & " and PR_Situacao not in('E') and PRB_Referencia = PR_Referencia " & _
          "and (Case When PR_LinhaProduto IS NULL Then  '990100' Else PR_LinhaProduto End) = LPR_Linha  and pr_codigofornecedor=fo_codigofornecedor " & _
          "Order By PR_CodigoFornecedor,PR_Descricao"
End If


    
    rsPesquisaPed.CursorLocation = adUseClient
    rsPesquisaPed.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsPesquisaPed.EOF Then
       AuxProdutoExiste = True
       'grdPrecos.Redraw = True
       txtPesquisar.Text = rsPesquisaPed("PR_Descricao")
       grdItensProduto.Redraw = False
       ReferenciaPreco = rsPesquisaPed("PR_PrecoVenda1")
       

        Do While Not rsPesquisaPed.EOF

            If Trim(rsPesquisaPed("PR_Classe")) = "P" Then
                wClasseProduto = "PROMOÇÃO"
            Else
                wClasseProduto = rsPesquisaPed("PR_Classe")
            End If
       

            grdItensProduto.AddItem rsPesquisaPed("PR_Referencia") & Chr(9) _
                & rsPesquisaPed("PR_Descricao") & Chr(9) _
                & Format(rsPesquisaPed("PR_PrecoVenda1"), "0.00") & Chr(9) _
                & rsPesquisaPed("EL_Estoque") & Chr(9) _
                & "0,00" & Chr(9) _
                & rsPesquisaPed("PRB_CodigoBarras") & Chr(9) _
                & Trim(rsPesquisaPed("LPR_Linha")) & Chr(9) _
                & rsPesquisaPed("LPR_Descricao") & Chr(9) & "0" & Chr(9) _
                & wClasseProduto & Chr(9) _
                & rsPesquisaPed("PR_Bloqueio") & Chr(9) _
                & Format(rsPesquisaPed("IcmsSaida"), "0.00") & Chr(9) _
                & Format(rsPesquisaPed("IcmsPdv"), "0.00") & Chr(9) _
                & Trim(rsPesquisaPed("PR_SubstituicaoTributaria")) & Chr(9) _
                & rsPesquisaPed("pr_indicePreco") & Chr(9) _
                & rsPesquisaPed("pr_classeFiscal") & Chr(9) _
                & rsPesquisaPed("FO_NOMEFANTASIA") & Chr(9) _
                & rsPesquisaPed("PR_ST") & Chr(9) _
                & rsPesquisaPed("PR_GarantiaEstendida") & Chr(9) _
                & rsPesquisaPed("EL_NaoComercializado") & Chr(9) _
                & rsPesquisaPed("EL_NaoComercializadoCONSO")
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
        'grdPrecos.Rows = 1
        'grdPrecos.Rows = 2
        
        wProdutoNaoExiste = False
        For Index = 1 To Len(txtPesquisar.Text)
            If Mid(txtPesquisar.Text, Index, Len(txtPesquisar.Text)) = "Nenhum Registro Encontrado" Then
                wProdutoNaoExiste = True
                Exit For
            End If
        Next Index
        
        If wProdutoNaoExiste = False Then
            txtPesquisar.Text = (txtPesquisar.Text & "     " & "Nenhum Registro Encontrado")
        End If
        
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

Private Sub MontaPrecos(CodigoCrediario As String, indicePreco As String)
    

  'grdPrecos.Rows = 1
  'grdPrecos.Redraw = False

  SQL = ""
  
  If CodigoCrediario = "AV" Then
     SQL = "Select * from CondicaoPagamento " & vbNewLine _
     & "where CP_Tipo = '" & CodigoCrediario & "' and CP_Codigo = 1 and cp_id = '" & wIndicePreco & "'"
  Else
     SQL = "Select * from CondicaoPagamento " & vbNewLine _
     & "where CP_Tipo = '" & CodigoCrediario & "' and cp_id = '" & wIndicePreco & "' Order By CP_Codigo"
  End If
   
  rsCondicaoFaturado.CursorLocation = adUseClient
  rsCondicaoFaturado.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  If CodigoCrediario = "FA" Then
      txtCondicaoFaturado.Visible = True
      mskDatafaturado.Visible = True
  Else
      txtCondicaoFaturado.Visible = False
      mskDatafaturado.Visible = False
  End If
 

     Do While Not rsCondicaoFaturado.EOF

        wValorTotalCalculado = Format((wValorVenda * rsCondicaoFaturado("CP_Coeficiente")), "###,###,###,##0.00")
        wPrecoCalculado = Format((wValorTotalCalculado / rsCondicaoFaturado("CP_Parcelas")), "###,###,###,##0.00")
        NomeParcela = rsCondicaoFaturado("CP_Condicao") & " " & Format(wPrecoCalculado, "###,###,###,##0.00")

        'If CodigoCrediario = "FA" Then
           '''grdPrecos.AddItem rsCondicaoFaturado("CP_Codigo") & Chr(9) _
               '''& Format(wValorTotalCalculado, "###,###,###,##0.00") & Chr(9) _
               '''& rsCondicaoFaturado("CP_Condicao")
        'ElseIf CodigoCrediario = "AV" Then
           '''grdPrecos.AddItem rsCondicaoFaturado("CP_Condicao") & Chr(9) _
               '''& Format(wValorTotalCalculado, "###,###,###,##0.00") & Chr(9) _
               '''& rsCondicaoFaturado("CP_Codigo")
        'Else
            '''grdPrecos.AddItem NomeParcela & Chr(9) _
               '''& Format(wValorTotalCalculado, "###,###,###,##0.00") & Chr(9) _
               '''& rsCondicaoFaturado("CP_Codigo")
        'End If
        rsCondicaoFaturado.MoveNext
     Loop

        '''grdPrecos.Col = 0
        '''grdPrecos.ColSel = 1
        '''grdPrecos.Redraw = True
        '''grdPrecos.TopRow = wGuardaLinha
        rsCondicaoFaturado.Close
End Sub

Private Sub GravaItens()

Dim wVlunit2 As Double
Dim wVltotitem As Double
Dim wIcms As Double
Dim wDesconto As Double

'On Error GoTo erronaInclusao

SQL = ""
SQL = "Select Referencia, Qtde From NFItens Where NumeroPed = " & txtPedido.Text & " and " _
      & "Referencia = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "' and TipoNota = 'PD'"

rsItensVenda.CursorLocation = adUseClient
rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    If rsItensVenda.EOF = True Then

      SQL = "Select max(item) as MaxItens from NFItens Where NumeroPed = " & txtPedido.Text
      rsComplementoVenda.CursorLocation = adUseClient
      rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

      If Not rsComplementoVenda.EOF Then
        auxItens = IIf(IsNull(rsComplementoVenda("MaxItens")), 0, rsComplementoVenda("MaxItens"))
      Else
        auxItens = 0
      End If
      rsComplementoVenda.Close

      

       auxItens = (auxItens + 1)
       wPreco = Format(grdItensProduto.TextMatrix(wGuardaLinha, 2), "#0.00")
       'Verificar Valores
       wVltotitem = Format((wPreco * txtQuantidade.Text), "##0.00")
       wDesconto = Format(0, "##0.00")
       'wIcms = Format(grdDadosProduto.TextMatrix(1, 5), "#0.00")
'----'
'       adoCNLoja.BeginTrans
       Screen.MousePointer = vbHourglass
'**************************** Insert na Tabela NFItens
        SQL = "Insert into NFItens (NF,NUMEROPED,SERIE,DATAEMI,REFERENCIA,QTDE,VLUNIT, " _
            & "VLTOTITEM,ICMS,DESCONTO,PLISTA,LOJAORIGEM,TIPONOTA,Item,SITUACAOPROCESSO,DATAPROCESSO,ICMSAplicado) Values (0," _
            & txtPedido.Text & ",'','" & Format(Date, "yyyy/mm/dd") & "','" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "'," _
            & txtQuantidade.Text & "," & ConverteVirgula(wPreco) & "," & ConverteVirgula(wVltotitem) & "," _
            & ConverteVirgula(wIcms) & "," & ConverteVirgula(wDesconto) & "," & ConverteVirgula(wPreco) & ",'" & Trim(wLoja) & "'," _
            & "'PD'," & auxItens & ",'A','" & Format(Date, "yyyy/mm/dd") & "',0)"
'''        SQL = "Insert into NFItens (NF,NUMEROPED,SERIE,DATAEMI,REFERENCIA,QTDE,VLUNIT, " _
'''            & "VLTOTITEM,icmpdv,ICMS,DESCONTO,PLISTA,LOJAORIGEM,TIPONOTA,Item,SITUACAOPROCESSO,DATAPROCESSO,ICMSAplicado,tipomovimentacao) Values (0," _
'''            & txtpedido.Text & ",'','" & Format(Date, "yyyy/mm/dd") & "','" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "'," _
'''            & txtQuantidade.Text & "," & ConverteVirgula(wPreco) & "," & ConverteVirgula(wVltotitem) & "," _
'''            & ConverteVirgula(wIcms) & "," & ConverteVirgula(wIcms) & "," & ConverteVirgula(wDesconto) & "," & ConverteVirgula(wPreco) & ",'" & Trim(wLoja) & "'," _
'''            & "'PD'," & auxItens & ",'A','" & Format(Date, "yyyy/mm/dd") & "',0,11)"
                
        adoCNLoja.Execute SQL
        Screen.MousePointer = vbNormal
'        adoCNLoja.CommitTrans
    Else
        auxItens = 0
        If MsgBox("Referência já cadastrada. Deseja somar a quantidade?", vbQuestion + vbYesNo, "Pedido") = vbYes Then
           SQL = ""
           SQL = "UPDATE NFItens set Qtde = (Qtde + " & txtQuantidade.Text & "), VLTOTITEM = ((vlunit - desconto) * (" & rsItensVenda("Qtde") & " + " & txtQuantidade.Text & ")) " _
                 & "Where NumeroPed = " & txtPedido.Text & " and Referencia = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "' and TipoNota = 'PD'"
'           adoCNLoja.BeginTrans
           adoCNLoja.Execute SQL
'           adoCNLoja.CommitTrans
        Else
           grdItensProduto.SetFocus
        End If
    End If
        
rsItensVenda.Close
Exit Sub
        
        
erronaInclusao:
MsgBox "Erro na Inclusão de itens " & Err.description, vbCritical, "Aviso"

adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal

Exit Sub

End Sub

Private Sub txtQuantidade_Change()
    If IsNumeric(txtQuantidade.Text) = False Then
        txtQuantidade.Text = ""
        txtQuantidade.SelStart = 0
        txtQuantidade.SelLength = Len(txtQuantidade.Text)
        txtQuantidade.SetFocus
    ElseIf txtQuantidade.Text <= 0 Then
        txtQuantidade.Text = ""
        txtQuantidade.SelStart = 0
        txtQuantidade.SelLength = Len(txtQuantidade.Text)
        txtQuantidade.SetFocus
    End If

        
End Sub



Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 46 Then
      txtQuantidade.Text = 0
      txtQuantidade.SelStart = 0
      txtQuantidade.SelLength = Len(txtQuantidade.Text)
      txtQuantidade.SetFocus
      Exit Sub
   End If
   
   If KeyAscii = 44 Then
      txtQuantidade.Text = 0
      txtQuantidade.SelStart = 0
      txtQuantidade.SelLength = Len(txtQuantidade.Text)
      txtQuantidade.SetFocus
      Exit Sub
   End If
   
   If KeyAscii = 13 Then
      If cmdFechaPedido.Visible = True Then
        If frmPedido.cmdBotoes(9).Visible = False Or _
           frmPedido.cmdBotoes(8).Visible = False Or _
           frmPedido.cmdBotoes(10).Visible = False Or _
           frmPedido.cmdBotoes(12).Visible = False Or _
           frmPedido.cmdTR.Visible = False Or _
           frmPedido.cmdBotoes(6).Visible = False Or _
           frmPedido.cmdBotoes(7).Visible = False Then
           txtQuantidade.SelLength = Len(txtQuantidade.Text)
           txtQuantidade.SetFocus
           cmbPedido.Visible = False
           cmdBotoes(1).Visible = False
           cmdBotoes(4).Visible = False
           cmdBotoes(11).Visible = False
           cmdBotoes(13).Visible = False
 cmdBotoes(14).Visible = False
           cmdBotoes(2).Visible = False
           
           
           Exit Sub
    
        End If
      Else
          If txtQuantidade.Text <> "" Then
             cmdBotoes(2).Visible = True
             cmbPedido.Visible = True
             cmdBotoes(1).Visible = True
             cmdBotoes(4).Visible = True
             cmdBotoes(11).Visible = True
             cmdBotoes(13).Visible = True
 cmdBotoes(14).Visible = True
             cmdBotoes(0).Visible = True
             
          End If
      End If
   
   End If
   
      
     If KeyAscii = 13 Then
         If grdItensProduto.Row <> 0 Then
             wGuardaLinha = grdItensProduto.Row
         End If
     If IsNumeric(Trim(txtQuantidade.Text)) Then

        Call GravaItens
        Call SomaItensVenda
        
        'SQL = ""
        SQL = "Update NFCapa Set ModalidadeVenda = '" & "A Vista" & "'" & _
              " Where NumeroPed = " & (txtPedido.Text)
        adoCNLoja.Execute SQL
        
        
        'SQL = ""
         SQL = "Update NFCapa set condpag = '1' where NumeroPed = " & txtPedido.Text
                adoCNLoja.Execute SQL
          
        SQL = ""
          
        'If grdPrecos.TextMatrix(0, 0) = "Faturado" Then
'            SQL = "Update NFCapa set condpag = '" & grdPrecos.TextMatrix(wGuardaLinha, 0) & _
'                  "' where NumeroPed = " & txtPedido.Text
'                  adoCNLoja.Execute SQL
'            If grdPrecos.TextMatrix(wGuardaLinha, 0) = "85" Then
'                wGuardaPagamento = " - " & mskDatafaturado.Text
'            Else
'                wGuardaPagamento = " - " & txtCondicaoFaturado.Text
'            End If
            
        'Else
'            If grdPrecos.TextMatrix(0, 0) = "Financiado" Or grdPrecos.TextMatrix(0, 0) = "Cartão" Then
'                SQL = "Update NFCapa set condpag = '" & grdPrecos.TextMatrix(wGuardaLinha, 2) & _
'                "' where NumeroPed = " & txtPedido.Text
'                adoCNLoja.Execute SQL
'                wGuardaPagamento = " - " & grdPrecos.TextMatrix(wGuardaLinha, 0)
'            Else
'                wGuardaPagamento = " "
'            End If
        'End If
           
        SQL = "Update NFCapa Set Parcelas = 0  Where ModalidadeVenda = 'A Vista' and NumeroPed = " & Val(txtPedido.Text)
        adoCNLoja.Execute SQL
        
        'grdPrecos.Enabled = False
        txtQuantidade.Text = ""
        txtQuantidade.Enabled = False
        txtPesquisar.Enabled = True
        txtPesquisar.SetFocus
        AuxProdutoExiste = False
     Else
        txtQuantidade.SelLength = Len(txtQuantidade.Text)
        txtQuantidade.SetFocus
     End If
   ElseIf KeyAscii = 27 Then
      txtQuantidade.Text = ""
      txtQuantidade.Enabled = False
      txtPesquisar.Enabled = True
      txtPesquisar.SetFocus
      Call VerificaItensVendas
      If auxItens = 0 Then
         'grdPrecos.Enabled = True
         'grdPrecos.SetFocus
      Else
         grdItensProduto.SetFocus
      End If
   End If
End Sub


Private Sub LerControleSistema()

  SQL = "Select CTS_CaminhoWeb2,CTS_Loja,CTS_CaminhoWeb1,CTS_CaminhoBanner,CTS_CaminhoWeb2,CTS_LogoPedido from ControleSistema"
  rdoControle.CursorLocation = adUseClient
  rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
  If rdoControle.EOF Then
     MsgBox "Problemas com o Sistema (Controle Sistema) entrar em contato com o TI", vbCritical, "Atenção"
     rdoControle.Close
     End
  Else
     wLoja = rdoControle("CTS_Loja")
     wBanner = Trim(rdoControle("CTS_CaminhoWeb1"))
     wBanner2 = Trim(rdoControle("CTS_CaminhoBanner"))
     wFichaTec = Trim(rdoControle("CTS_CaminhoWeb2"))
     GLB_logoPedido = Trim(rdoControle("CTS_LogoPedido"))
     rdoControle.Close
  End If
  
End Sub

Private Sub LerControleCaixa()
SQL = "Select * from ControleCaixa where CTR_SituacaoCaixa='A' and ctr_dataInicial >= '" & Format(Date, "yyyy/mm/dd") & "' "
rdoControle.CursorLocation = adUseClient
rdoControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

If rdoControle.EOF Then
   MsgBox "Caixa fechado.", vbCritical, "Atenção"
   rdoControle.Close
   End
Else
wProtocolo = rdoControle("CTR_Protocolo")
rdoControle.Close
End If

End Sub
Private Sub VerificaItensVendas()
'********************* NFItens
  SQL = "Select Count(*) as NroItens from NFItens Where NumeroPed = " & txtPedido.Text

  rsItensVenda.CursorLocation = adUseClient
  rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  auxQtdeItens = rsItensVenda("NroItens")
  rsItensVenda.Close
  

End Sub

Private Sub txtVendedor_Change()
    If txtVendedor.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtVendedor.Text = ""
        txtVendedor.SetFocus
        Exit Sub
    End If
End Sub


Private Sub txtVendedor_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   sairDoSistema
   End
End If

 If KeyAscii = 46 Then
      txtVendedor.Text = 0
      txtVendedor.SelStart = 0
      txtVendedor.SelLength = Len(txtVendedor.Text)
      txtVendedor.SetFocus
      Exit Sub
   End If
   
   If KeyAscii = 44 Then
      txtVendedor.Text = 0
      txtVendedor.SelStart = 0
      txtVendedor.SelLength = Len(txtVendedor.Text)
      txtVendedor.SetFocus
      Exit Sub
   End If

If KeyAscii = vbKeyReturn Then
  If IsNumeric(txtVendedor.Text) = False Then
     txtVendedor.Text = ""
  Else
  
  SQL = ""
    SQL = "Select numeroped,vendedor From nfcapa Where numeroped = " & txtPedido.Text
    rsVendedor.CursorLocation = adUseClient
    
    
    rsVendedor.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not rsVendedor.EOF Then
            If rsVendedor("Vendedor") <> txtVendedor.Text Then
                MsgBox "Numero do Pedido ou Vendedor incorreto"
                rsVendedor.Close
                Call LimpaForm
                Exit Sub
            Else
                txtPedido.Enabled = True
                txtVendedor.Width = 8000
                fradados.Enabled = False
                fradados.Width = 12640
                cmdBotoes(3).Visible = True
                cmdBotoes(11).Visible = True
                cmdBotoes(13).Visible = True
                cmdBotoes(14).Visible = True
                cmdBotoes(4).Visible = True
                cmdBotoes(5).Visible = True
                grdItensProduto.Enabled = True
                'grdPrecos.Enabled = True
                grdDadosProduto.Enabled = True
                txtQuantidade.Enabled = True
                txtPesquisar.Enabled = True
                txtPesquisar.SetFocus
            End If
    Else
        rsVendedor.Close
    
        SQL = "Select VE_Codigo, VE_Nome From Vende WHERE VE_Codigo = " & txtVendedor.Text
              rsVendedor.CursorLocation = adUseClient
               
        rsVendedor.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        If Not rsVendedor.EOF Then
            
            txtVendedor.MaxLength = 0
            wVendedor = rsVendedor("VE_Codigo") & " - " & rsVendedor("VE_Nome")
            auxVEN_CodigoVendedor = rsVendedor("VE_Codigo")
            txtVendedor.Text = wVendedor
            
            If Trim(txtVendedor.Text) = "" Then
               MsgBox wVendedor
               txtVendedor.SetFocus
               Exit Sub
            End If
            
            txtPedido.Enabled = True
            txtVendedor.Width = 7820
            fradados.Enabled = False
            fradados.Width = 8685
            Call CriaCapaPedido(txtPedido.Text)
            cmdBotoes(3).Visible = True
            cmdBotoes(11).Visible = True
            cmdBotoes(13).Visible = True
            cmdBotoes(14).Visible = True
            cmdBotoes(5).Visible = True
            cmdBotoes(2).Visible = True
            cmdBotoes(4).Visible = True
            txtPesquisar.Enabled = True
            txtPesquisar.SetFocus
            
            SQL = ""
            SQL = "Update LembreMe set LEM_situacao = 'O' from LembreMe, estoqueloja " & _
                  "where el_referencia = Lem_referencia and el_estoque > 0 and LEM_Situacao = 'E'"
            adoCNLoja.Execute SQL
            
            
            SQL = ""
            SQL = "Select LEM_Referencia from LembreMe " & _
                  "where lem_situacao = 'O' and lem_vendedor = '" & Mid(frmPedido.txtVendedor.Text, 1, 3) & _
                  "' Order by LEM_Data,LEM_Referencia"
            rsLembrete.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

            If Not rsLembrete.EOF Then
               rsLembrete.Close
               frmLembrete.Show 1
               frmLembrete.ZOrder
            Else
               rsLembrete.Close
            End If
               
         Else
            cmdLimpar.Caption = "Vendedor " & Mid(frmPedido.txtVendedor.Text, 1, 3) & " não Cadastrado"
            MsgBox "Vendedor " & Mid(frmPedido.txtVendedor.Text, 1, 3) & " não Cadastrado.", vbExclamation, "Atenção"
            txtVendedor.SetFocus
        End If
    End If
       rsVendedor.Close
  End If
End If
End Sub
Private Sub SomaItensVenda()
'******************* NFItens
  If rsItensVenda.State = 1 Then rsItensVenda.Close

  SQL = "Select TipoNota, sum(VLTOTITEM) as TotalVenda," _
        & "Count(*) as TotalItens, Max(Item) as UltimoReg From NFItens Where NumeroPed = " & txtPedido.Text & " and " _
        & "TipoNota = 'PD' Group By TipoNota"
      
      
  rsItensVenda.CursorLocation = adUseClient
  rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsItensVenda.EOF Then
     cmdTotalPedido.Caption = Format(rsItensVenda("TotalVenda") + GBL_Frete, "###,###,##0.00")
     'Do While Len(cmdTotalPedido.Caption) <= 12
         'cmdTotalPedido.Caption = cmdTotalPedido.Caption '+ " "
     'Loop
     auxQtdeItens = rsItensVenda("TotalItens")
     cmdQtdeItens.Caption = rsItensVenda("TotalItens")
     
     'Do While Len(cmdQtdeItens.Caption) <= 5
         'cmdQtdeItens.Caption = cmdQtdeItens.Caption + " "
     'Loop
     
     
     
'     auxItens = rsItensVenda("UltimoReg")
'     auxVendedordoPedido = rsItensVenda("Vendedor")
     cmdLimpar.Caption = ""
        
  Else
     txtPedido.Text = ""
     txtPedido.SetFocus
     cmdLimpar.Caption = "Pedido não cadastrado ou encerrado"
     rsItensVenda.Close
     Exit Sub
  End If
  rsItensVenda.Close
End Sub

Private Sub LerVendedordoPedido()
SQL = "Select * From Vende Where VE_Codigo = " & auxVendedordoPedido
      rsVendedor.CursorLocation = adUseClient
      rsVendedor.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        If Not rsVendedor.EOF Then
            auxVEN_CodigoVendedor = rsVendedor("VE_Codigo")
            txtVendedor.Width = 5020
            txtVendedor.MaxLength = 0
            txtVendedor.Text = rsVendedor("VE_Codigo") & " - " & rsVendedor("VE_Nome")
            txtVendedor.Enabled = True
            txtPesquisar.Enabled = True
            txtPesquisar.SetFocus
           
        End If
      rsVendedor.Close
End Sub

Private Sub txtVendedor_LostFocus()

    If GetAsyncKeyState(vbKeyTab) <> 0 Then
        If Not txtVendedor.Width = 7740 Then
            txtVendedor.Text = ""
            txtVendedor.SelStart = 0
            txtVendedor.SelLength = Len(txtVendedor.Text)
            fradados.Enabled = True
            txtVendedor.Enabled = True
            txtVendedor.SetFocus
        End If
        Exit Sub
    End If
    
    If Not txtVendedor.Text = "" Then
        cmdBotoes(4).Visible = True
        cmdBotoes(1).Visible = True
        cmdBotoes(5).Visible = True
        cmdBotoes(3).Visible = True
        cmdBotoes(11).Visible = True
        cmdBotoes(13).Visible = True
        cmdBotoes(14).Visible = True
        cmdBotoes(2).Visible = True
    End If

End Sub

Public Function Numeros(ByVal Texto As String) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim CharLido As String * 1
    Dim retorno As String
    
    Maximo = Len(Texto)
    retorno = ""
    For Char = 1 To Maximo Step 1
        CharLido = Mid(Texto, Char, 1)
        If IsNumeric(CharLido) Then
            retorno = retorno & CharLido
        End If
    Next Char
    Texto = retorno
    Numeros = Texto

End Function
Private Sub RemoveMenus()
  Dim hMenu As Long
  hMenu = GetSystemMenu(Hwnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Function CriaCapaPedido(ByVal NumeroPedido As Double)
      
    SQL = ""
    SQL = "Select count(referencia) as NumeroItem from NFItens " _
          & "where NumeroPed=" & NumeroPedido & ""
          
          rsComplementoVenda.CursorLocation = adUseClient
          rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    SQL = ""
    SQL = "Insert Into NFCapa(NUMEROPED,DATAEMI,LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA,cliente,volume,pesoLq,pgentra,desconto,fretecobr," _
        & "VendedorLojaVenda, LojaVenda,TM,qtditem, OutraLoja, OutroVend, baseicms, situacaoprocesso,dataprocesso) " _
        & "Values (" & NumeroPedido & ",'" & Format(Date, "yyyy/mm/dd") & "', " _
        & "'" & wLoja & "','PD'," & auxVEN_CodigoVendedor & ", " _
        & "'" & Format(Date, "yyyy/mm/dd") & "', '" & Format(Time, "hh:mm:ss") & "','999999',1,1,0,0,0, " _
        & auxVEN_CodigoVendedor & ", '" & wLoja & "',0," _
        & rsComplementoVenda("Numeroitem") & "," & wLoja & "," & auxVEN_CodigoVendedor & ",0,'A','" _
        & Format(Date, "yyyy/mm/dd") & "')"
        adoCNLoja.Execute (SQL)
     
     rsComplementoVenda.Close
     
     
End Function

Private Sub wbFichaTecnica_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Set objDoc = wbFichaTecnica.Document
    Set objWind = objDoc.parentWindow
End Sub

Private Sub wbInternet_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Set objDoc = wbFichaTecnica.Document
   ' Set objWind = objDoc.parentWindow
End Sub

Private Sub objWind_onerror(ByVal description As String, ByVal URL As String, ByVal line As Long)
    Set objEvent = objWind.event
    objEvent.returnValue = True
End Sub

Private Sub VerificaInternet()
Dim wVerificaInternet As String
       

  
End Sub

 Sub FechaPedido()

    Dim rsControle As New ADODB.Recordset
    

 auxItens = 0
 wCodigo = 1
 wSequencia = 1
 wValorDados = "V"
  
' On Error GoTo erronoUpdate
    Screen.MousePointer = vbHourglass
    
 
    SQL = ""
    SQL = "Select Referencia From NFItens Where NumeroPed = " & frmPedido.txtPedido.Text
 
    rsItensVenda.CursorLocation = adUseClient
    rsItensVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rsItensVenda.EOF = False Then
       adoCNLoja.BeginTrans
              
'************************ Verificando se Nota é Eletrônica


    SQL = "select cliente from nfcapa where numeroped = " & frmPedido.txtPedido.Text
    
    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If RTrim(LTrim(rsComplementoVenda("cliente"))) <> "999999" Then
    
        If validaDadosCliente(rsComplementoVenda("cliente")) = False Then
            MsgBox "Há erro(s) no cadastro desse cliente que impede o finalizamento desse pedido", vbExclamation, "Cliente"
            rsComplementoVenda.Close
            rsItensVenda.Close
            adoCNLoja.CommitTrans
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        rsComplementoVenda.Close

        
        SQL = "select CTS_SerieNota from ControleSistema"
        rsControle.CursorLocation = adUseClient
        rsControle.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If rsControle("CTS_SerieNota") = "NE" Then
            SQL = "Update NfCapa set Serie = 'NE' where NumeroPed = " & frmPedido.txtPedido.Text
            adoCNLoja.Execute (SQL)
        Else
            SQL = "select ce_Estado,ce_tipopessoa,cliente from fin_cliente,nfcapa where ce_CodigoCliente = Cliente and " & _
            "NumeroPed = " & frmPedido.txtPedido.Text
            rsComplementoVenda.CursorLocation = adUseClient
            rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            If RTrim(LTrim(rsComplementoVenda("ce_Estado"))) <> "SP" _
                 Or RTrim(LTrim(rsComplementoVenda("ce_TipoPessoa"))) = "O" Then
                 
                   MsgBox "ESTE PEDIDO IRÁ GERAR UMA NOTA FISCAL ELETRÔNICA, AVISE O CLIENTE.", vbInformation, "Atenção"
                   
                   SQL = "Update NfCapa set Serie = 'NE' where NumeroPed = " & frmPedido.txtPedido.Text
                   adoCNLoja.Execute (SQL)
                   
            End If
            'rsComplementoVenda.Close
        End If
        
        rsControle.Close
        
    End If
    If rsComplementoVenda.State <> 0 Then rsComplementoVenda.Close


'************************ Gravando Valores NFCapa
       SQL = ""
       SQL = "Exec SP_Totaliza_Capa_Nota_Fiscal_Loja " & frmPedido.txtPedido.Text
       adoCNLoja.Execute SQL
       
                  
       SQL = ""
       SQL = "Select count(referencia) as NumeroItem from NFItens " _
           & "where NumeroPed=" & frmPedido.txtPedido.Text & ""
          
            rsComplementoVenda.CursorLocation = adUseClient
            rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
       SQL = ""
       SQL = "Update NFCapa set TipoNota = 'PA', qtditem = " & rsComplementoVenda("NumeroItem") & "" _
             & " Where NumeroPed = " & frmPedido.txtPedido.Text
       adoCNLoja.Execute SQL
       
       rsComplementoVenda.Close
       
       
'************************ Gravando TipoNota NFItens
       SQL = "Update NFItens Set TipoNota = 'PA' Where NumeroPed = " & frmPedido.txtPedido.Text
       
       adoCNLoja.Execute SQL
       adoCNLoja.CommitTrans
       
'************************ Verifica se é entrada

       'If Mid(wVendedor, 1, 3) = "790" Then
       '     SQL = "exec SP_EST_Outras_Entradas_Estoque_Loja " & frmPedido.txtPedido.Text
       '     adoCNLoja.Execute SQL
       'End If

       Call LimpaForm

     Else
        MsgBox "Nenhum produto encontrado. Você não pode finalizar o pedido.", vbExclamation, "Atenção"
        txtPesquisar.SetFocus
        txtPesquisar.SelStart = 0
        txtPesquisar.SelLength = Len(txtPesquisar.Text)
     End If
     rsItensVenda.Close
     
     
     
     Screen.MousePointer = vbNormal
     Exit Sub
       
'erronoUpdate:
 '   MsgBox "Erro na atualizazão da situação do pedido " & Err.description, vbCritical, "Aviso"
 '   adoCNLoja.RollbackTrans
 '   Screen.MousePointer = vbNormal
       
End Sub

Sub VerificaBannerReferencia()

On Error GoTo trata_erro
  PicBanner.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerAcessorios\" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & ".swf")
  wBannerReferencia = True
    
  Exit Sub
  
trata_erro:
  

  If Err.Number = 53 Then
     wBannerReferencia = False
  Else
     MsgBox "Ocorreu o erro : " & Err.Number & " - " & Err.description
  End If

End Sub

Sub VerificaBannerLinha()
On Error GoTo trata_erro
  PicBanner.Picture = LoadPicture("C:\Sistemas\DMAC Venda\Imagens\BannerAcessorios\" & grdDadosProduto.TextMatrix(1, 0) & ".swf")
  wBannerLinha = True
    
  Exit Sub
  
trata_erro:
  

  If Err.Number = 53 Then
     wBannerLinha = False
  Else
     MsgBox "Ocorreu o erro : " & Err.Number & " - " & Err.description
  End If
End Sub

Private Function itemComGarantiaEstendida(ByRef NumeroPedido As String) As Boolean
    itemComGarantiaEstendida = True
End Function


Private Sub carregaProdutoGarantia()
        Dim rsProdGarantiaEstendida As New ADODB.Recordset
        
        SQL = "select count(*) itensGarantia " & _
        "from produtoLoja as p, nfitens as i, nfcapa as c " & _
        "where i.numeroPed = " & frmPedido.txtPedido & " and  " & _
        "p.pr_referencia = i.referencia and " & _
        "p.pr_garantiaEstendida = 'S' and i.numeroPed = c.numeroPed and " & _
        "c.vendedor not in (999,888,777)"
        
        rsProdGarantiaEstendida.CursorLocation = adUseClient
        rsProdGarantiaEstendida.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
            If Val(rsProdGarantiaEstendida("itensGarantia")) > 0 Then
                frmGarantiaEstendida.Show 1
                'frmGarantiaEstendida.ZOrder
            End If
        rsProdGarantiaEstendida.Close
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim MSG As Long
  tempoRestante = "00:00:10"
  
If (Button + Shift + Y) = 0 Then
    MSG = X / Screen.TwipsPerPixelX
    Select Case MSG
        'Case WM_LBUTTONDOWN
            'Me.WindowState = 0
            'Call SysTray(TrayDelete, Me.Hwnd, Me.Caption, Me.Icon)
            'Me.SetFocus
            'frmControleNotaFiscal.Visible = True
            'frmLogNotaFiscal.Visible = True
            'Call SysTray(TrayDelete, Me.Hwnd, Me.Caption, Me.Icon)
            'frmPedido.Visible = True
            'frmPedido.Visible = True
            'frmPedido.SetFocus
''        Case WM_LBUTTONDBLCLK
''            'Coloque aqui a rotina a ser executada
''            'quando ocorrer um duplo clique com o
''            'botão esquerdo no icon do System Tray.
''            'Neste exemplo, a janela será restaurada
''            'e o ícone retirado so System Tray.
''
''        Call criaIconeBarra(TrayDelete, Me.Hwnd, Me.Caption, Me.Icon)
''            Me.SetFocus
''        Case WM_RBUTTONDOWN
''            'Coloque aqui a rotina a ser executada
''            'quando ocorrer um clique com o botão
''            'direito do rato no icon do System Tray.
''        Case WM_RBUTTONDBLCLK
''            'Coloque aqui a rotina a ser executada
''            'quando ocorrer um duplo clique com o
''            'botão direito do rato no icon do System
''            'Tray.
    End Select
End If
'Se você precisar colocar algum outro código neste
'evento, pode coloca-lo aqui sem maiores problemas.

End Sub

Private Sub WebBrowser1_GotFocus()
    Dim SQL As String
    Dim rsBanner As New ADODB.Recordset
    
    picLimitadorBanner.Height = 7850
    
    SQL = "select CTS_CaminhoWeb2 from ControleSistema"
    
    rsBanner.CursorLocation = adUseClient
    rsBanner.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not rsBanner.EOF Then
        wBanner = rsBanner("CTS_CaminhoWeb2")
    End If
    rsBanner.Close
    WebBrowser1.Navigate (wBanner)
    
End Sub

Private Sub WebBrowser1_LostFocus()
    picLimitadorBanner.Height = 2025
    WebBrowser1.Navigate (wBanner2)
    tempoRestante = "00:00:10"
End Sub

Private Sub webInternet3_Click()
    If txtPedido.Enabled And txtPedido.Visible Then
        txtPedido.SetFocus
    ElseIf txtVendedor.Enabled And txtVendedor.Visible Then
        txtVendedor.SetFocus
    ElseIf txtPesquisar.Enabled And txtPesquisar.Visible Then
        txtPesquisar.SetFocus
    End If
End Sub
