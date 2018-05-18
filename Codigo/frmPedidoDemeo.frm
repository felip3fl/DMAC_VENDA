VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Begin VB.Form frmPedido 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   15300
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPedido.frx":0000
   Picture         =   "frmPedido.frx":08CA
   ScaleHeight     =   10950
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin Project1.chameleonButton cmbEMKT 
      Height          =   315
      Left            =   1515
      TabIndex        =   40
      Top             =   1785
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   556
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":1F9ED
      PICN            =   "frmPedido.frx":1FA09
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmbAgenda 
      Height          =   330
      Left            =   1125
      TabIndex        =   39
      Top             =   1785
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   582
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":20BBD
      PICN            =   "frmPedido.frx":20BD9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmbCliente 
      Height          =   315
      Left            =   720
      TabIndex        =   38
      Top             =   1785
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":21089
      PICN            =   "frmPedido.frx":210A5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmbPedido 
      Height          =   450
      Left            =   11820
      TabIndex        =   34
      Top             =   2190
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   794
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":215BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdItens 
      Height          =   375
      Left            =   12660
      TabIndex        =   33
      Top             =   2235
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   661
      BTYPE           =   11
      TX              =   "Item."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmPedido.frx":215DA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdValorTotal 
      Height          =   435
      Index           =   0
      Left            =   13320
      TabIndex        =   32
      Top             =   2205
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      BTYPE           =   11
      TX              =   "R$"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedido.frx":215F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdTotalPedido 
      Height          =   390
      Left            =   13695
      TabIndex        =   31
      Top             =   2220
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   688
      BTYPE           =   11
      TX              =   "0,00"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmPedido.frx":21612
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdQtdeItens 
      Height          =   330
      Left            =   12300
      TabIndex        =   30
      Top             =   2265
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      BTYPE           =   11
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmPedido.frx":2162E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash WebInternet1 
      Height          =   2115
      Left            =   9540
      TabIndex        =   25
      Top             =   2640
      Width           =   5715
      _cx             =   10081
      _cy             =   3731
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash WebInternet2 
      Height          =   2115
      Left            =   90
      TabIndex        =   23
      Top             =   2640
      Width           =   9525
      _cx             =   16801
      _cy             =   3731
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.PictureBox PicBanner 
      Height          =   5535
      Left            =   9870
      Picture         =   "frmPedido.frx":2164A
      ScaleHeight     =   5475
      ScaleWidth      =   6420
      TabIndex        =   22
      Top             =   5010
      Width           =   6480
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   65535
      Left            =   510
      Top             =   11940
   End
   Begin VB.PictureBox picQuadroGeral 
      BackColor       =   &H00B63C18&
      Height          =   7515
      Left            =   90
      ScaleHeight     =   7455
      ScaleWidth      =   15105
      TabIndex        =   14
      Tag             =   "&H00AE7411&"
      Top             =   3030
      Width           =   15165
      Begin VB.Frame fraCondicao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   6345
         TabIndex        =   26
         Top             =   5805
         Width           =   2340
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   -2055
            TabIndex        =   28
            Top             =   -30
            Width           =   2145
         End
         Begin VB.TextBox txtCondicaoFaturado 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   90
            TabIndex        =   27
            Top             =   0
            Width           =   2310
         End
         Begin MSMask.MaskEdBox mskDatafaturado 
            Height          =   345
            Left            =   90
            TabIndex        =   29
            Top             =   0
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   609
            _Version        =   393216
            BorderStyle     =   0
            ForeColor       =   12582912
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
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   7905
         TabIndex        =   4
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox txtPesquisar 
         BackColor       =   &H00FFFFFF&
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
         Left            =   15
         TabIndex        =   2
         Top             =   1950
         Width           =   7875
      End
      Begin VB.PictureBox fraMenu 
         BackColor       =   &H00B63C18&
         Height          =   5490
         Left            =   8685
         ScaleHeight     =   5430
         ScaleWidth      =   3000
         TabIndex        =   17
         Top             =   1950
         Visible         =   0   'False
         Width           =   3060
         Begin VB.CommandButton CmdCarimbo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F 9 - Car&imbo"
            Height          =   495
            Left            =   15
            TabIndex        =   13
            Top             =   4110
            Width           =   1480
         End
         Begin VB.CommandButton cmdLimpar 
            Caption         =   "CmdLimpar"
            Height          =   420
            Left            =   15
            TabIndex        =   18
            ToolTipText     =   "Clique aqui p/ voltar processo Ex: Tirar o desconto, Frete, etc.. "
            Top             =   4740
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F 12 - &Consulta"
            Height          =   495
            Left            =   15
            TabIndex        =   6
            Top             =   540
            Width           =   1480
         End
         Begin VB.CommandButton cmdTR 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F 7 - &TR"
            Height          =   495
            Left            =   15
            TabIndex        =   12
            Top             =   3600
            Width           =   1480
         End
         Begin VB.CommandButton cmdVendaDistancia 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F6-&Vda Distância"
            Height          =   495
            Left            =   15
            TabIndex        =   9
            Top             =   2070
            Width           =   1480
         End
         Begin VB.CommandButton cmdTransferencia 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F5-Tr&ansferência"
            Height          =   495
            Left            =   15
            TabIndex        =   8
            Top             =   1560
            Width           =   1480
         End
         Begin VB.CommandButton cmdFechaPedido 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F1- &Fechar"
            Height          =   495
            Left            =   15
            TabIndex        =   5
            Top             =   30
            Width           =   1480
         End
         Begin VB.CommandButton cmdDesconto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F 2 - &Desconto"
            Height          =   495
            Left            =   15
            TabIndex        =   10
            Top             =   2580
            Width           =   1480
         End
         Begin VB.CommandButton cmdFrete 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F 4 - F&rete"
            Height          =   495
            Left            =   15
            TabIndex        =   11
            Top             =   3090
            Width           =   1480
         End
         Begin VB.CommandButton cmdCotacao 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F 11 - C&otação"
            Height          =   495
            Left            =   15
            TabIndex        =   7
            Top             =   1050
            Width           =   1480
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00B63C18&
         Height          =   900
         Left            =   15
         ScaleHeight     =   840
         ScaleWidth      =   8700
         TabIndex        =   15
         Top             =   6630
         Width           =   8760
         Begin VB.PictureBox fradados 
            Appearance      =   0  'Flat
            BackColor       =   &H00B63C18&
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   -15
            ScaleHeight     =   420
            ScaleWidth      =   8610
            TabIndex        =   16
            Top             =   270
            Width           =   8640
            Begin VB.TextBox txtVendedor 
               BackColor       =   &H00FFFFFF&
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
               Left            =   870
               MaxLength       =   3
               TabIndex        =   1
               Top             =   0
               Width           =   7725
            End
            Begin VB.TextBox txtPedido 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
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
               Left            =   -15
               TabIndex        =   0
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Label lblPedidoVendedor 
            BackColor       =   &H00B63C18&
            Caption         =   " Pedido     Vendedor"
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
            Left            =   -15
            TabIndex        =   35
            Top             =   45
            Width           =   1815
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdDadosProduto 
         Height          =   555
         Left            =   15
         TabIndex        =   19
         Top             =   6075
         Width           =   8715
         _cx             =   15372
         _cy             =   979
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
         BackColor       =   8513786
         ForeColor       =   12582912
         BackColorFixed  =   8513786
         ForeColorFixed  =   12582912
         BackColorSel    =   8513786
         ForeColorSel    =   15002065
         BackColorBkg    =   8513786
         BackColorAlternate=   8513786
         GridColor       =   11432977
         GridColorFixed  =   11432977
         TreeColor       =   11432977
         FloodColor      =   192
         SheetBorder     =   -2147483642
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPedido.frx":AADEC
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
         BackColorFrozen =   8513786
         ForeColorFrozen =   0
         WallPaper       =   "frmPedido.frx":AAEB7
         WallPaperAlignment=   9
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   330
         OleObjectBlob   =   "frmPedido.frx":150909
         Top             =   6360
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdPrecos 
         Height          =   3435
         Left            =   6345
         TabIndex        =   20
         Top             =   2370
         Width           =   2985
         _cx             =   5265
         _cy             =   6059
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
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   11942936
         ForeColorFixed  =   16777215
         BackColorSel    =   16744576
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483647
         GridColorFixed  =   -2147483647
         TreeColor       =   12582912
         FloodColor      =   192
         SheetBorder     =   12582912
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   13
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPedido.frx":150B3D
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
         BackColorFrozen =   11942936
         ForeColorFrozen =   16777215
         WallPaperAlignment=   9
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdItensProduto 
         Height          =   3705
         Left            =   15
         TabIndex        =   3
         Top             =   2370
         Width           =   6945
         _cx             =   12250
         _cy             =   6535
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
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   11942936
         ForeColorFixed  =   16777215
         BackColorSel    =   16744576
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483647
         GridColorFixed  =   -2147483647
         TreeColor       =   12582912
         FloodColor      =   192
         SheetBorder     =   12582912
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   13
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPedido.frx":150BA2
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
         BackColorFrozen =   11942936
         ForeColorFrozen =   16777215
         WallPaperAlignment=   0
      End
      Begin VB.Label lblQtde 
         BackColor       =   &H00B63C18&
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
         Left            =   7935
         TabIndex        =   37
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00B63C18&
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
         Left            =   15
         TabIndex        =   36
         Top             =   1740
         Width           =   4845
      End
   End
   Begin SHDocVwCtl.WebBrowser wbFichaTecnica 
      Height          =   6090
      Left            =   8160
      TabIndex        =   21
      Top             =   4275
      Visible         =   0   'False
      Width           =   6990
      ExtentX         =   12330
      ExtentY         =   10742
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash webInternet3 
      Height          =   2565
      Left            =   90
      TabIndex        =   24
      Top             =   75
      Width           =   15165
      _cx             =   26749
      _cy             =   4524
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim auxVEN_CodigoVendedor As Integer
Dim auxItens As Double
Dim wGuardaLinha As Long
'Dim wGuardaLinha2 As Long
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

Dim wBanner As String
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
Dim wTotalpedido As Double
Dim auxPedido As String
Dim AuxProdutoExiste As Boolean
Dim wProdutoNaoExiste As Boolean
Dim auxVendedordoPedido As Integer
Dim wQuantidade As Integer


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

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Private Sub cmbAgenda_Click()
  If txtPesquisar.Enabled = True Then
    frmAgenda.Show 1
    frmAgenda.ZOrder
  End If
End Sub

Private Sub cmbEMKT_Click()
 If txtPesquisar.Enabled = True Then
  frmEmarketing.Show 1
  frmEmarketing.ZOrder
 End If
End Sub

Private Sub CmdCarimbo_Click()
frmCarimbos.Show 1
frmCarimbos.ZOrder
End Sub

Private Sub CmdConsulta_Click()

 On Error GoTo ErronaDelecao
 
       rdoCNLoja.BeginTrans
       Screen.MousePointer = vbHourglass
       SQL = "Delete NFItens Where NumeroPed = " & txtPedido.Text & " and TipoNota = 'PD'"
       rdoCNLoja.Execute SQL
       SQL = "Delete NFCapa Where TipoNota = 'PD' and NumeroPed = " & txtPedido.Text
       Screen.MousePointer = vbNormal
       rdoCNLoja.CommitTrans
       Call LimpaForm
       Exit Sub
       
ErronaDelecao:
MsgBox "Erro na deleção do pedido " & Err.description, vbCritical, "Aviso"
rdoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
      
End Sub

Private Sub cmdCotacao_Click()
 FrmCotacao.Show 1
 FrmCotacao.ZOrder

End Sub

Private Sub cmdDesconto_Click()
 frmDesconto.Show 1
 frmDesconto.ZOrder
End Sub

Private Sub CmdDesfaz_Click()
  FrmDesfazProcesso.txtPedido = txtPedido.Text
  FrmDesfazProcesso.Show 1
  FrmDesfazProcesso.ZOrder
End Sub

Private Sub cmdFechaPedido_Click()
    Call FechaPedido
End Sub

Private Sub cmdFrete_Click()
  frmFrete.Show 1
  frmFrete.ZOrder
End Sub


Private Sub cmdPagamento_Click()
  frmPagamento.Show 1
  frmPagamento.ZOrder
End Sub

Private Sub cmdTR_Click()
  frmTR.Show 1
'  frmTR.ZOrder
End Sub

Private Sub cmdTransferencia_Click()
   frmTransferencia.Show 1
   frmTransferencia.ZOrder
End Sub

Private Sub cmdVendaDistancia_Click()
   frmVendaDistancia.Show 1
   frmVendaDistancia.ZOrder
   
End Sub

Private Sub cmdVerPedido_Click()
 'frmCompras.Show
 'frmCompras.ZOrder
End Sub

Private Sub Command1_Click()
'frmConsCliente.Top = 2100
'frmConsCliente.Left = 2625
frmConsCliente.Show 1
frmConsCliente.ZOrder
End Sub

Private Sub Form_Load()
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2
    frmPedido.txtVendedor.Width = 915
  frmPedido.fradados.Width = 1830

' Skin1.LoadSkin App.Path & "\Skin\SkinNovo.skn"
' Skin1.ApplySkin Me.hwnd

 On Error GoTo erro
 Call LerControleSistema
 Call LerControleCaixa
 Call VerificaInternet
 PicBanner.Visible = True
 
 wbFichaTecnica.Visible = False
' WebInternet2.Movie = "http://wwwimages.adobe.com/www.adobe.com/homepage/pt_br/fma_rotation/fma0/fma.swf?config=/homepage/pt_br/fma_rotation/fma0/fma_max2010_config.xml"
' WebInternet2.Play
 WebInternet2.Movie = "F:\sistemas\TraderBalcao\imagens\bradesco\demeo\mas.swf"
 WebInternet2.Play
 WebInternet1.Movie = "F:\sistemas\TraderBalcao\imagens\bradesco\demeo\TESTE LAMPADA BALCAO.swf"
 WebInternet1.Play
 webInternet3.Movie = "F:\sistemas\TraderBalcao\imagens\bradesco\demeo\BARRA LIMPA MAIOR C CARRINHO.swf"
 webInternet3.Play
 

 grdItensProduto.Enabled = False
 PicBanner.ZOrder
 RemoveMenus
  
erro:
    Exit Sub
End Sub

Private Sub grdItensProduto_DblClick()
      'If PicBanner.Visible = False Then
      wbFichaTecnica.Navigate2 wFichaTec & grdItensProduto.TextMatrix(grdItensProduto.Row, 0)
   '   PicBanner.Visible = True
      wbFichaTecnica.Visible = True
      wbFichaTecnica.ZOrder
     ' grdItensProduto.SetFocus
   'ElseIf PicBanner.Visible = True Then
      PicBanner.Visible = False
   '      wbFichaTecnica.Visible = True
   'End If

End Sub

Private Sub grdItensProduto_EnterCell()
    If PicBanner.Visible = False Then
       PicBanner.Visible = True
       wbFichaTecnica.Visible = False
    End If
End Sub

Private Sub grdItensProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       txtPesquisar.Enabled = True
       txtPesquisar.SetFocus
       txtPesquisar.SelStart = 0
       txtPesquisar.SelLength = Len(txtPesquisar.Text)
    ElseIf KeyCode = 13 Then
       grdItensProduto.BackColorSel = &HC0C000
       If grdPrecos.Enabled = True Then
          grdPrecos.SetFocus
          grdPrecos.Row = 1
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
  'PicBanner.Visible = True
  'wbFichaTecnica.Visible = False
End Sub

Private Sub grdItensProduto_RowColChange()

 txtPesquisar.Text = grdItensProduto.TextMatrix(grdItensProduto.Row, 1)
 PicBanner.Picture = LoadPicture("F:\sistemas\TraderBalcao\Imagens\BRADESCO\demeo\TESTE ACESSORIOS.jpg")
 grdDadosProduto.TextMatrix(1, 0) = grdItensProduto.TextMatrix(grdItensProduto.Row, 6)  'Linha
 grdDadosProduto.TextMatrix(1, 1) = grdItensProduto.TextMatrix(grdItensProduto.Row, 7)  'Descrição
' grdDadosProduto.TextMatrix(1, 2) = grdItensProduto.TextMatrix(grdItensProduto.Row, 8)  'Bloqueio
 grdDadosProduto.TextMatrix(1, 2) = grdItensProduto.TextMatrix(grdItensProduto.Row, 10)  'Bloqueio
 grdDadosProduto.TextMatrix(1, 3) = grdItensProduto.TextMatrix(grdItensProduto.Row, 9)  'Classe
' grdDadosProduto.TextMatrix(1, 4) = grdItensProduto.TextMatrix(grdItensProduto.Row, 10) 'ICMS
 grdDadosProduto.TextMatrix(1, 4) = grdItensProduto.TextMatrix(grdItensProduto.Row, 11) 'ICMS
 grdDadosProduto.TextMatrix(1, 5) = grdItensProduto.TextMatrix(grdItensProduto.Row, 12) 'ICMS Reduzido
 grdDadosProduto.TextMatrix(1, 6) = grdItensProduto.TextMatrix(grdItensProduto.Row, 13) 'ST
 wValorVenda = Format(grdItensProduto.TextMatrix(grdItensProduto.Row, 3), "0.00")
' Call MontaCrediario
 
 If grdPrecos.TextMatrix(0, 0) = "Financiado" Then
    Call MontaPrecos(2)
 ElseIf grdPrecos.TextMatrix(0, 0) = "Faturado" Then
    Call CarregaCondicaoFaturado
 ElseIf grdPrecos.TextMatrix(0, 0) = "A Vista" Then
    Call MontaPrecos(2)
 ElseIf grdPrecos.TextMatrix(0, 0) = "Cartão" Then
    Call MontaPrecos(3)
 End If
 
End Sub

Private Sub grdPrecos_DblClick()
    If grdPrecos.Col = 0 And cmdQtdeItens.Caption = 0 Then
       grdPrecos.Enabled = True
       txtCondicaoFaturado.Visible = True
       If grdPrecos.TextMatrix(0, 0) = "Financiado" Then
          grdPrecos.TextMatrix(0, 0) = "Cartão"
          Call MontaPrecos(3)
       ElseIf grdPrecos.TextMatrix(0, 0) = "Faturado" Then
          grdPrecos.TextMatrix(0, 0) = "A Vista"
          Call MontaPrecos(2)
       ElseIf grdPrecos.TextMatrix(0, 0) = "A Vista" Then
          grdPrecos.TextMatrix(0, 0) = "Financiado"
          Call MontaPrecos(2)
       ElseIf grdPrecos.TextMatrix(0, 0) = "Cartão" Then
          grdPrecos.TextMatrix(0, 0) = "Faturado"
          Call CarregaCondicaoFaturado
          'Call MontaPrecos(3)
       End If
    Else
       grdPrecos.Enabled = False
    End If
End Sub

Private Sub grdPrecos_EnterCell()

        txtCondicaoFaturado.Visible = True
        txtCondicaoFaturado.Text = grdPrecos.TextMatrix(grdPrecos.Row, 2)
        mskDatafaturado.Text = "__/__/____"

End Sub

Private Sub grdPrecos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
          grdItensProduto.BackColorSel = &HC00000
          grdItensProduto.SetFocus
     
          txtPesquisar.Enabled = True
          txtPesquisar.SetFocus
      
    ElseIf KeyCode = 13 Then
       
        If grdPrecos.TextMatrix(grdPrecos.Row, 0) = "85" Then
          fraCondicao.Enabled = True
'
'         txtCondicaoFaturado.SetFocus
          txtCondicaoFaturado.Visible = False
          txtCondicaoFaturado.Text = ""
          mskDatafaturado.SetFocus
                           
         Exit Sub
        End If
       
           grdPrecos.BackColorSel = &HC00000
           txtQuantidade.Enabled = True
           txtQuantidade.SetFocus
           grdPrecos.Enabled = False
           
    End If

End Sub

Private Sub cmbPedido_Click()
If cmdQtdeItens.Caption <> 0 Then
    frmConsultaItensdoPedido.Show 1
    frmConsultaItensdoPedido.ZOrder
End If
 
End Sub

Private Sub cmbPedido_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = vbCustom
End Sub

Private Sub cmbCliente_Click()
  If txtPesquisar.Enabled = True Then
    frmConsCliente.ZOrder
    frmConsCliente.Show 1
  End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub lblPesquisa1_DragDrop(Source As Control, x As Single, y As Single)

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
        If mskDatafaturado.Text = "" Then
            grdPrecos.SetFocus
        Else
            wGuardaLinha = grdPrecos.Row
        End If
    End If
End Sub


Private Sub s_Change()

End Sub

Private Sub tmrRefresh_Timer()

    Call LerControleSistema
    Call VerificaInternet

End Sub


Private Sub txtPedido_Change()
If IsNumeric(txtPedido.Text) = False Then
   txtPedido.Text = ""
End If
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   txtPedido.Text = ""
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
           rsPegaNumeroPedido.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 
           On Error GoTo erronoUpdate
        
           If Not rsPegaNumeroPedido.EOF Then
              rdoCNLoja.BeginTrans
              Screen.MousePointer = vbHourglass
              SQL = ""
              SQL = "Update ControleSistema set CTS_NumeroPedido=(CTS_NumeroPedido + 1)"
                    rdoCNLoja.Execute SQL
                    Screen.MousePointer = vbNormal
                    rdoCNLoja.CommitTrans
            
              txtPedido.Text = (rsPegaNumeroPedido("CTS_NumeroPedido"))
              auxPedido = (rsPegaNumeroPedido("CTS_NumeroPedido"))
              txtPedido.Enabled = False
              txtVendedor.Enabled = True
              txtPesquisar.Enabled = True
              txtVendedor.SetFocus
              rsPegaNumeroPedido.Close
              Exit Sub
          Else
              cmdLimpar.Caption = "Erro no Controle do Sistema avise o CPD"
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
      
      Call LerVendedordoPedido
      fradados.Enabled = False
      cmbPedido.Visible = True
      cmbCliente.Visible = True
      
      wPesquisaCodigo = 1
      inibebotoes (frmPedido.txtPedido)
     
      Exit Sub
   End If
  
End If
Exit Sub
erronoUpdate:
MsgBox "Erro na atualização do número do pedido " & Err.description, vbCritical, "Aviso"
rdoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal
rsPegaNumeroPedido.Close

End Sub

Private Sub txtPedido_LostFocus()
If txtPedido.Text = "" Then
   txtPedido.SetFocus
End If

End Sub

Private Sub txtPesquisar_Change()
  
 If fraMenu.Visible = True Then
    fraMenu.Visible = False
    fraMenu.Enabled = False
    PicBanner.Visible = True
    cmbPedido.Visible = True
    cmbCliente.Visible = True
 End If
 
End Sub

Private Sub txtPesquisar_GotFocus()
   txtPesquisar.SelStart = 0
   txtPesquisar.SelLength = Len(txtPesquisar.Text)
   If wbFichaTecnica.Visible = True Then
      wbFichaTecnica.Visible = False
      PicBanner.Visible = True
   End If
End Sub


Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   Call VerificaItensVendas
   
   If auxItens <> 0 Then
   
     If fraMenu.Visible = True Then
         Call FechaPedido
     Else
            txtPesquisar.SelStart = 0
            txtPesquisar.SelLength = Len(txtPesquisar.Text)
            cmbPedido.Visible = False
            cmbCliente.Visible = False
            PicBanner.Visible = False
            fraMenu.Visible = True
            fraMenu.Enabled = True
            txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF2 Then
      If fraMenu.Visible = True Then
         If cmdDesconto.Enabled = False Then
            Exit Sub
         End If
         frmDesconto.txtTotalPedido.Text = cmdTotalPedido.Caption
         frmDesconto.txtPedido = txtPedido.Text
         frmDesconto.Show 1
         frmDesconto.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
'  ElseIf KeyCode = vbKeyF3 Then
'      If fraMenu.Visible = True Then
'         If cmdPagamento.Enabled = False Then
'            Exit Sub
'         End If
'         frmPagamento.Show 1
'         frmPagamento.ZOrder
'      Else
'         txtPesquisar.SetFocus
'      End If
  ElseIf KeyCode = vbKeyF4 Then
      If fraMenu.Visible = True Then
         If cmdFrete.Enabled = False Then
            Exit Sub
         End If
         frmFrete.txtTotalPedido.Text = cmdTotalPedido.Caption
         frmFrete.txtPedido = txtPedido.Text
         frmFrete.Show 1
         frmFrete.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF5 Then
      If fraMenu.Visible = True Then
         If cmdTransferencia.Enabled = False Then
            Exit Sub
         End If
         frmTransferencia.Show 1
         frmTransferencia.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF6 Then
      If fraMenu.Visible = True Then
         If cmdVendaDistancia.Enabled = False Then
            Exit Sub
         End If
         frmVendaDistancia.Show 1
         frmVendaDistancia.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF7 Then
      If fraMenu.Visible = True Then
         If cmdTR.Enabled = False Then
            Exit Sub
         End If
          frmTR.Show 1
          frmTR.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF8 Then
      If fraMenu.Visible = True Then
         FrmDesfazProcesso.txtPedido = txtPedido.Text
         FrmDesfazProcesso.Show 1
         FrmDesfazProcesso.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF9 Then
      If fraMenu.Visible = True Then
         If CmdCarimbo.Enabled = False Then
            Exit Sub
         End If
         
         frmCarimbos.txtPedido = txtPedido.Text
         frmCarimbos.Show 1
         frmCarimbos.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF11 Then
      If fraMenu.Visible = True Then
         FrmCotacao.Show 1
         FrmCotacao.ZOrder
      Else
         txtPesquisar.SetFocus
      End If
  ElseIf KeyCode = vbKeyF12 Then
'***********************
'      Deleta Capa e Itens na consulta
       On Error GoTo ErronaDelecao
          rdoCNLoja.BeginTrans
          Screen.MousePointer = vbHourglass
          SQL = "Delete NFItens Where NumeroPed = " & txtPedido.Text & " and TipoNota = 'PD'"
          rdoCNLoja.Execute SQL
          
          SQL = "Delete NFCapa Where TipoNota = 'PD' and NumeroPed = " & txtPedido.Text
          rdoCNLoja.Execute SQL
          
          Screen.MousePointer = vbNormal
          rdoCNLoja.CommitTrans
          Call LimpaForm
'***********************
  ElseIf KeyCode = vbKeyTab Then
         txtPesquisar.SetFocus
  End If
 End If
Exit Sub

ErronaDelecao:
MsgBox "Erro na deleção do pedido " & Err.description, vbCritical, "Aviso"
rdoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal

End Sub

Private Sub txtPesquisar_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  If fraMenu.Visible = True Then
    fraMenu.Visible = False
    fraMenu.Enabled = False
    PicBanner.Visible = True
  End If

    If txtPesquisar.Text <> "" Then
            If IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) = 3 Then
                wWhere = " PR_CodigoFornecedor =" & Trim(txtPesquisar.Text) & " "
                PesquisarProduto wWhere ' Pesquisa por fonecedor
            ElseIf IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) = 7 Then
                wWhere = "PR_Referencia ='" & Trim(txtPesquisar.Text) & "' "
                PesquisarProduto wWhere ' Pesquisa por referencia
            ElseIf IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) > 3 Then
                wWhere = "PRB_CodigoBarras = '" & Trim(txtPesquisar.Text) & "' "
                PesquisarProduto wWhere ' Pesquisa por codigo de barras
            ElseIf IsNumeric(Trim(txtPesquisar.Text)) = False Then
                If IsNumeric(Mid(txtPesquisar.Text, 1, 3)) = True And Trim(Mid(txtPesquisar.Text, 4, 1)) = "" Then
                     wWhere = "PR_Descricao Like '" & Trim(UCase(Mid(Trim(txtPesquisar.Text), 4, _
                     Len(Trim(Trim(txtPesquisar.Text)))))) & "%' and PR_CodigoFornecedor = " _
                     & Mid(txtPesquisar, 1, 3)
                    PesquisarProduto wWhere  ' Pesquisa por Fornecedor e Descrição
                Else
                    wWhere = "PR_Descricao Like '" & UCase(Trim(txtPesquisar.Text)) & "%' "
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
                rdoControle.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

                If rdoControle.EOF = False Then
                    grdPrecos.TextMatrix(0, 0) = Trim(rdoControle("ModalidadeVenda"))
         
                    Select Case UCase(Trim(rdoControle("ModalidadeVenda")))
                         Case "FINANCIADO"
                              Call MontaPrecos(2)
                         Case "FATURADO"
                              Call CarregaCondicaoFaturado
                         Case "A VISTA"
                              Call MontaPrecos(2)
                         Case "CARTÃO"
                              Call MontaPrecos(3)
                    End Select
         
                    For Index = 1 To grdPrecos.Rows - 1
                        If Trim(Mid(grdPrecos.TextMatrix(Index, 0), 1, 2)) = Trim(rdoControle("Parcelas")) Then
                            grdPrecos.Row = Index
                            grdPrecos.Enabled = False
                            Exit For
                        End If
                    Next Index
         
                End If
                rdoControle.Close
         
            ElseIf cmdQtdeItens.Caption = 0 Then
               grdPrecos.Enabled = True
            End If
    Else
            txtPesquisar.SetFocus
    End If
ElseIf KeyAscii = 27 Then
     If txtPesquisar.Text <> "" Then
        fraMenu.Visible = False
        PicBanner.Visible = True
        cmbPedido.Visible = True
        cmbCliente.Visible = True
        txtPesquisar.SelStart = 0
        txtPesquisar.SelLength = Len(txtPesquisar.Text)
        txtPesquisar.SetFocus
     Else
        fraMenu.Visible = False
        PicBanner.Visible = True
        Call VerificaItensVendas
        
On Error GoTo ErroDeletaNFCapa
         If auxItens = 0 Then
            rdoCNLoja.BeginTrans
            SQL = "Delete NFCapa Where TipoNota = 'PD' and NumeroPed = " & txtPedido.Text
            rdoCNLoja.Execute SQL
            rdoCNLoja.CommitTrans
            End
        Else
           cmbPedido.Visible = True
           cmbCliente.Visible = True
           txtPesquisar.SetFocus
        End If
    End If
ElseIf (KeyAscii <> 27) Or (KeyAscii <> 13) Then
     txtPesquisar.SetFocus
End If
Exit Sub
 
ErroDeletaNFCapa:
    rdoCNLoja.RollbackTrans
    Exit Sub
    
End Sub

Function PesquisarProduto(ByVal wWhere As String)
    grdItensProduto.Rows = 1
      
    cmdLimpar.Caption = "Pesquisando ..."
    Screen.MousePointer = 11
            
    If auxItens <= 0 Then
        grdPrecos.TextMatrix(0, 0) = "A Vista"
    End If
            
'***********************************  Estoque Loja ***********************************************
'    SQL = "Select PR_ICMSSaida,PR_IcmPdv,PRB_CodigoBarras,PR_Referencia,PR_Descricao,PR_PrecoVenda1,EL_Estoque,PR_Classe," _
'        & "PR_Bloqueio,PR_SubstituicaoTributaria,LPR_Linha,LPR_Descricao " _
'        & "From ProdutoLoja, Produtobarras, EstoqueLoja, LinhaProduto " _
'        & "Where EL_Referencia=PR_Referencia and " & wWhere & " and PR_Situacao not in('E') and PRB_Referencia = PR_Referencia " _
'        & "and (Case When PR_LinhaProduto IS NULL Then '990100' Else PR_LinhaProduto End) = LPR_Linha and PRB_TipoCodigo = 'D' Order By PR_CodigoFornecedor,PR_Descricao"
            
    SQL = ""
    SQL = "Select (CASE WHEN PR_SubstituicaoTributaria = 'N' THEN PR_ICMSSaida ELSE PR_ICMSSaidaIva End) as IcmsSaida," & _
          "(CASE WHEN PR_SubstituicaoTributaria = 'N' THEN PR_IcmPdv ELSE PR_ICMSPDVSaidaIva End) as IcmsPdv,PRB_CodigoBarras,PR_Referencia,PR_Descricao,PR_PrecoVenda1,EL_Estoque,PR_Classe," & _
          "PR_Bloqueio,PR_SubstituicaoTributaria,LPR_Linha,LPR_Descricao " & _
          "From ProdutoLoja, Produtobarras, EstoqueLoja, LinhaProduto " & _
          "Where EL_Referencia=PR_Referencia and " & wWhere & " and PR_Situacao not in('E') and PRB_Referencia = PR_Referencia " & _
          "and (Case When PR_LinhaProduto IS NULL Then '990100' Else PR_LinhaProduto End) = LPR_Linha and PRB_TipoCodigo = 'D' " & _
          "Order By PR_CodigoFornecedor,PR_Descricao"
            
    rsPesquisaPed.CursorLocation = adUseClient
    rsPesquisaPed.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsPesquisaPed.EOF Then
       AuxProdutoExiste = True
       grdPrecos.Redraw = True
       txtPesquisar.Text = rsPesquisaPed("PR_Descricao")
       grdItensProduto.Redraw = False
       ReferenciaPreco = rsPesquisaPed("PR_PrecoVenda1")
        
        Do While Not rsPesquisaPed.EOF
            grdItensProduto.AddItem rsPesquisaPed("PR_Referencia") & Chr(9) _
                & rsPesquisaPed("PR_Descricao") & Chr(9) _
                & rsPesquisaPed("EL_Estoque") & Chr(9) _
                & Format(rsPesquisaPed("PR_PrecoVenda1"), "0.00") & Chr(9) _
                & "0,00" & Chr(9) _
                & rsPesquisaPed("PRB_CodigoBarras") & Chr(9) _
                & Trim(rsPesquisaPed("LPR_Linha")) & Chr(9) _
                & rsPesquisaPed("LPR_Descricao") & Chr(9) & "0" & Chr(9) _
                & rsPesquisaPed("PR_Classe") & Chr(9) _
                & rsPesquisaPed("PR_Bloqueio") & Chr(9) _
                & Format(rsPesquisaPed("IcmsSaida"), "0.00") & Chr(9) _
                & Format(rsPesquisaPed("IcmsPdv"), "0.00") & Chr(9) _
                & Trim(rsPesquisaPed("PR_SubstituicaoTributaria"))
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

Private Sub MontaPrecos(CodigoCrediario As Integer)
'    wGuardaLinha = grdPrecos.Row
    
  '-------------------------------------------------------------------------------
'   If wGuardaLinha = 1 Then
'       SQL = "Update NFCapa set condpag = 0 where NumeroPed = " & txtPedido.Text
'   Else
'       SQL = "Update NFCapa set condpag = '" & grdPrecos.TextMatrix(wGuardaLinha, 0) & _
'             "' where NumeroPed = " & txtPedido.Text
'   End If
'   rsCondicaoFaturado.CursorLocation = adUseClient
'   rsCondicaoFaturado.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  '-------------------------------------------------------------------------------------
    
    
    SQL = ""
    SQL = "Select * From Crediario Where CRE_CodigoCrediario = " & CodigoCrediario
    rsCrediario.CursorLocation = adUseClient
    rsCrediario.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         
    If Not rsCrediario.EOF Then
        grdPrecos.Rows = 1
        grdPrecos.Redraw = False
            
            For Index = 1 To rsCrediario("CRE_NroCampos")
                NomeParcela = "CRE_Parcelas" & Index
                NomeCoefic = "CRE_Coeficiente" & Index
                If UCase(Trim(rsCrediario(NomeParcela))) = "A VISTA" Then
                    QtdeParcelas = 1
                Else
                    QtdeParcelas = Trim(Mid(rsCrediario(NomeParcela), 1, 2))
                End If
            
                wValorTotalCalculado = Format((wValorVenda * rsCrediario(NomeCoefic)), "0.00")
                wPrecoCalculado = Format((wValorTotalCalculado / QtdeParcelas), "0.00")
    
                If UCase(Trim(rsCrediario(NomeParcela))) = "A VISTA" Then
                    grdPrecos.AddItem Trim(rsCrediario(NomeParcela)) & Chr(9) _
                                      & Format(wValorTotalCalculado, "###,###,###,##0.00")
                
                Else
                    grdPrecos.AddItem Trim(rsCrediario(NomeParcela)) & " " & Format(wPrecoCalculado, "###,###,###,##0.00") & Chr(9) _
                                      & Format(wValorTotalCalculado, "###,###,###,##0.00")
                
                End If
            Next Index
            
'        grdPrecos.Row = wGuardaLinha
        
        grdPrecos.Col = 0
        grdPrecos.ColSel = 1
        grdPrecos.Redraw = True
            
    End If
    
   If UCase(Trim(grdPrecos.TextMatrix(0, 0))) = "A VISTA" Then
      grdPrecos.Rows = 2
      grdPrecos.Row = 1
   Else
      If grdPrecos.Enabled = False Then
  '      grdPrecos.Enabled = True
         grdPrecos.Row = wGuardaLinha
         grdPrecos.Enabled = False
      Else
         grdPrecos.Row = wGuardaLinha
      End If
   End If
''   grdPrecos.Redraw = True
   rsCrediario.Close

End Sub

Private Sub MontaCrediario()

'    wGuardaLinha = grdPrecos.Row

    
    SQL = "Select * from Crediario where CRE_CodigoCrediario = 1 "

    rsCrediario.CursorLocation = adUseClient
    rsCrediario.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsCrediario.EOF Then
        grdPrecos.Rows = 1
        grdPrecos.Redraw = False
        For Index = 1 To rsCrediario("CRE_NroCampos")
            NomeParcela = "CRE_Parcelas" & Index '  NomeColuna = "TP_Preco" & i
            NomeCoefic = "CRE_Coeficiente" & Index
            If Mid(rsCrediario(NomeParcela), 1, 2) = "A " Then
               QtdeParcelas = 1
            Else
               QtdeParcelas = Mid(rsCrediario(NomeParcela), 1, 2)
            End If

            If rsCrediario(NomeParcela) <> " " Then
               wPrecoCalculado = Format((wValorVenda * rsCrediario(NomeCoefic)) / 100, "0.00")
               wValorTotalCalculado = Format(wPrecoCalculado + wValorVenda, "0.00")
               wPrecoCalculado = Format((wValorTotalCalculado) / QtdeParcelas, "0.00")
               grdPrecos.AddItem rsCrediario(NomeParcela) & Chr(9) _
               & Format(wValorTotalCalculado, "###,###,###,##0.00")
            End If
        Next Index
        
        grdPrecos.Row = wGuardaLinha
        grdPrecos.Col = 0
        grdPrecos.ColSel = 1
        grdPrecos.Redraw = True
       
    End If
   rsCrediario.Close
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
rsItensVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    If rsItensVenda.EOF = True Then
       auxItens = (auxItens + 1)
'       wPreco = Format(grdPrecos.TextMatrix(grdPrecos.Row, 1), "#0.00")
       wPreco = Format(grdPrecos.TextMatrix(wGuardaLinha, 1), "#0.00")
'       wVlunit2 = Format(wPreco, "#0.00")

       'Verificar Valores
       wVltotitem = Format((wPreco * txtQuantidade.Text), "##0.00")
       
       wDesconto = Format(0, "##0.00")
 '      wDesconto = wVltotitem - wVlunit2
       wIcms = Format(grdDadosProduto.TextMatrix(1, 5), "#0.00")
'----'
'       rdoCNLoja.BeginTrans
       Screen.MousePointer = vbHourglass
'**************************** Insert na Tabela NFItens
        SQL = "Insert into NFItens (NF,NUMEROPED,SERIE,DATAEMI,REFERENCIA,QTDE,VLUNIT, " _
            & "VLTOTITEM,ICMS,DESCONTO,PLISTA,LOJAORIGEM,TIPONOTA,Item) Values (0," _
            & txtPedido.Text & ",'','" & Format(Date, "yyyy/mm/dd") & "','" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "'," _
            & txtQuantidade.Text & "," & ConverteVirgula(wPreco) & "," & ConverteVirgula(wVltotitem) & "," _
            & ConverteVirgula(wIcms) & "," & ConverteVirgula(wDesconto) & "," & ConverteVirgula(wPreco) & ",'" & Trim(wLoja) & "'," _
            & "'PD'," & auxItens & ")"
                
        rdoCNLoja.Execute SQL
        Screen.MousePointer = vbNormal
'        rdoCNLoja.CommitTrans
    Else
        If MsgBox("Referência já cadastrada. Deseja somar a quantidade?", vbQuestion + vbYesNo, "Pedido") = vbYes Then
           SQL = ""
           SQL = "UPDATE NFItens set Qtde = (Qtde + " & txtQuantidade.Text & "), VLTOTITEM = ((vlunit - desconto) * (" & rsItensVenda("Qtde") & " + " & txtQuantidade.Text & ")) " _
                 & "Where NumeroPed = " & txtPedido.Text & " and Referencia = '" & grdItensProduto.TextMatrix(grdItensProduto.Row, 0) & "' and TipoNota = 'PD'"
'           rdoCNLoja.BeginTrans
           rdoCNLoja.Execute SQL
'           rdoCNLoja.CommitTrans
        Else
           grdItensProduto.SetFocus
        End If
    End If
        
rsItensVenda.Close
Exit Sub
        
        
erronaInclusao:
MsgBox "Erro na Inclusão de itens " & Err.description, vbCritical, "Aviso"

rdoCNLoja.RollbackTrans
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
      If fraMenu.Visible = True Then
        If frmPedido.cmdDesconto.Enabled = False Or _
           frmPedido.cmdFrete.Enabled = False Or _
           frmPedido.cmdCotacao.Enabled = False Or _
           frmPedido.cmdTR.Enabled = False Or _
           frmPedido.cmdVendaDistancia.Enabled = False Or _
           frmPedido.cmdTransferencia.Enabled = False Or _
           frmPedido.CmdCarimbo.Enabled = False Then
           txtQuantidade.SelLength = Len(txtQuantidade.Text)
           txtQuantidade.SetFocus
           cmbPedido.Visible = False
           cmbCliente.Visible = False
           
           Exit Sub
    
        End If
      Else
          If txtQuantidade.Text <> "" Then
             cmbPedido.Visible = True
             cmbCliente.Visible = True
          End If
      End If
   
   End If
   
      
     If KeyAscii = 13 Then
         If grdPrecos.Row <> 0 Then
             wGuardaLinha = grdPrecos.Row
         End If
     If IsNumeric(Trim(txtQuantidade.Text)) Then

        Call GravaItens
        Call SomaItensVenda
        
        SQL = ""
        SQL = "Update NFCapa Set ModalidadeVenda = '" & grdPrecos.TextMatrix(0, 0) & "'" & _
              " Where NumeroPed = " & (txtPedido.Text)
        rdoCNLoja.Execute SQL
        
        
        SQL = ""
     '-----------------------------------------------------------------------------------------
        If grdPrecos.TextMatrix(0, 0) = "Faturado" Then
            SQL = "Update NFCapa set condpag = '" & grdPrecos.TextMatrix(wGuardaLinha, 0) & _
                  "' where NumeroPed = " & txtPedido.Text
            If grdPrecos.TextMatrix(wGuardaLinha, 0) = "85" Then
                wGuardaPagamento = " - " & mskDatafaturado.Text
            Else
                wGuardaPagamento = " - " & txtCondicaoFaturado.Text
            End If
        Else
            If grdPrecos.TextMatrix(0, 0) = "Financiado" Then
                SQL = "Update NFCapa set condpag = '3' where NumeroPed = " & txtPedido.Text
                wGuardaPagamento = " - " & grdPrecos.TextMatrix(wGuardaLinha, 0)
            Else
                If grdPrecos.TextMatrix(0, 0) = "Cartão" Then
                    wGuardaPagamento = " - " & grdPrecos.TextMatrix(wGuardaLinha, 0)
                Else
                    wGuardaPagamento = " "
                End If
                SQL = "Update NFCapa set condpag = '0' where NumeroPed = " & txtPedido.Text
            End If
        End If
        '--------------------------- Manutencao Isnara 25 Maio 2011 --------------------------__
        rdoCNLoja.Execute SQL
        
                
        SQL = "Update NFCapa Set Parcelas = 0  Where ModalidadeVenda='A Vista' and NumeroPed = " & Val(txtPedido.Text)
        rdoCNLoja.Execute SQL

        
        grdPrecos.Enabled = False
        
        grdItensProduto.BackColorSel = &HFF8080
        grdPrecos.BackColorSel = &HFF8080
        txtQuantidade.Text = ""
        txtQuantidade.Enabled = False
        txtPesquisar.Enabled = True
'        grdItensProduto.SetFocus
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
         grdPrecos.BackColorSel = &HC00000
         grdPrecos.Enabled = True
         grdPrecos.SetFocus
      Else
         grdItensProduto.SetFocus
         grdItensProduto.BackColorSel = &HC00000
      End If
   End If
End Sub

Private Sub LerControleSistema()
  SQL = "Select * from ControleSistema"
  rdoControle.CursorLocation = adUseClient
  rdoControle.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
  If rdoControle.EOF Then
     MsgBox "Problemas com o Sistema (Controle Sistema) entrar em contato com o CPD", vbCritical, "Atenção"
     rdoControle.Close
     End
  Else
     wLoja = rdoControle("CTS_Loja")
     wBanner = Trim(rdoControle("CTS_CaminhoWeb1"))
     wFichaTec = Trim(rdoControle("CTS_CaminhoWeb2"))
     rdoControle.Close
  End If
  
End Sub

Private Sub LerControleCaixa()
SQL = "Select * from ControleCaixa where CTR_SituacaoCaixa='A' "
rdoControle.CursorLocation = adUseClient
rdoControle.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

If rdoControle.EOF Then
   MsgBox "Problemas com o Sistema (Controle Caixa) entrar em contato com o CPD", vbCritical, "Atenção"
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
  rsItensVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  auxItens = rsItensVenda("NroItens")
  rsItensVenda.Close
  
End Sub

Private Sub FechaPedido()
 auxItens = 0
 wCodigo = 1
 wSequencia = 1
 wValorDados = "V"
  
 On Error GoTo erronoUpdate
    Screen.MousePointer = vbHourglass
 
    SQL = ""
    SQL = "Select Referencia From NFItens Where NumeroPed = " & txtPedido.Text
 
    rsItensVenda.CursorLocation = adUseClient
    rsItensVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rsItensVenda.EOF = False Then
       rdoCNLoja.BeginTrans
       
       
'************************ Verificando se Nota é Eletrônica

       SQL = ""
       SQL = "select ce_Estado,ce_tipopessoa from cliente,nfcapa where ce_CodigoCliente = Cliente and " & _
             "NumeroPed = " & frmPedido.txtPedido.Text
            
            rsComplementoVenda.CursorLocation = adUseClient
            rsComplementoVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            If RTrim(LTrim(rsComplementoVenda("ce_Estado"))) <> "SP" _
                 Or RTrim(LTrim(rsComplementoVenda("ce_TipoPessoa"))) = "O" Then
                 
                   MsgBox "ESTE PEDIDO IRÁ GERAR UMA NOTA FISCAL ELETRÔNICA, AVISE O CLIENTE.", vbInformation, "Atenção"
                   
                   SQL = ""
                   SQL = "Update NfItens set Serie = 'NE' where NumeroPed = " & frmPedido.txtPedido.Text

                   rdoCNLoja.Execute (SQL)
            End If
            rsComplementoVenda.Close


'************************ Gravando Valores NFCapa
       SQL = ""
       SQL = "Exec SP_Totaliza_Capa_Nota_Fiscal " & txtPedido.Text
       rdoCNLoja.Execute SQL
       
                  
       SQL = ""
       SQL = "Select count(referencia) as NumeroItem from NFItens " _
           & "where NumeroPed=" & txtPedido.Text & ""
          
            rsComplementoVenda.CursorLocation = adUseClient
            rsComplementoVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
       SQL = ""
       SQL = "Update NFCapa set TipoNota = 'PA', qtditem = " & rsComplementoVenda("NumeroItem") & "" _
             & " Where NumeroPed = " & txtPedido.Text
       rdoCNLoja.Execute SQL
       
       rsComplementoVenda.Close
       
'************************ Gravando TipoNota NFItens
       SQL = "Update NFItens Set TipoNota = 'PA' Where NumeroPed = " & txtPedido.Text
       
       rdoCNLoja.Execute SQL
       rdoCNLoja.CommitTrans
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
       
erronoUpdate:
    MsgBox "Erro na atualizazão da situação do pedido " & Err.description, vbCritical, "Aviso"
    rdoCNLoja.RollbackTrans
    Screen.MousePointer = vbNormal
       
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)

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
    SQL = "Select VE_Codigo, VE_Nome From Vende WHERE VE_Codigo = " & txtVendedor.Text
        rsVendedor.CursorLocation = adUseClient
               
        rsVendedor.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
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
            txtVendedor.Width = 7740
            fradados.Enabled = False
            fradados.Width = 8640
            Call CriaCapaPedido(txtPedido.Text)
            cmbAgenda.Visible = True
            cmbEMKT.Visible = True
            txtPesquisar.Enabled = True
            txtPesquisar.SetFocus
        Else
            cmdLimpar.Caption = "Vendedor Não Cadastrado"
            MsgBox "Vendedor Não Cadastrado.", vbInformation, "Atenção"
            txtVendedor.SetFocus
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
  rsItensVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsItensVenda.EOF Then
     cmdTotalPedido.Caption = Format(rsItensVenda("TotalVenda"), "###,###,##0.00")
     cmdQtdeItens.Caption = rsItensVenda("TotalItens")
     auxItens = rsItensVenda("UltimoReg")
'     auxVendedordoPedido = rsItensVenda("Vendedor")
     cmdLimpar.Caption = ""
     
     If cmdQtdeItens.Caption <= 1 Then
        cmdItens.Caption = "Item."
     Else
        cmdItens.Caption = "Itens."
     End If
        
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
      rsVendedor.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
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

If txtVendedor.Text = "" Then
   txtVendedor.SetFocus
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
  hMenu = GetSystemMenu(hwnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Function CriaCapaPedido(ByVal NumeroPedido As Double)
      
    SQL = ""
    SQL = "Select count(referencia) as NumeroItem from NFItens " _
          & "where NumeroPed=" & NumeroPedido & ""
          
          rsComplementoVenda.CursorLocation = adUseClient
          rsComplementoVenda.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    SQL = ""
    SQL = "Insert Into NFCapa(NUMEROPED,DATAEMI,LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA, " _
        & "VendedorLojaVenda, LojaVenda,TM,qtditem, OutraLoja, OutroVend) " _
        & "Values (" & NumeroPedido & ",'" & Format(Date, "mm/dd/yyyy") & "', " _
        & "'" & wLoja & "','PD'," & auxVEN_CodigoVendedor & ", " _
        & "'" & Format(Date, "mm/dd/yyyy") & "', '" & Format(Time, "hh:mm:ss") & "', " _
        & auxVEN_CodigoVendedor & ", '" & wLoja & "',0," _
        & rsComplementoVenda("Numeroitem") & "," & wLoja & "," & auxVEN_CodigoVendedor & ")"
        rdoCNLoja.Execute (SQL)
     
     rsComplementoVenda.Close
     
End Function

'Private Sub LimpaForm()
'Dim cObjeto As Control
'Dim wColuna As Integer
  
'  For Each cObjeto In Me.Controls
'      If (TypeOf cObjeto Is TextBox) Then
'        cObjeto.Text = ""
'      End If
'  Next
'
'  For Each cObjeto In Me.Controls
'      If (TypeOf cObjeto Is CommandButton) Then
'        cObjeto.Enabled = True
'      End If
'  Next
'
'  For wColuna = 0 To frmPedido.grdDadosProduto.Cols - 1
'     frmPedido.grdDadosProduto.TextMatrix(1, wColuna) = ""
'  Next wColuna
  
'  frmPedido.grdItensProduto.Rows = 1
'  frmPedido.grdPrecos.Rows = 1
'  frmPedido.fraMenu.Visible = False
'  frmPedido.txtPesquisar.Enabled = False
'  frmPedido.txtQuantidade.Enabled = False
'  frmPedido.grdItensProduto.Enabled = False
'  frmPedido.grdPrecos.Enabled = False
'  frmPedido.grdDadosProduto.Enabled = False
'  frmPedido.wbFichaTecnica.Visible = False
              
'  frmPedido.cmdQtdeItens.Caption = 0
'  frmPedido.txtTotalPedido.text = 0
'  auxItens = 0
              
'  frmPedido.PicBanner.Visible = True
'  frmPedido.cmbPedido.Visible = True
'  frmPedido.cmbCliente.Visible = True
'  frmPedido.fraMenu.Visible = False
'  frmPedido.fraMenu.Enabled = False
  'picQuadroGeral.Width = 9975

'  frmPedido.fradados.Enabled = True
'  frmPedido.txtVendedor.Width = 915
'  frmPedido.txtVendedor.Enabled = False
'  frmPedido.txtPedido.Enabled = True
'  frmPedido.txtPedido.SetFocus
'
'End Sub

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
       
'On Error GoTo ErroInternet
'  Inet1.RequestTimeout = 3
'  wVerificaInternet = Inet1.OpenURL(wBanner)
'  If Trim(wVerificaInternet) = "" Or Trim(wVerificaInternet) = "Invalid Request" Then
'     wBanner = App.Path & "\Skin\DEMEO.swf"
'  End If
'  Exit Sub


'ErroInternet:
    wBanner = App.Path & "\Skin\DEMEO.swf"
'    Exit Sub
  
End Sub

Private Sub CarregaCondicaoFaturado()

  grdPrecos.Rows = 1
  grdPrecos.Redraw = False

  SQL = ""
  SQL = "Select * from CondicaoPagamento where CP_Tipo = 'FA' Order By CP_Codigo"
  rsCondicaoFaturado.CursorLocation = adUseClient
  rsCondicaoFaturado.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

     Do While Not rsCondicaoFaturado.EOF

 '       wValorTotalCalculado = Format((wValorVenda * rsCondicaoFaturado("CP_Coeficiente")), "0.00")
        wValorTotalCalculado = Format((wValorVenda * 1), "0.00")
 '       wValorTotalCalculado = Format(wValorVenda, "0.00")
        wPrecoCalculado = Format((wValorTotalCalculado / rsCondicaoFaturado("CP_Parcelas")), "0.00")

        grdPrecos.AddItem rsCondicaoFaturado("CP_Codigo") & Chr(9) _
        & Format(wPrecoCalculado, "###,###,###,##0.00") & Chr(9) _
        & rsCondicaoFaturado("CP_Condicao")

        rsCondicaoFaturado.MoveNext
     Loop



        grdPrecos.Col = 0
        grdPrecos.ColSel = 1

        grdPrecos.Redraw = True

        
        grdPrecos.TopRow = wGuardaLinha
        
        
 '      grdPrecos.BackColorSel = &HFF8080
        
 '      grdPrecos.Row(wGuardaLinha).DefaultCellStyle.BackColor = &HFF8080





  rsCondicaoFaturado.Close
End Sub



