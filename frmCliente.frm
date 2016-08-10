VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCliente 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   5895
   ClientLeft      =   1950
   ClientTop       =   3270
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14820
      TabIndex        =   68
      Top             =   5085
      Width           =   14820
   End
   Begin VB.Timer tmrTroca 
      Left            =   14700
      Top             =   210
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00505050&
      Caption         =   "Dados do Cliente"
      ForeColor       =   &H00E0E0E0&
      Height          =   4440
      Left            =   60
      TabIndex        =   26
      Top             =   350
      Width           =   15030
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   70
         Top             =   585
         Width           =   1095
      End
      Begin VB.TextBox txtRazaoSocial 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2715
         MaxLength       =   40
         TabIndex        =   1
         Top             =   585
         Width           =   5175
      End
      Begin VB.ComboBox cmbPessoa 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7965
         TabIndex        =   2
         Top             =   585
         Width           =   1665
      End
      Begin VB.TextBox txtClienteFidelidade 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10110
         MaxLength       =   18
         TabIndex        =   34
         Top             =   3900
         Width           =   1740
      End
      Begin VB.ComboBox cmbTipoCliente 
         BackColor       =   &H00A3A3A3&
         Height          =   315
         Left            =   11235
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   1590
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00505050&
         Caption         =   "Dados de Cobrança"
         ForeColor       =   &H00E0E0E0&
         Height          =   1140
         Left            =   105
         TabIndex        =   57
         Top             =   2400
         Width           =   14745
         Begin VB.TextBox txtNumCobranca 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5595
            MaxLength       =   5
            TabIndex        =   21
            Top             =   585
            Width           =   750
         End
         Begin VB.TextBox txtEnderecoCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1140
            MaxLength       =   40
            TabIndex        =   20
            Top             =   585
            Width           =   4380
         End
         Begin VB.TextBox txtComplCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10080
            MaxLength       =   15
            TabIndex        =   25
            Top             =   585
            Width           =   1710
         End
         Begin VB.TextBox txtBairroCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11865
            MaxLength       =   15
            TabIndex        =   27
            Top             =   585
            Width           =   2775
         End
         Begin VB.TextBox txtMunicipioCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6420
            MaxLength       =   60
            TabIndex        =   22
            Top             =   585
            Width           =   2940
         End
         Begin VB.TextBox txtEstadoCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9435
            MaxLength       =   2
            TabIndex        =   23
            Top             =   585
            Width           =   570
         End
         Begin VB.TextBox mskCepCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   105
            MaxLength       =   8
            TabIndex        =   19
            Top             =   585
            Width           =   960
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Município"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   17
            Left            =   6420
            TabIndex        =   64
            Top             =   315
            Width           =   705
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   15
            Left            =   1155
            TabIndex        =   63
            Top             =   315
            Width           =   690
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   16
            Left            =   11865
            TabIndex        =   62
            Top             =   315
            Width           =   405
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   19
            Left            =   9435
            TabIndex        =   61
            Top             =   315
            Width           =   495
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   20
            Left            =   105
            TabIndex        =   60
            Top             =   315
            Width           =   315
         End
         Begin VB.Label lblNumeroCobranca 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "N.º"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   5595
            TabIndex        =   59
            Top             =   315
            Width           =   225
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   12
            Left            =   10080
            TabIndex        =   58
            Top             =   315
            Width           =   960
         End
      End
      Begin VB.ComboBox cmbSituacao 
         BackColor       =   &H00A3A3A3&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12885
         TabIndex        =   5
         Top             =   585
         Width           =   1935
      End
      Begin VB.TextBox mskFax 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11610
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1890
         Width           =   1485
      End
      Begin VB.TextBox mskCelular 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10005
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1890
         Width           =   1530
      End
      Begin VB.TextBox mskTelefone 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1890
         Width           =   1530
      End
      Begin VB.TextBox mskCep 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1260
         Width           =   1305
      End
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1260
         MaxLength       =   14
         TabIndex        =   0
         Top             =   585
         Width           =   1395
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5235
         TabIndex        =   33
         Top             =   3900
         Width           =   4785
      End
      Begin VB.TextBox txtBairro 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10530
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1260
         Width           =   4320
      End
      Begin VB.TextBox txtComplemento 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7035
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1260
         Width           =   3420
      End
      Begin VB.TextBox txtMunicipio 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   90
         MaxLength       =   60
         TabIndex        =   11
         Top             =   1890
         Width           =   4545
      End
      Begin VB.TextBox txtEndereco 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1485
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1260
         Width           =   4710
      End
      Begin VB.TextBox txtInscricaoEstadual 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9690
         MaxLength       =   15
         TabIndex        =   3
         Top             =   585
         Width           =   1470
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdMunicipio 
         Height          =   360
         Left            =   75
         TabIndex        =   24
         Top             =   2205
         Visible         =   0   'False
         Width           =   4575
         _cx             =   8070
         _cy             =   635
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
         GridLinesFixed  =   0
         GridLineWidth   =   0
         Rows            =   8
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCliente.frx":0000
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
         WallPaperAlignment=   9
      End
      Begin VB.TextBox txtCodMun 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4725
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1890
         Width           =   1395
      End
      Begin VB.ComboBox cmbSegmento 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2745
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3900
         Width           =   2400
      End
      Begin VB.ComboBox cmbRamoAtiv 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3900
         Width           =   2550
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6270
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1260
         Width           =   690
      End
      Begin VB.ComboBox cmbUf 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6195
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1890
         Width           =   750
      End
      Begin VB.ComboBox cmbPraca 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7020
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1890
         Width           =   1305
      End
      Begin VB.CheckBox chkCart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00505050&
         Caption         =   "Pagamento Carteira"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   13140
         TabIndex        =   18
         Tag             =   "Cart"
         Top             =   1935
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskDataCadastro 
         Height          =   315
         Left            =   13575
         TabIndex        =   71
         Tag             =   "DataCadastro"
         Top             =   3900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataNascimento 
         Height          =   315
         Left            =   11940
         TabIndex        =   36
         Tag             =   "DataCadastro"
         Top             =   3900
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente Fidelidade"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   10110
         TabIndex        =   67
         Top             =   3645
         Width           =   1245
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome *"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2745
         TabIndex        =   66
         Top             =   345
         Width           =   615
      End
      Begin VB.Label lblTipCli 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cliente *"
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   11235
         TabIndex        =   65
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   12855
         TabIndex        =   56
         Top             =   315
         Width           =   510
      End
      Begin VB.Label lblClite 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Município *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   4725
         TabIndex        =   55
         Top             =   1635
         Width           =   1380
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Segmento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   36
         Left            =   2745
         TabIndex        =   54
         Top             =   3645
         Width           =   720
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Ramo de Atividade"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   35
         Left            =   105
         TabIndex        =   53
         Top             =   3645
         Width           =   1350
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Nascimento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   30
         Left            =   11925
         TabIndex        =   52
         Top             =   3645
         Width           =   1455
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   18
         Left            =   10005
         TabIndex        =   51
         Tag             =   "Cel"
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   7035
         TabIndex        =   50
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "N.º *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6270
         TabIndex        =   49
         Top             =   1005
         Width           =   330
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   5235
         TabIndex        =   48
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   8
         Left            =   10530
         TabIndex        =   47
         Top             =   1020
         Width           =   510
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   7
         Left            =   1485
         TabIndex        =   46
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   14
         Left            =   11610
         TabIndex        =   45
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblCep 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   105
         TabIndex        =   44
         ToolTipText     =   "Click para consultar o cep da rua "
         Top             =   1005
         Width           =   420
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "UF *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   11
         Left            =   6195
         TabIndex        =   43
         Top             =   1635
         Width           =   315
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Praça"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   10
         Left            =   7020
         TabIndex        =   42
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "I.E. *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   9705
         TabIndex        =   41
         Top             =   345
         Width           =   345
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   13
         Left            =   8400
         TabIndex        =   40
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Municipio *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   39
         Top             =   1635
         Width           =   780
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   7965
         TabIndex        =   35
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Cadastro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   13560
         TabIndex        =   32
         Top             =   3645
         Width           =   1245
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   1260
         TabIndex        =   30
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Código *"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   28
         Top             =   345
         Width           =   600
      End
   End
   Begin Project1.chameleonButton cmdGravar 
      Height          =   405
      Left            =   13920
      TabIndex        =   37
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Grava"
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
      MICON           =   "frmCliente.frx":0050
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton CmdLimpa 
      Height          =   405
      Left            =   12765
      TabIndex        =   38
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Limpa"
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
      MICON           =   "frmCliente.frx":006C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton CmdFichaFinanc 
      Height          =   405
      Left            =   11085
      TabIndex        =   72
      Top             =   5280
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Ficha Financeira"
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
      MICON           =   "frmCliente.frx":0088
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
      TabIndex        =   69
      Top             =   0
      Width           =   15630
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim WMunicipio As String
Dim Cliente As Double
Dim auxCMBUF As String
Dim wPreencheInicio As Boolean
Dim Ln As Integer
Dim wMskInscEst As String
Dim wNumeros As String
Dim adoBuscaClienteFaturado As New ADODB.Recordset
Dim adoCodigoCliente As New ADODB.Recordset
Dim adoCliente As New ADODB.Recordset
Dim adoClienteCNPJ As New ADODB.Recordset
Dim adoMunicipio As New ADODB.Recordset
Dim rsSituacaoCliente As New ADODB.Recordset
Dim wPreencherCliente As Boolean
'Dim wNumeroClientePedido As Integer
Dim novoCodigo As Integer
Dim I As Integer
Dim pagamentoCarteira As String
Dim dataNascimento As String
Dim wLimpar As Boolean
Dim SQL As String


Function FU_ValidaCPF(CPF As String) As Integer

    Dim Soma As Integer
    Dim Resto As Integer
    Dim I As Integer

    If Len(CPF) <> 11 Then
        FU_ValidaCPF = False
        Exit Function
    End If

    Soma = 0
    For I = 1 To 9
        Soma = Soma + Val(Mid$(CPF, I, 1)) * (11 - I)
    Next I
    Resto = 11 - (Soma - (Int(Soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
        
    Soma = 0
    For I = 1 To 10
        Soma = Soma + Val(Mid$(CPF, I, 1)) * (12 - I)
    Next I
    Resto = 11 - (Soma - (Int(Soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
    
    FU_ValidaCPF = True

End Function

Function FU_ValidaCGC(CGC As String) As Integer
        Dim retorno, a, j, I, d1, d2
        If Len(CGC) = 8 And Val(CGC) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For I = 1 To 7
               a = Val(Mid(CGC, I, 1))
               If (I Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next I
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(CGC, 8, 1)) Then
              FU_ValidaCGC = True
           Else
              FU_ValidaCGC = False
           End If
        Else
           If Len(CGC) = 14 And Val(CGC) > 0 Then
              a = 0
              I = 0
              d1 = 0
              d2 = 0
              j = 5
              For I = 1 To 12 Step 1
                  a = a + (Val(Mid(CGC, I, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next I
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              I = 0
              j = 6
              For I = 1 To 13 Step 1
                  a = a + (Val(Mid(CGC, I, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next I
              a = a Mod 11
              d2 = IIf(a > 1, 11 - a, 0)
              If (d1 = Val(Mid(CGC, 13, 1)) And d2 = Val(Mid(CGC, 14, 1))) Then
                 FU_ValidaCGC = True
              Else
                 FU_ValidaCGC = False
              End If
           Else
              FU_ValidaCGC = False
           End If
        End If
End Function

Private Sub validaDados(codigoCliente As Double)

    Dim adoValidaCliente As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "exec SP_GLB_Valida_Cliente '" & wNumeroClientePedido & "'"
    adoCNLoja.Execute SQL
    
    SQL = "select campoErrado as campoErrado from temp_Fin_Cliente_Erro"
    adoValidaCliente.CursorLocation = adUseClient
    adoValidaCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Do While Not adoValidaCliente.EOF
        If adoValidaCliente("campoErrado") = "CGC" Then lblClite(4).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "InscricaoEstadual" Then lblClite(6).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Razao" Then lblNome.ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Endereco" Then lblClite(7).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Bairro" Then lblClite(8).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Municipio" Then lblClite(9).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "CEP" Then lblCep.ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Telefone" Then lblClite(13).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Fax" Then lblClite(14).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "EMail" Then lblEmail(0).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Numero" Then lblNumero.ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Celular" Then lblClite(18).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "Mun_Codigo" Then lblClite(3).ForeColor = vbRed
        If adoValidaCliente("campoErrado") = "CodigoMunicipio" Then lblClite(3).ForeColor = vbRed
        adoValidaCliente.MoveNext
    Loop
    
    adoValidaCliente.Close
    
End Sub

Private Sub chameleonButton1_Click()

End Sub

Private Sub chkCart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 83 Then
        chkCart.Value = 1
    ElseIf KeyCode = 78 Then
        chkCart.Value = 0
    ElseIf KeyCode = 13 Then
        ProximoCampo txtInscricaoEstadual
        SelecionaCampo txtInscricaoEstadual
    ElseIf KeyCode = 27 Then
        ProximoCampo cmbSituacao
    End If
End Sub

Private Sub cmbPessoa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbPessoa.Text <> "" Then
            cmbPessoa.Text = UCase(cmbPessoa.Text)
            If Mid(UCase(cmbPessoa.Text), 1, 1) = "J" Then
                cmbPessoa.Text = "JURÍDICA"
                Call CarregaRamo("JURÍDICA", txtCodigo.Text)
            ElseIf Mid(UCase(cmbPessoa.Text), 1, 1) = "U" Or Mid(UCase(cmbPessoa), 1, 2) = "FU" Then
                cmbPessoa.Text = "FUNCIONÁRIO"
                Call CarregaRamo("FUNCIONÁRIO", txtCodigo.Text)
            ElseIf Mid(UCase(cmbPessoa.Text), 1, 1) = "F" Then
                   cmbPessoa.Text = "FÍSICA"
                   Call CarregaRamo("FÍSICA", txtCodigo.Text)
            ElseIf Mid(UCase(cmbPessoa.Text), 1, 1) = "O" Then
                cmbPessoa.Text = "ÓRGÃO PÚBLICO"
                Call CarregaRamo("ÓRGÃO PÚBLICO", txtCodigo.Text)
            Else
                cmbPessoa.SelStart = 0
                cmbPessoa.SelLength = Len(cmbPessoa.Text)
            End If
            If cmbSituacao.Enabled = True Then
                ProximoCampo cmbSituacao
                SelecionaCampo cmbSituacao
            Else
                ProximoCampo chkCart
            End If
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtCnpj
        SelecionaCampo txtCnpj
    End If
End Sub

Private Sub cmbPessoa_LostFocus()
    If cmbPessoa.Text <> "" Then
        cmbPessoa.Text = UCase(cmbPessoa.Text)
        If Mid(UCase(cmbPessoa.Text), 1, 1) = "J" Then
            cmbPessoa.Text = "JURÍDICA"
            txtInscricaoEstadual.Locked = True
            txtInscricaoEstadual.Locked = False
            'txtInscricaoEstadual.Text = ""
            txtCnpj.MaxLength = 14
            mskDataNascimento.Enabled = False
        ElseIf Mid(UCase(cmbPessoa.Text), 1, 1) = "U" Or Mid(UCase(cmbPessoa), 1, 2) = "FU" Then
            txtInscricaoEstadual.Text = "ISENTO"
            cmbPessoa.Text = "FUNCIONÁRIO"
            txtInscricaoEstadual.Locked = False
            txtCnpj.MaxLength = 11
        ElseIf Mid(UCase(cmbPessoa.Text), 1, 1) = "F" Then
            txtInscricaoEstadual.Text = "ISENTO"
            cmbPessoa.Text = "FÍSICA"
            txtInscricaoEstadual.Locked = False
            txtCnpj.MaxLength = 11
        ElseIf Mid(UCase(cmbPessoa.Text), 1, 1) = "Ó" Then
            cmbPessoa.Text = "ÓRGÃO PÚBLICO"
            txtInscricaoEstadual.Locked = True
            txtInscricaoEstadual.Locked = False
            txtInscricaoEstadual.Text = ""
            txtCnpj.MaxLength = 14
            mskDataNascimento.Enabled = False
        Else
            cmbPessoa.SelStart = 0
            'cmbPessoa.SelLength = Len(cmbPessoa.Text)
            MsgBox "Tipo de Pessoa inválido! Informe: Física, Jurídica, Funcionário ou Órgão Público.", vbCritical, "Atenção"
            cmbPessoa.ListIndex = 0
            cmbPessoa.SetFocus
        End If
    Else
    MsgBox "Tipo de Pessoa inválido! Informe: Física, Jurídica, Funcionário ou Órgão Público.", vbCritical, "Atenção"
    cmbPessoa.ListIndex = 0
    cmbPessoa.SetFocus
    End If
    
    'If cmbPessoa.Text = "FÍSICA" Or cmbPessoa.Text = "FUNCIONÁRIO" And txtCodigo.Text = "" Then
    '    txtInscricaoEstadual.locked = False
    '    txtCNPJ.Text = ""
    'Else'
'
'    End If
    
    
    Call CarregaRamo(cmbPessoa.Text, txtCodigo.Text)
End Sub

Private Sub cmbPraca_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If cmbPraca.Text <> "" Then
            ProximoCampo cmbUF
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtMunicipio
        SelecionaCampo txtMunicipio
    End If

End Sub

Private Sub cmbRamoAtiv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbRamoAtiv.Text <> "" Then
            ProximoCampo cmbSegmento
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo mskDataNascimento
        SelecionaCampo mskDataNascimento
    End If

End Sub

Private Sub cmbRamoAtiv_LostFocus()
     Call carregaSegmento(Mid(cmbRamoAtiv.Text, 1, 2), txtCodigo.Text)
End Sub

Private Sub cmbSegmento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbSegmento.Text <> "" Then
            ProximoCampo txtEMail
            SelecionaCampo txtEMail
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo cmbRamoAtiv
    End If
End Sub

Private Sub cmbSituacao_Change()
        If cmbSituacao.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        cmbSituacao.Text = ""
        cmbSituacao.SetFocus
        Exit Sub
        End If
End Sub

Private Sub cmbTipoCliente_LostFocus()
    If Mid(cmbTipoCliente.Text, 1, 1) = "F" Then
    
        SQL = "Exec SP_Busca_Codigo_Cliente_Faturado"
    
        adoCodigoCliente.CursorLocation = adUseClient
        adoCodigoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        Call BuscaClienteFaturado
    End If
End Sub

Private Sub BuscaClienteFaturado()
    SQL = "Select cts_codigoclientefaturado from controlesistema"
 
    adoBuscaClienteFaturado.CursorLocation = adUseClient
    adoBuscaClienteFaturado.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    txtCodigo.Text = adoBuscaClienteFaturado("cts_codigoclientefaturado")
    adoBuscaClienteFaturado.Close
End Sub


Private Sub cmbUf_Change()
    txtEstadoCobranca.Text = cmbUF.Text
    
End Sub

Private Sub cmbUf_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If cmbUF.Text <> "" Then
            cmbUF.Text = UCase(cmbUF.Text)
            ProximoCampo mskTelefone
            SelecionaCampo mskTelefone
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo cmbPraca
    End If
End Sub

Private Sub cmdAtualizaDadosCliente_Click()

End Sub

Private Sub CmdFichaFinanc_Click()

    If rdoCNMatriz.State <> 1 Then
        
        If ConexaoDLLAdo.abrirConexaoADO(rdoCNMatriz, GLB_Servidor, GLB_Banco) Then
            GLB_ConectouOK = True
        End If
        
    End If

    wConexao = "Balcao"
    wCodigoCliFinan = txtCodigo.Text
        

        
        Cliente = txtCodigo.Text
        SQL = "exec SP_fin_Ler_Cliente_Limite_Credito " & Cliente & ", '" & Trim(GLB_Loja) & "'"
        rdoCNMatriz.Execute (SQL)
        
        SQL = " Exec SP_FIN_Pesquisa_Cliente_Ficha_Financeira_Por_Codigo '" & Cliente & "'"
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
'
'   If adoCliente.EOF = True Then
'         MsgBox "Este cliente não possue ficha finaceira", vbCritical, "Atenção"
'         adoCliente.Close
'         Exit Sub
'   Else
         FrmFichaFinanceira.Show 1
'   End If

   ' adoCliente.Close

End Sub

Private Sub CmdGravar_Click()


lblClite(4).ForeColor = &HE0E0E0
lblClite(6).ForeColor = &HE0E0E0
lblNome.ForeColor = &HE0E0E0
lblClite(7).ForeColor = &HE0E0E0
lblClite(8).ForeColor = &HE0E0E0
lblClite(9).ForeColor = &HE0E0E0
lblCep.ForeColor = &HE0E0E0
lblClite(13).ForeColor = &HE0E0E0
lblClite(14).ForeColor = &HE0E0E0
lblEmail(0).ForeColor = &HE0E0E0
lblNumero.ForeColor = &HE0E0E0
lblClite(18).ForeColor = &HE0E0E0
lblClite(3).ForeColor = &HE0E0E0
lblClite(3).ForeColor = &HE0E0E0

    If verificaCamposNulos = True Then

        Call lerClientePorCNPJ
        
        Screen.MousePointer = 11
        If adoClienteCNPJ.EOF Then
            
            cmbSituacao.Enabled = False

            Call GravaCliente

        Else
            If adoClienteCNPJ("CE_CodigoCliente") = Val(txtCodigo.Text) Then
                Call AtualizaCliente(txtCodigo.Text)
            Else
                MsgBox "CNPJ/CPJ já possui cadastro"
                txtCnpj.Locked = True
                txtCnpj.SetFocus
                txtCnpj.SelStart = 0
                txtCnpj.SelLength = Len(txtCnpj.Text)
           End If
       End If
       adoClienteCNPJ.Close
       Screen.MousePointer = 0
'       Call Limpar
    End If
End Sub
Private Sub lerClientePorCNPJ()

 SQL = " exec SP_FIN_Ler_Clientes_Por_Parametro_Cnpj '" & txtCnpj.Text & "'"
    
     adoClienteCNPJ.CursorLocation = adUseClient
     adoClienteCNPJ.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     


End Sub

Private Sub cmdLimpa_Click()
    Call Limpar
End Sub
Private Sub Limpar()
    wLimpar = True
    txtRazaoSocial.Text = ""
    txtCnpj.Text = ""
    txtCnpj.Locked = False
    txtCnpj.Locked = True
    cmbPessoa.Enabled = False
    cmbPessoa.Text = ""
    cmbSituacao.Text = "ATIVO"
    chkCart.Value = False
    txtInscricaoEstadual.Text = ""
    mskCep.Text = ""
    txtEndereco.Text = ""
    txtNumero.Text = ""
    txtMunicipio.Text = ""
    txtCodMun.Text = ""
    cmbUF.ListIndex = -1
    txtComplemento.Text = ""
    txtBairro.Text = ""
    cmbPraca.ListIndex = -1
    mskTelefone.Text = ""
    mskCelular.Text = ""
    mskFax.Text = ""
    txtClienteFidelidade.Text = ""
    mskDataNascimento.Text = ""
    cmbRamoAtiv.ListIndex = -1
    txtEMail.Text = ""
    cmbSegmento.ListIndex = -1
    mskDataCadastro.Text = ""
    txtEnderecoCobranca.Text = ""
    txtNumCobranca.Text = ""
    mskCepCobranca.Text = ""
    txtComplCobranca.Text = ""
    txtBairroCobranca.Text = ""
    txtMunicipioCobranca.Text = ""
    txtEstadoCobranca.Text = ""
    txtCodigo.Locked = True
    txtCodigo.Locked = False
    txtCodigo.Text = ""
    
    mskDataCadastro.Text = Date
    wNumeroClientePedido = 0

End Sub

Private Sub Form_Activate()

    If wNumeroClientePedido <> 0 Then
        Cliente = wNumeroClientePedido
        wPreencherCliente = True
        txtCnpj.Locked = True
    Else
        Cliente = 0
        wPreencherCliente = False
        txtCnpj.Locked = False
    End If

    If wPreencherCliente = True Then
        DescricaoOperacao "Pesquisando Cliente"
        PreencheDadosCliente wNumeroClientePedido
        DescricaoOperacao "Pronto"
    Else
        SQL = "SP_FIN_Pesquisa_Ultimo_Numero_Cliente"
        
        rsNumeroCliente.CursorLocation = adUseClient
        rsNumeroCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        SQL = "SP_FIN_Atualizando_Codigo " & rsNumeroCliente("UltNumCliente")
        
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        If Not rsNumeroCliente.EOF Then
            txtCodigo.Text = rsNumeroCliente("UltNumCliente")
            txtCodigo.Locked = True
            cmbPraca.AddItem "1 - SP"
            cmbPraca.AddItem "2 - Outro"
            cmbPraca.ListIndex = 0
            txtCodigo.Locked = True
            txtRazaoSocial.Enabled = True
            cmbPessoa.Enabled = True
            cmbSituacao.Enabled = True
            txtEndereco.Enabled = True
            txtBairro.Enabled = True
'            txtCnpj.Locked = True
            txtInscricaoEstadual.Locked = True
            txtMunicipio.Enabled = True
            mskCep.Enabled = True
            cmbUF.Enabled = True
            mskDataCadastro.Enabled = False
            mskTelefone.Enabled = True
            mskFax.Enabled = True
            txtEnderecoCobranca.Enabled = True
            txtBairroCobranca.Enabled = True
            txtMunicipioCobranca.Enabled = True
            txtEstadoCobranca.Enabled = True
            mskCepCobranca.Enabled = True
            mskDataNascimento.Enabled = True
            mskCelular.Enabled = True
            cmbRamoAtiv.Enabled = True
            cmbSegmento.Enabled = True
        End If
        rsNumeroCliente.Close
    End If



 tmrTroca.Interval = 1
 validaDados wNumeroClientePedido
 
End Sub

Private Sub Form_Load()
    
lblTituloJanela.Caption = frmCliente.Caption

mskDataCadastro = Date

  Me.top = 4680
  Me.left = 90
  Me.Width = 15180
  Me.Height = 5790
  cmbPessoa.AddItem "FÍSICA"
  cmbPessoa.AddItem "JURÍDICA"
  cmbPessoa.AddItem "FUNCIONÁRIO"
  cmbPessoa.AddItem "ÓRGÃO PÚBLICO"
  grdMunicipio.Height = 990

  cmbTipoCliente.AddItem "L - LOJA"
  cmbTipoCliente.AddItem "F - FATURADO"
  cmbTipoCliente.ListIndex = 0
  
If cmbPessoa.Text = "FÍSICA" Then
  cmbPessoa.TabIndex = 4
End If

  SQL = "SP_FIN_Situacao_Cliente"

    adoCliente.CursorLocation = adUseClient
    adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not adoCliente.EOF Then
      Do While Not adoCliente.EOF
         cmbSituacao.AddItem adoCliente("SI_CodigoSituacao") & " - " & RTrim(LTrim(adoCliente("SI_Descricao")))
         adoCliente.MoveNext
      Loop
      cmbSituacao.ListIndex = 0
    End If

    adoCliente.Close

  
  wPreencheInicio = True
  cmbPessoa.ListIndex = 0

    Dim rsNumeroCliente As New ADODB.Recordset
    Dim preencheUF As Boolean

    SQL = "SP_FIN_Ler_Estado"

    adoCliente.CursorLocation = adUseClient
    adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not adoCliente.EOF Then
        preencheUF = True
    Else
        preencheUF = False
    End If
     
    If preencheUF = True Then
        Do While Not adoCliente.EOF
            cmbUF.AddItem UCase(adoCliente("UF_Estado"))
            adoCliente.MoveNext
        Loop
        For I = 0 To cmbUF.ListCount
            cmbUF.ListIndex = I
            If cmbUF.Text = "SP" Then
                cmbUF.ListIndex = I
                Exit For
            End If
        Next I
    End If
      
    If preencheUF = True Then
        Do While Not adoCliente.EOF
            cmbUF.AddItem UCase(adoCliente("UF_Estado"))
            adoCliente.MoveNext
        Loop
        For I = 0 To cmbUF.ListCount
            cmbUF.ListIndex = I
            If cmbUF.Text = "SP" Then
                cmbUF.ListIndex = I
                Exit For
            End If
        Next I
        
    End If

    
       adoCliente.Close
        
    
End Sub

Function preencheUF() As Boolean

    SQL = "SP_FIN_Ler_Estado"
    adoCliente.CursorLocation = adUseClient
    adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not adoCliente.EOF Then
        preencheUF = True
    Else
        preencheUF = False
    End If
    adoCliente.Close
End Function
Private Sub mskMunicipio_GotFocus()
    grdMunicipio.ZOrder
    grdMunicipio.Visible = True
End Sub
Private Sub mskMunicipio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
       grdMunicipio.SetFocus
    End If
End Sub

Private Sub grdMunicipio_Click()
'    txtMunicipio.Text = UCase(Trim(grdMunicipio.TextMatrix(Row, 0)))
End Sub

Private Sub grdMunicipio_EnterCell()
    txtEstadoCobranca.Text = cmbUF.Text
    txtCodMun.Locked = True
    cmbUF.Locked = True
    cmbPraca.Locked = True
    mskTelefone.SetFocus
    
End Sub

Private Sub grdMunicipio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Or KeyAscii = 27 Then
        grdMunicipio.Visible = False
    End If
    
End Sub

Private Sub grdMunicipio_LostFocus()

If grdMunicipio.Row < 0 Then
 Exit Sub
Else

    txtMunicipio.Text = UCase(grdMunicipio.TextMatrix(grdMunicipio.Row, 0))
    txtMunicipioCobranca.Text = txtMunicipio.Text

    grdMunicipio.Visible = False


    If txtMunicipio.Text = "SÃO PAULO" Then
        cmbPraca.Text = "1 - SP"
        ElseIf txtMunicipio.Text <> "SÃO PAULO" Then

        cmbPraca.Text = "2 - Outro"
    End If
  End If
End Sub
Private Sub grdMunicipio_RowColChange()
   On Error GoTo SaidaRotina

    cmbUF.Text = UCase(grdMunicipio.TextMatrix(grdMunicipio.Row, 2))
    txtCodMun.Text = grdMunicipio.TextMatrix(grdMunicipio.Row, 1)

SaidaRotina:
    Exit Sub
End Sub


Private Sub PreencheGridMunicipio()

    grdMunicipio.Rows = 0
    
    With grdMunicipio
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeSpill
        .Editable = flexEDNone
    End With
 
Ln = 0

    SQL = "SP_FIN_Pesquisa_Municipio"

        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

         Do While Not adoCliente.EOF

                With grdMunicipio
                     .AddItem Trim(adoCliente("Mun_nome")) & Chr(9) & Trim(adoCliente("Mun_Codigo")) & Chr(9) _
                    & Trim(adoCliente("Mun_UF"))
                     .IsSubtotal(.Rows) = True
                     .RowOutlineLevel(.Rows) = 3
                     .Cell(flexcpFontBold, .Rows, 0) = False
                     .Redraw = flexRDBuffered
                End With
        
            adoCliente.MoveNext
            Ln = Ln + 1
         Loop
            Ln = Ln - 1
            Do While Ln > 0
                grdMunicipio.IsCollapsed(Ln) = flexOutlineCollapsed
                Ln = Ln - 1
            Loop
        adoCliente.Close

End Sub

Private Sub PreencheGridMunicipioPesquisa()

If wPreencheInicio = True Then
   Exit Sub
End If

    grdMunicipio.Rows = 0
    
    With grdMunicipio
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeSpill
        .Editable = flexEDNone
    End With


Ln = 0

        If Len(txtMunicipio.Text) > 0 Then
           SQL = "SP_FIN_Ler_Codigo_Municipio_Por_Parametro '" & txtMunicipio.Text & "'"
           
            adoCliente.CursorLocation = adUseClient
            adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
                
         Else
         'adoCliente.Close
         Exit Sub
           
         End If

         Do While Not adoCliente.EOF
           
                With grdMunicipio
                     .AddItem Trim(adoCliente("Mun_nome")) & Chr(9) & Trim(adoCliente("Mun_Codigo")) & Chr(9) _
                    & Trim(adoCliente("Mun_UF"))
                     .IsSubtotal(.Rows - 1) = True
                     .RowOutlineLevel(.Rows - 1) = 3
                     .Cell(flexcpFontBold, .Rows - 1, 0) = False
                     .Redraw = flexRDBuffered
                End With

            adoCliente.MoveNext
            Ln = Ln + 1
         Loop
     
            Ln = Ln - 1
            Do While Ln >= 0
                grdMunicipio.IsCollapsed(Ln) = flexOutlineCollapsed
                Ln = Ln - 1
            Loop
            adoCliente.Close
End Sub

Private Sub carregaCampos()
    If Not adoCliente.EOF Then

        Do While Not adoCliente.EOF
            txtRazaoSocial.Text = adoCliente("CE_Razao")
                           
    adoCliente.MoveNext
        Loop
    Else
        MsgBox "Nenhum registro encontrado!", vbCritical, "ATENÇÃO"
        txtRazaoSocial.SetFocus
        adoCliente.Close
    End If
    
    adoCliente.Close
End Sub

Private Sub mskCelular_Change()
    Numeros (mskCelular.Text)
    mskCelular.Text = wNumeros
End Sub

Private Sub mskCelular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mskTelefone.Text <> "" Then
            ProximoCampo mskFax
            SelecionaCampo mskFax
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo mskTelefone
    End If
End Sub

Private Sub mskCelular_LostFocus()
     If mskCelular.Text <> "" Then
        If Not IsNumeric(mskCelular.Text) Then
            MsgBox "Digite apenas números!", vbCritical, "ATENÇÃO"
            mskCelular.Text = ""
            mskCelular.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub mskCep_Change()
    Numeros (mskCep.Text)
    mskCep.Text = wNumeros
End Sub

Private Sub mskCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mskCep.Text <> "" Then
            If IsNumeric(mskCep.Text) = True Then
                If ConsultaCep(mskCep.Text) = False Then
                    ProximoCampo txtEndereco
                    SelecionaCampo txtEndereco
                Else
                    ProximoCampo txtNumero
                End If
            Else
                MsgBox "Digite apenas números!", vbCritical, "Atenção"
                SelecionaCampo mskCep
            End If
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtInscricaoEstadual
        SelecionaCampo txtInscricaoEstadual
    End If
End Sub

Private Sub mskCep_LostFocus()
    If Len(mskCep.Text) <> 8 Then
            MsgBox "O CEP deve ser informado com 8 posições!", vbCritical, "Atenção"
            mskCep.SelStart = 0
            mskCep.SelLength = Len(mskCep.Text)
            Screen.MousePointer = 0
            'mskCep.SetFocus
            Exit Sub
        End If
       If mskCep.Text <> "" Then
        If IsNumeric(mskCep.Text) = True Then
            
                If ConsultaCep(mskCep.Text) = False Then
                    
                        ProximoCampo txtEndereco
                        SelecionaCampo txtEndereco
                        txtEndereco.Text = ""
                        txtNumero.Text = ""
                        txtBairro.Text = ""
                        txtMunicipio.Text = ""
                        Call preencheUF
                        txtComplemento.Text = ""
                        txtCodMun.Text = ""
                  
                Else
                        ProximoCampo txtEndereco
                        Call ConsultaCep(mskCep.Text)
                End If
         
            
        Else
            MsgBox "Digite apenas números!", vbCritical, "Atenção"
            ProximoCampo mskCep
            SelecionaCampo mskCep
        End If
    End If
    If Len(txtEndereco.Text) > 40 Then
        MsgBox "Limite de caracter alcançado, abrevie o endereço!", vbCritical, "Atenção"
        txtEndereco.SetFocus
    End If
    mskCepCobranca.Text = mskCep.Text
    
    If mskCep.Text = "" Then
        MsgBox "Digite o CEP!", vbCritical, "ATENÇÃO"
        mskCep.SetFocus
        Exit Sub
    End If

    If txtEndereco.Text = "" Then
        
    Else
        txtNumero.SetFocus
    End If
    
End Sub

Private Sub mskCepCobranca_Change()
    Numeros (mskCepCobranca.Text)
    mskCepCobranca.Text = wNumeros
End Sub

Private Sub mskCepCobranca_GotFocus()
       
    mskCepCobranca.SelStart = 0
    mskCepCobranca.SelLength = Len(mskCepCobranca.Text)
    
End Sub

Private Sub mskCepCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mskCepCobranca.Text <> "" Then
            mskCepCobranca.Text = UCase(mskCepCobranca.Text)
            ProximoCampo txtBairroCobranca
            SelecionaCampo txtBairroCobranca
        Else
            mskCepCobranca.Text = UCase(mskCep.Text)
            ProximoCampo txtBairroCobranca
            SelecionaCampo txtBairroCobranca
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtComplCobranca
        SelecionaCampo txtComplCobranca
    End If

End Sub

Private Sub mskDataCadastro_Change()
        If Len(mskDataCadastro.Text) = 2 Then
            mskDataCadastro.Text = mskDataCadastro.Text & "/"
            mskDataCadastro.SelStart = 3
        ElseIf Len(mskDataCadastro.Text) = 5 Then
            mskDataCadastro.Text = mskDataCadastro.Text & "/"
            mskDataCadastro.SelStart = 6
        End If
   
End Sub

Private Sub mskDataCadastro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mskDataCadastro.Text <> "" Then
            ProximoCampo txtEnderecoCobranca
            SelecionaCampo txtEnderecoCobranca
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtEMail
        SelecionaCampo txtEMail
    End If
    
End Sub


Private Sub mskDataNascimento_Change()
    If mskDataNascimento.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        mskDataNascimento.Text = ""
        mskDataNascimento.SetFocus
        Exit Sub
    End If

End Sub

Private Sub mskDataNascimento_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            ProximoCampo cmbRamoAtiv
        ElseIf KeyAscii = 27 Then
            Call Limpar
            Unload Me
            ProximoCampo mskFax
            SelecionaCampo mskFax
        End If
        
        If Len(mskDataNascimento.Text) = 2 Then
            mskDataNascimento.Text = mskDataNascimento.Text & "/"
            mskDataNascimento.SelStart = 3
        ElseIf Len(mskDataNascimento.Text) = 5 Then
            mskDataNascimento.Text = mskDataNascimento.Text & "/"
            mskDataNascimento.SelStart = 6
        End If
           
End Sub

Private Sub mskFax_Change()
    Numeros (mskFax.Text)
    mskFax.Text = wNumeros
End Sub

Private Sub mskFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mskFax.Text <> "" Then
            ProximoCampo mskDataNascimento
            SelecionaCampo mskDataNascimento
        Else
            ProximoCampo mskDataNascimento
            SelecionaCampo mskDataNascimento
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo mskCelular
        SelecionaCampo mskCelular
    End If
End Sub

Private Sub mskTelefone_Change()
    Numeros (mskTelefone.Text)
    mskTelefone.Text = wNumeros
End Sub

Private Sub mskTelefone_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If mskTelefone.Text <> "" Then
            ProximoCampo mskCelular
            SelecionaCampo mskCelular
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo cmbUF
    End If
    
End Sub

'Private Sub tmrTroca_Timer()
'    Call TrocaBannerTopo1
'End Sub

Private Sub txtBairro_Change()
    If txtBairro.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtBairro.Text = ""
        txtBairro.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtBairro.Text <> "" Then
            txtBairro.Text = UCase(txtBairro.Text)
            ProximoCampo txtMunicipio
            SelecionaCampo txtMunicipio
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtComplemento
        SelecionaCampo txtComplemento
    End If
        
End Sub

Private Sub txtBairro_LostFocus()
     txtBairro.Text = UCase(txtBairro.Text)
     txtBairroCobranca.Text = txtBairro.Text
End Sub

Private Sub txtBairroCobranca_Change()
    If txtBairroCobranca.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtBairroCobranca.Text = ""
        txtBairroCobranca.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtBairroCobranca_GotFocus()
    txtBairroCobranca.SelStart = 0
    txtBairroCobranca.SelLength = Len(txtBairroCobranca.Text)
End Sub

Private Sub txtBairroCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtBairroCobranca.Text <> "" Then
            txtBairroCobranca.Text = UCase(txtBairroCobranca.Text)
            ProximoCampo txtMunicipioCobranca
            SelecionaCampo txtMunicipioCobranca
        Else
            txtBairroCobranca.Text = UCase(txtBairro.Text)
            ProximoCampo txtMunicipioCobranca
            SelecionaCampo txtMunicipioCobranca
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo mskCepCobranca
        SelecionaCampo mskCepCobranca
    End If
End Sub

Private Sub txtBairroCobranca_LostFocus()
     txtBairroCobranca.Text = UCase(txtBairroCobranca.Text)
End Sub

Private Sub txtClienteFidelidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call Limpar
        Unload Me
    End If
End Sub

Private Sub txtCNPJ_Change()
    Numeros (txtCnpj.Text)
    txtCnpj.Text = wNumeros
End Sub

Private Sub txtCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCnpj.Text <> "" And IsNumeric(txtCnpj) = True Then
            ProximoCampo txtRazaoSocial
        'Else
            'txtCnpj.SetFocus
            'txtCnpj.SelStart = 0
            'txtCnpj.SelLength = Len(txtCnpj.Text)
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtRazaoSocial
        SelecionaCampo txtRazaoSocial
    End If
End Sub

Private Sub txtCNPJ_LostFocus()
    lblClite(4).ForeColor = &HE0E0E0
    
    If cmbPessoa.Text = "FÍSICA" Or cmbPessoa.Text = "FUNCIONÁRIO" Then
        If Len(txtCnpj.Text) < 11 And txtCnpj.Text <> "" Then
            'MsgBox "Digite todos os dígitos do CPF!", vbCritical, "ATENÇÃO"
            'txtCnpj.SetFocus
            lblClite(4).ForeColor = vbRed
            txtCnpj.SelStart = 0
            txtCnpj.SelLength = Len(txtCnpj.Text)
            Screen.MousePointer = 0
            Exit Sub
        End If
    ElseIf cmbPessoa.Text = "JURÍDICA" Or cmbPessoa.Text = "ÓRGÃO PÚBLICO" Then
        If Len(txtCnpj.Text) < 14 And txtCnpj.Text <> "" Then
            'MsgBox "Digite todos os dígitos do CNPJ!", vbCritical, "ATENÇÃO"
            'txtCnpj.SetFocus
            lblClite(4).ForeColor = vbRed
            txtCnpj.SelStart = 0
            txtCnpj.SelLength = Len(txtCnpj.Text)
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    If txtCnpj.Text <> "" Then
        If Not IsNumeric(txtCnpj.Text) Then
            'MsgBox "Digite apenas números!", vbCritical, "ATENÇÃO"
            lblClite(4).ForeColor = vbRed
            txtCnpj.Text = ""
            txtCnpj.SetFocus
            Exit Sub
        End If
        If VerificaClienteExisteCNPJ(txtCnpj.Text) = True Then
            cmbSituacao.Locked = False
            If txtCodigo.Text <> "" And txtCnpj.Text <> "" Then
                txtCnpj.Locked = False
            End If
        ElseIf VerificaClienteExisteCNPJ(txtCnpj.Text) = False Then
            cmbSituacao.Locked = True
            cmbSituacao.Text = "0 - ATIVO"
        End If
    End If

End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        If txtCodigo.Text <> "" And IsNumeric(txtCodigo.Text) = True Then
            ProximoCampo txtCnpj
            SelecionaCampo txtCnpj
        Else
            ProximoCampo txtCodigo
            SelecionaCampo txtCodigo
        End If
    ElseIf KeyAscii = 27 Then
        Unload Me
        'frmManutencaoCliente.txtCampo.SetFocus
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> "" Then
        If Not IsNumeric(txtCodigo.Text) Then
            MsgBox "Digite apenas números!", vbCritical, " ATENÇÃO"
            txtCodigo.Text = ""
            txtCodigo.SetFocus
            Exit Sub
        End If
        If VerificaClienteExiste(txtCodigo.Text) = True Then
            PreencheDadosCliente txtCodigo.Text
            cmbPessoa.Enabled = True
            
'        Else
'            MsgBox "Código não cadastrado!", vbCritical, "ATENÇÃO"
'            txtCodigo.Text = ""
'            txtCodigo.SetFocus
'            Exit Sub
        End If
    End If
    
End Sub

Private Sub txtCodMun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call Limpar
        Unload Me
    End If
End Sub

Private Sub txtComplCobranca_Change()
    If txtComplCobranca.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtComplCobranca.Text = ""
        txtComplCobranca.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtComplCobranca_GotFocus()
    txtComplCobranca.SelStart = 0
    txtComplCobranca.SelLength = Len(txtComplCobranca.Text)
End Sub

Private Sub txtComplCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call Limpar
        Unload Me
    End If
End Sub

Private Sub txtComplemento_Change()
    If txtComplemento.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtComplemento.Text = ""
        txtComplemento.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtComplemento.Text <> "" Then
            txtComplemento.Text = UCase(txtComplemento.Text)
            ProximoCampo txtBairro
            SelecionaCampo txtBairro
        Else
            txtComplemento.Text = UCase(txtComplemento.Text)
            ProximoCampo txtComplemento
            SelecionaCampo txtComplemento
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtNumero
        SelecionaCampo txtNumero
    End If
End Sub

Private Sub txtComplemento_LostFocus()
     If Len(txtEndereco.Text & txtComplemento.Text) > 40 Then
        MsgBox "Favor abreviar o Complemento ou Endereço do cliente", vbCritical, "Atenção"
        txtComplemento.SetFocus
    Else
        SelecionaCampo txtNumero
    End If
    txtComplCobranca.Text = UCase(txtComplCobranca.Text)
    txtComplCobranca.Text = txtComplemento.Text
End Sub

Private Sub txtEmail_Change()
    If txtEMail.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtEMail.Text = ""
        txtEMail.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskDataCadastro.Text = Format(Date, "yyyy/mm/dd")
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo cmbSegmento
    End If
End Sub

Private Sub txtEndereco_Change()
    If txtEndereco.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtEndereco.Text = ""
        txtEndereco.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEndereco.Text <> "" Then
            txtEndereco.Text = UCase(txtEndereco.Text)
            ProximoCampo txtNumero
            SelecionaCampo txtNumero
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo mskCep
        SelecionaCampo mskCep
    End If
End Sub

Private Sub txtEndereco_LostFocus()
    txtEnderecoCobranca.Text = txtEndereco.Text
End Sub

Private Sub txtEnderecoCobranca_Change()
    If txtEnderecoCobranca.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtEnderecoCobranca.Text = ""
        txtEnderecoCobranca.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEnderecoCobranca_GotFocus()
    txtEnderecoCobranca.SelStart = 0
    txtEnderecoCobranca.SelLength = Len(txtEnderecoCobranca.Text)
End Sub

Private Sub txtEnderecoCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEnderecoCobranca.Text <> "" Then
            txtEnderecoCobranca.Text = UCase(txtEnderecoCobranca.Text)
            ProximoCampo txtNumCobranca
            SelecionaCampo txtNumCobranca
        Else
            txtEnderecoCobranca.Text = UCase(txtEndereco.Text)
            ProximoCampo txtNumCobranca
            SelecionaCampo txtNumCobranca
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    End If
End Sub

Private Sub txtEnderecoCobranca_LostFocus()
     txtEnderecoCobranca.Text = UCase(txtEnderecoCobranca.Text)
End Sub

Private Sub txtEstadoCobranca_Change()
    If txtEstadoCobranca.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtEstadoCobranca.Text = ""
        txtEstadoCobranca.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEstadoCobranca_GotFocus()
    txtEstadoCobranca.SelStart = 0
    txtEstadoCobranca.SelLength = Len(txtEstadoCobranca.Text)
End Sub

Private Sub txtEstadoCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEstadoCobranca.Text <> "" Then
            txtEstadoCobranca.Text = UCase(txtEstadoCobranca.Text)
        Else
            txtEstadoCobranca.Text = UCase(cmbUF.Text)
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtMunicipioCobranca
        SelecionaCampo txtMunicipioCobranca
    End If
End Sub

Private Sub txtEstadoCobranca_LostFocus()
    If IsNumeric(txtEstadoCobranca) Then
        MsgBox "Digite apenas as siglas do estado!", vbCritical, "ATENÇÃO"
        txtEstadoCobranca.Text = ""
        txtEstadoCobranca.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtInscricaoEstadual_Change()
    If txtInscricaoEstadual.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtInscricaoEstadual.Text = ""
        txtInscricaoEstadual.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtInscricaoEstadual_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
      If Mid(UCase(cmbPessoa.Text), 1, 1) = "F" Then
         txtInscricaoEstadual.Text = "ISENTO"
      End If
      ProximoCampo mskCep
      SelecionaCampo mskCep
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo chkCart
    End If
End Sub

Private Sub txtInscricaoEstadual_LostFocus()
    If Mid(UCase(cmbPessoa.Text), 1, 1) = "F" Then
       txtInscricaoEstadual.Text = "ISENTO"
    End If
    lblClite(2).ForeColor = &HE0E0E0
End Sub




''Private Sub txtLimiteCredito_LostFocus()
''
''    If txtLimiteCredito.Text = "'" Then
''        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
''        txtLimiteCredito.Text = ""
''        txtLimiteCredito.SetFocus
''        Exit Sub
''    End If
''    txtLimiteCredito.Text = Format(txtLimiteCredito.Text, "####,###,##.00")
''End Sub

Private Sub txtMunicipio_Change()
        If txtMunicipio.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtMunicipio.Text = ""
        txtMunicipio.SetFocus
        Exit Sub
    End If
' Comando para agilizar o andamento do botão limpar
    If wLimpar = True Then
        wLimpar = False
        Exit Sub
    End If
'--------------------------------------------------
    
    grdMunicipio.ZOrder
    If wPreencheInicio = False Then
       grdMunicipio.Visible = True
       PreencheGridMunicipioPesquisa
    End If
    If Trim(txtMunicipio.Text) = "" Then
       grdMunicipio.Visible = False
    End If
    

    cmbPraca.AddItem "1 - SP"
    cmbPraca.AddItem "2 - Outro"
    If txtMunicipio.Text = "SÃO PAULO" Then
        cmbPraca.Text = "1 - SP"
    ElseIf txtMunicipio.Text <> "SÃO PAULO" Then
        cmbPraca.Text = "2 - Outro"
    End If
    txtMunicipioCobranca.Text = txtMunicipio.Text
    
    

End Sub

Private Sub txtMunicipio_GotFocus()
If (grdMunicipio.TextMatrix(2, 1) <> "") Then
    grdMunicipio.ZOrder
    grdMunicipio.Visible = True
End If
End Sub

Private Sub txtMunicipio_KeyDown(KeyCode As Integer, Shift As Integer)
   ' If KeyCode = 40 Then
     '   grdMunicipio.SetFocus
    'End If

End Sub

Private Sub txtMunicipio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Exit Sub
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtBairro
        SelecionaCampo txtBairro
    End If
    wPreencheInicio = False
End Sub



Private Sub txtMunicipioCobranca_Change()
 If txtMunicipioCobranca.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtMunicipioCobranca.Text = ""
        txtMunicipioCobranca.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtMunicipioCobranca_GotFocus()
    txtMunicipioCobranca.SelStart = 0
    txtMunicipioCobranca.SelLength = Len(txtMunicipioCobranca.Text)
End Sub

Private Sub txtMunicipioCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMunicipioCobranca.Text <> "" Then
            txtMunicipioCobranca.Text = UCase(txtMunicipioCobranca.Text)
            ProximoCampo txtEstadoCobranca
            SelecionaCampo txtEstadoCobranca
        Else
            txtMunicipioCobranca.Text = UCase(txtMunicipio.Text)
            ProximoCampo txtEstadoCobranca
            SelecionaCampo txtEstadoCobranca
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtBairroCobranca
        SelecionaCampo txtBairroCobranca
    End If
End Sub

Private Sub txtMunicipioCobranca_LostFocus()
    'txtMunicipioCobranca.Text = UCase(txtMunicipioCobranca.Text)
End Sub

Private Sub txtNumCobranca_Change()
    Numeros (txtNumCobranca.Text)
    txtNumCobranca.Text = wNumeros
End Sub

Private Sub txtNumCobranca_GotFocus()
    txtNumCobranca.SelStart = 0
    txtNumCobranca.SelLength = Len(txtNumCobranca.Text)
End Sub

Private Sub txtNumCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNumCobranca.Text <> "" Then
            txtNumCobranca.Text = UCase(txtNumCobranca.Text)
            ProximoCampo txtComplCobranca
            SelecionaCampo txtComplCobranca
        Else
            txtNumCobranca.Text = UCase(txtNumero.Text)
            ProximoCampo txtComplCobranca
            SelecionaCampo txtComplCobranca
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtEnderecoCobranca
        SelecionaCampo txtEnderecoCobranca
    End If
End Sub

Private Sub txtNumCobranca_LostFocus()
     txtNumCobranca.Text = UCase(txtNumCobranca.Text)
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If txtNumero.Text <> "" Then
            txtNumero.Text = UCase(txtNumero.Text)
            ProximoCampo txtComplemento
            SelecionaCampo txtComplemento
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo txtEndereco
        SelecionaCampo txtEndereco
    End If
End Sub

Private Sub txtNumero_LostFocus()
    lblNumero.ForeColor = &HE0E0E0
    txtNumCobranca.Text = txtNumero.Text
    If Not IsNumeric(txtNumero.Text) Then
        MsgBox "Digite somente números!", vbCritical, "ATENÇÃO"
        txtNumero.Text = ""
        'txtNumero.SetFocus
    End If
End Sub

Private Sub txtRazaoSocial_Change()
    If txtRazaoSocial.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtRazaoSocial.Text = ""
        txtRazaoSocial.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtRazaoSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If txtCodigo.Locked = False Then
            Call Limpar
            Unload Me
 '           frmManutencaoCliente.txtCampo.SetFocus
        Else
            ProximoCampo txtCodigo
            SelecionaCampo txtCodigo
        End If
    ElseIf KeyCode = 192 Then
        If Len(txtRazaoSocial.Text) = 0 Then
            txtRazaoSocial.Text = ""
        Else
            txtRazaoSocial.Text = Mid(txtRazaoSocial.Text, 1, Len(txtRazaoSocial.Text) - 1)
        End If
    End If
End Sub

Function ProximoCampo(ByRef NomeProxCampo)
    On Error Resume Next
    NomeProxCampo.SetFocus
End Function

Function SelecionaCampo(ByRef NomeCampo)
    NomeCampo.SelStart = 0
    NomeCampo.SelLength = Len(NomeCampo.Text)
End Function

Function VerificaClienteExiste(ByVal Cliente As Double) As Boolean
    SQL = "SP_FIN_Pesquisa_Codigo " & Cliente & ""
    
    adoCliente.CursorLocation = adUseClient
    adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not adoCliente.EOF Then
        VerificaClienteExiste = True
    Else
        VerificaClienteExiste = False
    End If
    adoCliente.Close
    
End Function
Function VerificaClienteExisteCNPJ(ByVal Cliente As Double) As Boolean
    SQL = "SP_FIN_Ler_Clientes_Por_Parametro_Cnpj'" & Cliente & "'"
    
    adoCliente.CursorLocation = adUseClient
    adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not adoCliente.EOF Then
        VerificaClienteExisteCNPJ = True
    Else
        VerificaClienteExisteCNPJ = False
        
    End If

    adoCliente.Close
End Function
    

Private Sub txtRazaoSocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRazaoSocial.Text <> "" Then
            ProximoCampo cmbPessoa
            SelecionaCampo cmbPessoa
            txtRazaoSocial.Text = UCase(txtRazaoSocial.Text)
        End If
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
    End If
End Sub

Function PreencheDadosCliente(ByVal Cliente As String)
    If PesquisaCliente(1, Cliente) = True Then
        If (adoCliente("CE_PagamentoCarteira")) = "S" Then
            chkCart.Value = 1
        Else
            chkCart.Value = 0
        End If
        txtCodigo.Text = adoCliente("CE_CodigoCliente")
        txtCodigo.Locked = False
        
        txtRazaoSocial.Text = adoCliente("CE_Razao")
        
        If Trim(adoCliente("CE_TipoPessoa")) = "J" Then
            cmbPessoa.Text = "JURÍDICA"
            mskDataNascimento.Enabled = False
        ElseIf Trim(adoCliente("CE_TipoPessoa")) = "F" Then
            cmbPessoa.Text = "FÍSICA"
        ElseIf Trim(adoCliente("CE_TipoPessoa")) = "U" Then
            cmbPessoa.Text = "FUNCIONÁRIO"
        ElseIf Trim(adoCliente("CE_TipoPessoa")) = "O" Then
            cmbPessoa.Text = "ÓRGÃO PÚBLICO"
            cmbPessoa.Enabled = False
            
        End If
        
        If Trim(adoCliente("CE_Situacao")) = 0 Then
            cmbSituacao.Text = "ATIVO"
        ElseIf Trim(adoCliente("CE_Situacao")) = 1 Then
            cmbSituacao.Text = "EM COB.JUDICIAL"
        ElseIf Trim(adoCliente("CE_Situacao")) = 2 Then
            cmbSituacao.Text = "CONCORDATA REQUERIDA"
        ElseIf Trim(adoCliente("CE_Situacao")) = 3 Then
        cmbSituacao.Text = "ALERTA DA ACESP"
        ElseIf Trim(adoCliente("CE_Situacao")) = 4 Then
            cmbSituacao.Text = "PAGA EM CARTORIO"
        ElseIf Trim(adoCliente("CE_Situacao")) = 5 Then
            cmbSituacao.Text = "FALÊNCIA REQUERIDA"
        ElseIf Trim(adoCliente("CE_Situacao")) = 6 Then
            cmbSituacao.Text = "ALERTA DE MEO"
        ElseIf Trim(adoCliente("CE_Situacao")) = 7 Then
            cmbSituacao.Text = "TEM PROTESTOS"
        End If
        
        
      
        
        txtNumero.Text = IIf(IsNull(adoCliente("CE_Numero")), "", adoCliente("CE_Numero"))
        txtComplemento.Text = IIf(IsNull(adoCliente("CE_Complemento")), "", adoCliente("CE_Complemento"))
        txtEndereco.Text = adoCliente("CE_Endereco")
        txtBairro.Text = adoCliente("CE_Bairro")
        txtCnpj.Text = adoCliente("CE_CGC")
        txtCnpj.Locked = False
        txtInscricaoEstadual.Text = adoCliente("CE_InscricaoEstadual")
        txtMunicipio.Text = adoCliente("CE_Municipio")
        txtEMail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
       
        If Trim(adoCliente("CE_TipoPessoa")) = "F" Then
           txtInscricaoEstadual.Locked = True
        End If

        SQL = "Exec SP_FIN_Ler_Municipio '" & adoCliente("CE_CodigoMunicipio") & "'"

        adoMunicipio.CursorLocation = adUseClient
        adoMunicipio.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        If adoMunicipio.EOF = False Then
           txtMunicipio.Text = Trim(adoMunicipio("Mun_Nome"))
        End If
        
        cmbPraca.AddItem "1 - SP"
        cmbPraca.AddItem "2 - Outro"
        For I = 0 To cmbPraca.ListCount - 1
            cmbPraca.ListIndex = I
            If Val(Mid(cmbPraca.Text, 1, 1)) = adoCliente("CE_Praca") Then
                cmbPraca.ListIndex = I
                Exit For
            End If
        Next I
        mskCep.Text = adoCliente("CE_Cep")
        For I = 0 To cmbUF.ListCount
            cmbUF.ListIndex = I
            If cmbUF.Text = UCase(adoCliente("CE_Estado")) Then
                auxCMBUF = UCase(adoCliente("CE_Estado"))
                cmbUF.ListIndex = I
                Exit For
            End If
        Next I
        mskDataCadastro.Text = Format(adoCliente("CE_DataCadastro"), "yyyy/mm/dd")
        mskTelefone.Text = adoCliente("CE_Telefone")
        mskFax.Text = adoCliente("CE_Fax")
        txtEnderecoCobranca.Text = adoCliente("CE_EnderecoCobranca")
        txtNumCobranca = IIf(IsNull(adoCliente("CE_NumeroCobranca")), "", adoCliente("CE_NumeroCobranca"))
        txtComplCobranca = IIf(IsNull(adoCliente("CE_ComplCobranca")), "", adoCliente("CE_ComplCobranca"))
        txtBairroCobranca.Text = adoCliente("CE_BairroCobranca")
        txtMunicipioCobranca.Text = adoCliente("CE_MunicipioCobranca")
        txtEstadoCobranca.Text = adoCliente("CE_EstadoCobranca")
        mskCepCobranca.Text = adoCliente("CE_CepCobranca")
 '       txtLimiteCredito.Text = Format(adoCliente("CE_LimiteCredito"), "####,###,##.00")
 '       mskDataLimite.Text = adoCliente("CE_DataLimiteCredito")
        
        If IsNull(adoCliente("CE_CodigoMunicipio")) = False Then
           txtCodMun.Text = adoCliente("CE_CodigoMunicipio")
          ' adoCliente.Close
        End If
 
        txtEMail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
        mskDataNascimento.Enabled = True
        mskDataNascimento.Text = IIf(IsNull(adoCliente("CE_DataNasc")), "01/01/1900", adoCliente("CE_DataNasc"))
        txtClienteFidelidade.Text = IIf(IsNull(adoCliente("ce_clienteFidelidade")), "", adoCliente("ce_clienteFidelidade"))
        mskCelular.Text = IIf(IsNull(adoCliente("CE_Celular")), "00000000", adoCliente("CE_Celular"))
        Call CarregaRamo(cmbPessoa.Text, txtCodigo.Text)
        txtCodigo.Locked = True
        txtRazaoSocial.SelStart = 0
        txtRazaoSocial.SelLength = Len(txtRazaoSocial.Text)
        adoCliente.Close
        adoMunicipio.Close
    End If
End Function

Function PreencheDadosClienteCNPJ(ByVal Cliente As String)
    If PesquisaCliente(2, Cliente) = True Then
        If (adoCliente("CE_PagamentoCarteira")) = "S" Then
            chkCart.Value = 1
        Else
            chkCart.Value = 0
        End If
        txtCodigo.Text = adoCliente("CE_CodigoCliente")
        txtCodigo.Locked = False
        
        txtRazaoSocial.Text = adoCliente("CE_Razao")
        
        If Trim(adoCliente("CE_TipoPessoa")) = "J" Then
            cmbPessoa.Text = "JURÍDICA"
        ElseIf Trim(adoCliente("CE_TipoPessoa")) = "F" Then
            cmbPessoa.Text = "FÍSICA"
        ElseIf Trim(adoCliente("CE_TipoPessoa")) = "U" Then
            cmbPessoa.Text = "FUNCIONÁRIO"
        ElseIf Trim(adoCliente("CE_TipoPessoa")) = "O" Then
            cmbPessoa.Text = "ÓRGÃO PÚBLICO"
        End If
        
        If Trim(adoCliente("CE_Situacao")) = 0 Then
            cmbSituacao.Text = "ATIVO"
        ElseIf Trim(adoCliente("CE_Situacao")) = 1 Then
            cmbSituacao.Text = "EM COB.JUDICIAL"
        ElseIf Trim(adoCliente("CE_Situacao")) = 2 Then
            cmbSituacao.Text = "CONCORDATA REQUERIDA"
        ElseIf Trim(adoCliente("CE_Situacao")) = 3 Then
        cmbSituacao.Text = "ALERTA DA ACESP"
        ElseIf Trim(adoCliente("CE_Situacao")) = 4 Then
            cmbSituacao.Text = "PAGA EM CARTORIO"
        ElseIf Trim(adoCliente("CE_Situacao")) = 5 Then
            cmbSituacao.Text = "FALÊNCIA REQUERIDA"
        ElseIf Trim(adoCliente("CE_Situacao")) = 6 Then
            cmbSituacao.Text = "ALERTA DE MEO"
        ElseIf Trim(adoCliente("CE_Situacao")) = 7 Then
            cmbSituacao.Text = "TEM PROTESTOS"
        End If

        txtNumero.Text = IIf(IsNull(adoCliente("CE_Numero")), "", adoCliente("CE_Numero"))
        txtComplemento.Text = IIf(IsNull(adoCliente("CE_Complemento")), "", adoCliente("CE_Complemento"))
        txtEndereco.Text = adoCliente("CE_Endereco")
        txtBairro.Text = adoCliente("CE_Bairro")
        txtCnpj.Text = adoCliente("CE_CGC")
        txtInscricaoEstadual.Text = adoCliente("CE_InscricaoEstadual")
        txtMunicipio.Text = adoCliente("CE_Municipio")
        txtEMail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))

        If Trim(adoCliente("CE_TipoPessoa")) = "F" Then
           txtInscricaoEstadual.Locked = True
        End If

        SQL = "Exec SP_FIN_Ler_Municipio '" & adoCliente("CE_CodigoMunicipio") & "'"

        adoMunicipio.CursorLocation = adUseClient
        adoMunicipio.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

        If adoMunicipio.EOF = False Then
           txtMunicipio.Text = Trim(adoMunicipio("Mun_Nome"))
        End If
        
        'cmbPraca.AddItem "1 - SP"
        'cmbPraca.AddItem "2 - Outro"
        For I = 0 To cmbPraca.ListCount - 1
            cmbPraca.ListIndex = I
            If Val(Mid(cmbPraca.Text, 1, 1)) = adoCliente("CE_Praca") Then
                cmbPraca.ListIndex = I
                Exit For
            End If
        Next I
        mskCep.Text = adoCliente("CE_Cep")
        For I = 0 To cmbUF.ListCount
            cmbUF.ListIndex = I
            If cmbUF.Text = UCase(adoCliente("CE_Estado")) Then
                auxCMBUF = UCase(adoCliente("CE_Estado"))
                cmbUF.ListIndex = I
                Exit For
            End If
        Next I
        mskDataCadastro.Text = Format(adoCliente("CE_DataCadastro"), "yyyy/mm/dd")
        mskTelefone.Text = adoCliente("CE_Telefone")
        mskFax.Text = adoCliente("CE_Fax")
        txtEnderecoCobranca.Text = adoCliente("CE_EnderecoCobranca")
        txtNumCobranca = IIf(IsNull(adoCliente("CE_NumeroCobranca")), "", adoCliente("CE_NumeroCobranca"))
        txtComplCobranca = IIf(IsNull(adoCliente("CE_ComplCobranca")), "", adoCliente("CE_ComplCobranca"))
        txtBairroCobranca.Text = adoCliente("CE_BairroCobranca")
        txtMunicipioCobranca.Text = adoCliente("CE_MunicipioCobranca")
        txtEstadoCobranca.Text = adoCliente("CE_EstadoCobranca")
        mskCepCobranca.Text = adoCliente("CE_CepCobranca")
'        txtLimiteCredito.Text = Format(adoCliente("CE_LimiteCredito"), "####,###,##.00")
'        mskDataLimite.Text = adoCliente("CE_DataLimiteCredito")
        
        If IsNull(adoCliente("CE_CodigoMunicipio")) = False Then
           txtCodMun.Text = adoCliente("CE_CodigoMunicipio")
        End If
        txtClienteFidelidade.Text = IIf(IsNull(adoCliente("ce_codigoFidelidade")), "", adoCliente("ce_codigoFidelidade"))
        txtEMail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
        mskDataNascimento.Enabled = True
        mskDataNascimento.Text = IIf(IsNull(adoCliente("CE_DataNasc")), "01/01/1900", adoCliente("CE_DataNasc"))
        mskCelular.Text = (adoCliente("CE_Celular"))
        Call CarregaRamo(cmbPessoa.Text, txtCodigo.Text)
        txtCodigo.Locked = True
        txtRazaoSocial.SelStart = 0
        txtRazaoSocial.SelLength = Len(txtRazaoSocial.Text)
        adoCliente.Close
        adoMunicipio.Close
    End If

    'adoCliente.Close
'        adoMunicipio.Close
End Function


Function PesquisaCliente(ByVal tipoPesquisa As Integer, ByVal Cliente As String) As Boolean
    If tipoPesquisa = 1 Then
        SQL = "SP_FIN_Pesquisa_Codigo_Cliente " & Cliente & ""
        
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
      
    
'Pesquisa por cgc ou cpf (2)

    ElseIf tipoPesquisa = 2 Then
        SQL = "SP_FIN_Ler_Clientes_Por_Parametro_Cnpj '" & Cliente & "'"
        
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
'Pesquisa Pelo Nome Cliente (3)

    ElseIf tipoPesquisa = 3 Then
        SQL = "SP_FIN_Pesquisa_Razao_Cliente '" & Cliente & "'"
    
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
'Pesquisa Cliente Tela frmCadCliente(4)

    ElseIf tipoPesquisa = 4 Then
        SQL = "SP_FIN_Ler_Clientes_Por_Código " & Cliente & ""
        
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

    Else
        Exit Function
    End If
    If Not adoCliente.EOF Then
        PesquisaCliente = True
    Else
        PesquisaCliente = False
    End If
        
End Function

Function CarregaRamo(ByVal Pessoa As String, ByVal codigo As String)
    Dim adoRamo As New ADODB.Recordset
    Dim adoRamo2 As New ADODB.Recordset
    Dim Index As Integer
    Index = 0

    cmbRamoAtiv.Clear
    
    If Pessoa = "Jurídica" Or Pessoa = "JURÍDICA" Then
        Pessoa = "J"
    ElseIf Pessoa = "Física" Or Pessoa = "FÍSICA" Then
        Pessoa = "F"
    ElseIf Pessoa = "Funcionário" Or Pessoa = "FUNCIONÁRIO" Then
        Pessoa = "U"
    ElseIf Pessoa = "Órgão Público" Or Pessoa = "ÓRGÃO PÚBLICO" Then
        Pessoa = "O"
    End If
    
    SQL = "SP_FIN_Pesquisa_Ramo_Atividade_Por_Pessoa '" & Pessoa & "'"

    adoRamo.CursorLocation = adUseClient
    adoRamo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
     
    SQL = "SP_FIN_Pesquisa_Ramo_Atividade_Por_Codigo '" & codigo & "'"
                      
    adoRamo2.CursorLocation = adUseClient
    adoRamo2.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
   
    If Not adoRamo.EOF Then
        Do While Not adoRamo.EOF
              cmbRamoAtiv.AddItem adoRamo("RMO_Codigo") & " - " & RTrim(LTrim(adoRamo("RMO_DescricaoRamo")))
              If Not adoRamo2.EOF Then
                If adoRamo("RMO_Codigo") = adoRamo2("ce_ramoatividade") Then
                   cmbRamoAtiv.ListIndex = Index
                End If
              End If
              If cmbRamoAtiv.Text = "" Then
                  If RTrim(LTrim(adoRamo("RMO_DescricaoRamo"))) = "Indefinido" Or RTrim(LTrim(adoRamo("RMO_DescricaoRamo"))) = "Orgao Publico" Then
                     cmbRamoAtiv.ListIndex = Index
                  End If
              End If
              Index = Index + 1
              adoRamo.MoveNext
        Loop

    Call carregaSegmento(Mid(cmbRamoAtiv.Text, 1, 2), txtCodigo.Text)
    End If
    adoRamo.Close
    adoRamo2.Close
    
End Function

Function carregaSegmento(ByVal RamoAtividade As String, ByVal codigoCliente As String)
    Dim adoSegmento As New ADODB.Recordset
    Dim adoSegmento2 As New ADODB.Recordset
    Dim Index As Integer
    Index = 0
    
    cmbSegmento.Clear
    
    SQL = "SP_FIN_Pesquisa_Segmento_Por_Ramo_Atividade '" & RamoAtividade & "'"
    
    adoSegmento.CursorLocation = adUseClient
    adoSegmento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
    SQL = "SP_FIN_Pesquisa_Segmento_Por_Codigo_Cliente '" & codigoCliente & "'"
    
    adoSegmento2.CursorLocation = adUseClient
    adoSegmento2.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not adoSegmento.EOF Then
        Do While Not adoSegmento.EOF
            cmbSegmento.AddItem adoSegmento("SEG_CodigoSegmento") & " - " & adoSegmento("SEG_Descricao")
            If Not adoSegmento2.EOF Then
               If adoSegmento("SEG_CodigoSegmento") = adoSegmento2("ce_Segmento") And _
                  adoSegmento("SEG_RamoAtividade") = adoSegmento2("ce_RamoAtividade") Then
                    cmbSegmento.ListIndex = Index
               End If
            End If
            If cmbSegmento.Text = "" Then
               If RTrim(LTrim(adoSegmento("SEG_descricao"))) = "Indefinido" Then
                 cmbSegmento.ListIndex = Index
               End If
            End If
            Index = Index + 1
            adoSegmento.MoveNext
        Loop
    End If
    adoSegmento.Close
    adoSegmento2.Close
    adoSegmento2.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

End Function

Private Sub cmbSituacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ProximoCampo chkCart
        cmbSituacao.Text = UCase(cmbSituacao.Text)
        cmbSituacao.Text = 0
    ElseIf KeyAscii = 27 Then
        Call Limpar
        Unload Me
        ProximoCampo cmbPessoa
        SelecionaCampo cmbPessoa
    End If
End Sub

Function ConsultaCep(ByVal Cep As String) As Boolean
Dim adoCodigo As New ADODB.Recordset
    
    
    SQL = " SP_FIN_Pesquisa_Cep '" & Cep & "'"
    
    adoCliente.CursorLocation = adUseClient
    adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not adoCliente.EOF Then
        ConsultaCep = True
        txtEndereco.Text = adoCliente("LOGRADOURO")
        txtBairro.Text = adoCliente("BAIRRO")
        txtMunicipio.Text = adoCliente("MUNICIPIO")
        txtMunicipioCobranca.Text = adoCliente("MUNICIPIO")
        txtBairroCobranca.Text = adoCliente("BAIRRO")
        For I = 0 To cmbUF.ListCount
            cmbUF.ListIndex = I
            If cmbUF.Text = UCase(adoCliente("UF")) Then
                cmbUF.ListIndex = I
                Exit For
            End If
        Next I
        txtEstadoCobranca.Text = adoCliente("UF")
        
    Else
        ConsultaCep = False
    End If
    
    adoCNLoja.BeginTrans
    
    SQL = "SP_FIN_Pesquisa_Municipio_Por_Parametro '" & txtMunicipio.Text & "'"
     
        adoCodigo.CursorLocation = adUseClient
        adoCodigo.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    adoCNLoja.Execute (SQL)
    adoCNLoja.CommitTrans
         
     If Not adoCodigo.EOF Then
        
        'MsgBox "Banco já cadastrado!", vbCritical, "ATENÇÃO"
        txtCodMun.Text = adoCodigo("Mun_Codigo")
        cmbUF.Text = adoCodigo("Mun_UF")
        
        mskTelefone.SetFocus
        adoCodigo.Close
        
     End If
    
    'adoCodigo.Close
    
    
    adoCliente.Close
End Function

Function verificaCamposNulos() As Boolean
    verificaCamposNulos = True
    
    If txtRazaoSocial.Text = "" Then
        lblClite(4).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtRazaoSocial
        SelecionaCampo txtRazaoSocial
        
    ElseIf cmbPessoa.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo cmbPessoa
        SelecionaCampo cmbPessoa
        
    ElseIf txtEndereco.Text = "" Then
        lblClite(7).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtEndereco
        SelecionaCampo txtEndereco
        
    ElseIf txtNumero.Text = "" Then
        lblNumero.ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtNumero
        SelecionaCampo txtNumero
        
    ElseIf txtBairro.Text = "" Then
        lblClite(8).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtBairro
        SelecionaCampo txtBairro
        
    ElseIf txtCnpj.Text = "" Then
        lblClite(4).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtCnpj
        SelecionaCampo txtCnpj
        
    ElseIf txtInscricaoEstadual.Text = "" Then
        lblClite(6).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtInscricaoEstadual
        SelecionaCampo txtInscricaoEstadual
        
    ElseIf txtMunicipio.Text = "" Then
        lblClite(3).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo txtMunicipio
        SelecionaCampo txtMunicipio
        
    ElseIf cmbPraca.Text = "" Then
        'lblClite(4).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo cmbPraca
        
    ElseIf cmbUF.Text = "" Then
        'lblClite(4).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo cmbUF
        
    ElseIf mskCep.Text = "" Then
        lblCep.ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo mskCep
        SelecionaCampo mskCep
        
    ElseIf mskTelefone.Text = "" Then
        lblClite(13).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo mskTelefone
        SelecionaCampo mskTelefone
        
    ElseIf IsDate(mskDataCadastro.Text) = False Then
        'lblClite(4).ForeColor = vbRed
        verificaCamposNulos = False
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
        
    ElseIf txtEstadoCobranca.Text = "" Then
        'lblClite(4).ForeColor = vbRed
        txtEstadoCobranca.Text = cmbUF.Text
        verificaCamposNulos = True
        
    ElseIf mskCepCobranca.Text = "" Then
        'lblClite(4).ForeColor = vbRed
        mskCepCobranca.Text = mskCep.Text
        verificaCamposNulos = True
        
    ElseIf IsDate(mskDataNascimento.Text) = False Then
        'lblClite(4).ForeColor = vbRed
        verificaCamposNulos = True
        ProximoCampo mskDataNascimento
    End If
    
    If verificaCamposNulos = False Then
        MsgBox "Preencha todos os campos obrigatórios", vbInformation, "Atenção"
        Exit Function
    End If
    
End Function

Function AtualizaCliente(ByVal codigo As Double) As Boolean
    Dim Pessoa As String
        
    AtualizaCliente = True

    If Trim(mskCep.Text) = "0" Or Trim(mskCep.Text) = "" Then
        lblCep.ForeColor = vbRed
        mskCep.SelStart = 0
        mskCep.SelLength = Len(mskCep.Text)
        Screen.MousePointer = 0
        AtualizaCliente = False
    End If

    If Trim(txtCnpj.Text) <> "" Then

        If Len(txtCnpj.Text) = 11 And UCase(Mid(cmbPessoa.Text, 1, 1)) <> "F" Then
              lblClite(4).ForeColor = vbRed
              'txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              AtualizaCliente = False
        End If
        
        If Len(txtCnpj.Text) = 14 And UCase(Mid(cmbPessoa.Text, 1, 1)) = "F" Then
              lblClite(4).ForeColor = vbRed
              'txtCnpj.Locked = True
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              AtualizaCliente = False
        End If
    
        If Len(txtCnpj.Text) = 11 Then
           If FU_ValidaCPF(txtCnpj.Text) = False Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.Locked = True
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              AtualizaCliente = False
           End If
        End If

        If Len(txtCnpj.Text) = 14 Then
           If FU_ValidaCGC(txtCnpj.Text) = False Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.Locked = True
              txtCnpj.SetFocus

              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
           AtualizaCliente = False
           End If
        Else

        End If
      End If

      If Trim(txtCnpj.Text) = "" Then
         txtCnpj.Text = right(String(14, "0") & txtCnpj.Text, 15)
      End If
      If txtCnpj.Text = 0 Then
         txtCnpj.Text = right(String(14, "0") & txtCnpj.Text, 15)
      End If
    
    If cmbPessoa.Text = "FÍSICA" Then
        Pessoa = "F"
    ElseIf cmbPessoa.Text = "JURÍDICA" Then
        Pessoa = "J"
    ElseIf Trim(UCase(cmbPessoa.Text)) = "FUNCIONÁRIO" Then
        Pessoa = "U"
    ElseIf Trim(UCase(cmbPessoa.Text)) = "ÓRGÃO PÚBLICO" Then
        Pessoa = "O"
    End If

    If txtNumero.Text = "" Then
        txtNumero.Text = 0
    End If
    If txtNumCobranca.Text = "" Then
        txtNumCobranca.Text = 0
    End If
    
    If chkCart.Value = True Then
        pagamentoCarteira = "S"
        Else
        pagamentoCarteira = "N"
    End If

    If Trim(UCase(cmbPessoa.Text)) <> "FÍSICA" Then
       If IsNumeric(txtInscricaoEstadual.Text) = False Then
          If Trim(UCase(txtInscricaoEstadual.Text)) <> "ISENTO" Then
             MsgBox "Inscrição Estadual inválida!", vbCritical, "Atenção"
             txtInscricaoEstadual.SetFocus
             Screen.MousePointer = 0
             AtualizaCliente = False
          End If
       End If
    End If

          If Trim(txtNumero.Text) = "" Or Trim(txtNumero.Text) = 0 Then
             lblNumero.ForeColor = vbRed
             txtNumero.SetFocus
             txtNumero.SelStart = 0
             txtNumero.SelLength = Len(txtNumero.Text)
             Screen.MousePointer = 0
             AtualizaCliente = False
          End If

    If Trim(txtCodMun.Text) = "" Then
       lblClite(3).ForeColor = vbRed
       txtNumero.SetFocus
       txtCodMun.SelStart = 0
       txtCodMun.SelLength = Len(mskCep.Text)
       Screen.MousePointer = 0
       AtualizaCliente = False
    End If

    If IsNumeric(mskTelefone.Text) = False Or Trim(mskTelefone.Text) = "" Then
       lblClite(13).ForeColor = vbRed
       mskTelefone.SetFocus
       Screen.MousePointer = 0
       AtualizaCliente = False
    End If
    


    Numeros (mskFax.Text)
    mskFax.Text = wNumeros
    
    dataNascimento = Format(mskDataNascimento.Text, "yyyy/mm/dd")

    If AtualizaCliente = True Then

        Err.Number = 0
        adoCNLoja.BeginTrans
        SQL = "SP_FIN_Altera_Cliente " & codigo & ",'" & txtRazaoSocial.Text & "','" & txtCnpj.Text & "','" _
                                        & Mid(Pessoa, 1, 2) & "', " & Val(cmbSituacao.Text) & ",'" _
                                        & pagamentoCarteira & "', '" _
                                        & txtInscricaoEstadual.Text & "','" & mskCep.Text & "','" _
                                        & txtEndereco.Text & "','" & txtNumero.Text & "','" _
                                        & txtMunicipio.Text & "'," & txtCodMun.Text & ",'" _
                                        & cmbUF.Text & "','" & txtComplemento.Text & "','" _
                                        & txtBairro.Text & "'," & Mid(cmbPraca.Text, 1, 1) & ",'" _
                                        & mskTelefone.Text & "','" & mskCelular.Text & "','" _
                                        & mskFax.Text & "','" & dataNascimento & "'," _
                                        & Val(cmbRamoAtiv.Text) & ",'" & txtEMail.Text & "'," _
                                        & Val(cmbSegmento.Text) & ", '" & txtEnderecoCobranca.Text & "','" _
                                        & txtNumCobranca.Text & "','" & txtComplCobranca.Text & "','" _
                                        & mskCepCobranca.Text & "','" & txtBairroCobranca.Text & "','" _
                                        & txtMunicipioCobranca.Text & "','" & txtEstadoCobranca.Text & "','0.00'"
                                        

            adoCliente.CursorLocation = adUseClient
            adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
         If Err.Number = 0 Then
            adoCNLoja.CommitTrans
            Screen.MousePointer = 0
            MsgBox "Cliente alterado com sucesso!", vbInformation, "Sucesso"
            AtualizaCliente = True
            txtCodigo.Locked = False
 
        Else
            adoCNLoja.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "Erro no processo de alteração!", vbCritical, "ERRO"
        End If
        Unload Me
    Else
        MsgBox "Informações invalidas"
    End If
    
End Function

Function GravaCliente() As Boolean
    Dim Pessoa As String
        
    GravaCliente = True
    cmbPessoa.Enabled = True
    
    If Trim(mskCep.Text) = "0" Or Trim(mskCep.Text) = "" Then
        lblCep.ForeColor = vbRed
        mskCep.SelStart = 0
        mskCep.SelLength = Len(mskCep.Text)
        GravaCliente = False
    End If



    If Trim(txtCnpj.Text) <> "" Then

        If Len(txtCnpj.Text) = 11 And UCase(Mid(cmbPessoa.Text, 1, 1)) <> "F" Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              GravaCliente = False
        End If
        
        If Len(txtCnpj.Text) = 14 And UCase(Mid(cmbPessoa.Text, 1, 1)) = "F" Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              GravaCliente = False
        End If
    
        If Len(txtCnpj.Text) = 11 Then
           If FU_ValidaCPF(txtCnpj.Text) = False Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              GravaCliente = False
           End If
        End If

        If Len(txtCnpj.Text) = 14 Then
           If FU_ValidaCGC(txtCnpj.Text) = False Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
           GravaCliente = False
           End If
        Else
           If Len(txtCnpj.Text) <> 11 Then
              lblClite(4).ForeColor = vbRed
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
           GravaCliente = False
           End If
        End If
      End If

      If Trim(txtCnpj.Text) = "" Then
         txtCnpj.Text = right(String(14, "0") & txtCnpj.Text, 15)
      End If
      If txtCnpj.Text = 0 Then
         txtCnpj.Text = right(String(14, "0") & txtCnpj.Text, 15)
      End If
    
    If cmbPessoa.Text = "FÍSICA" Then
        Pessoa = "F"
    ElseIf cmbPessoa.Text = "JURÍDICA" Then
        Pessoa = "J"
    ElseIf Trim(UCase(cmbPessoa.Text)) = "FUNCIONÁRIO" Then
        Pessoa = "U"
    ElseIf Trim(UCase(cmbPessoa.Text)) = "ÓRGÃO PÚBLICO" Then
        Pessoa = "O"
    End If

    If txtNumero.Text = "" Then
        txtNumero.Text = 0
    End If
    If txtNumCobranca.Text = "" Then
        txtNumCobranca.Text = 0
    End If
    
    If chkCart.Value = True Then
        pagamentoCarteira = "S"
        Else
        pagamentoCarteira = "N"
    End If

    If Trim(UCase(cmbPessoa.Text)) <> "FÍSICA" Then
       If IsNumeric(txtInscricaoEstadual.Text) = False Then
          If Trim(UCase(txtInscricaoEstadual.Text)) <> "ISENTO" Then
             lblClite(6).ForeColor = vbRed
             txtInscricaoEstadual.SetFocus
             Screen.MousePointer = 0
             GravaCliente = False
          End If
       End If
    End If

    If Trim(txtNumero.Text) = "" Or Trim(txtNumero.Text) = 0 Then
       lblNumero.ForeColor = vbRed
       txtNumero.SetFocus
       txtNumero.SelStart = 0
       txtNumero.SelLength = Len(txtNumero.Text)
       Screen.MousePointer = 0
       GravaCliente = False
    End If

    If Trim(txtCodMun.Text) = "" Then
       lblClite(3).ForeColor = vbRed
       txtNumero.SetFocus
       txtCodMun.SelStart = 0
       txtCodMun.SelLength = Len(mskCep.Text)
       Screen.MousePointer = 0
       GravaCliente = False
    End If

    If IsNumeric(mskTelefone.Text) = False Or Trim(mskTelefone.Text) = "" Then
       lblClite(13).ForeColor = vbRed
       mskTelefone.SetFocus
       Screen.MousePointer = 0
       GravaCliente = False
    End If

    Numeros (mskFax.Text)
    mskFax.Text = wNumeros
    
    dataNascimento = Format(mskDataNascimento.Text, "yyyy/mm/dd")
    
    If GravaCliente = True Then
    
        On Error Resume Next
        adoCNLoja.BeginTrans
        'SQL = ""
        SQL = "SP_FIN_Grava_Cliente_Loja '" & txtCodigo.Text & "','" & txtRazaoSocial.Text & "','" & txtCnpj.Text & "','" _
                                            & Pessoa & "'," & Mid(cmbSituacao.Text, 1, 2) & ",'" & pagamentoCarteira _
                                            & "','" _
                                            & txtInscricaoEstadual.Text & "','" & mskCep.Text & "','" _
                                            & txtEndereco.Text & "','" & txtNumero.Text & "','" _
                                            & txtMunicipio.Text & "','" & txtCodMun.Text & "','" _
                                            & cmbUF.Text & "','" & txtComplemento.Text & "','" _
                                            & txtBairro.Text & "'," & Mid(cmbPraca.Text, 1, 1) & ",'" _
                                            & mskTelefone.Text & "','" & mskCelular.Text & "','" _
                                            & mskFax.Text & "','" _
                                            & dataNascimento & "','" _
                                            & Mid(cmbRamoAtiv.Text, 1, 2) & "','" & txtEMail.Text & "','" _
                                            & Mid(cmbSegmento.Text, 1, 2) & "', '" & txtEnderecoCobranca.Text & "','" _
                                            & txtNumCobranca.Text & "','" & txtComplCobranca.Text & "','" _
                                            & mskCepCobranca.Text & "','" & txtBairroCobranca.Text & "','" _
                                            & txtMunicipioCobranca.Text & "','" & txtEstadoCobranca.Text & "',0.00,'" _
                                            & Trim(cmbTipoCliente.Text) & "', '" _
                                            & Trim(txtClienteFidelidade) & "', '" _
                                            & left(frmPedido.txtVendedor.Text, 3) & "', '" _
                                            & Trim(GLB_Loja) & "'"
    
        adoCNLoja.Execute (SQL)
        adoCNLoja.CommitTrans
        
        If Err.Number = 0 Then
         
                SQL = "SP_FIN_Atualizando_Tela_De_Cadastro"
            
                adoCliente.CursorLocation = adUseClient
                adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
                adoCliente.Close ' SDSDSD
        '             frmManutencaoCliente.txtCampo.Text = txtCodigo.Text
                Screen.MousePointer = 0
                MsgBox "Cliente cadastrado com sucesso!", vbInformation, "Sucesso"
                frmConsCliente.txtPesquisaCliente.Text = frmCliente.txtCodigo.Text
                
                Unload Me
                
        Else
                adoCNLoja.RollbackTrans
                Screen.MousePointer = 0
                MsgBox "Erro ao cadastrar o cliente!", vbCritical, "ERRO"
                GravaCliente = False
        End If
        adoCliente.Close
    Else
        MsgBox "Cadastro ainda possui erros!", vbExclamation, "Atenção"
    End If
End Function


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
        Else
          '  retorno = ""
        End If
    Next Char
    
    Texto = retorno
    
    Numeros = Texto
    wNumeros = Texto

End Function

Function preencheStatus() As Boolean
   
    SQL = "SP_FIN_Ler_Situacao"
    rsSituacaoCliente.CursorLocation = adUseClient
    rsSituacaoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not rsSituacaoCliente.EOF Then
        preencheStatus = True
    Else
        preencheStatus = False
    End If
    
        If preencheStatus = True Then
        Do While Not rsSituacaoCliente.EOF
            cmbSituacao.AddItem UCase(rsSituacaoCliente("SI_CodigoSituacao")) & "-" & _
                                        UCase(rsSituacaoCliente("SI_Descricao"))
            rsSituacaoCliente.MoveNext
        Loop
        For I = 0 To cmbSituacao.ListCount
            cmbSituacao.ListIndex = I
            If Mid(cmbSituacao.Text, 1, 1) = 0 Then
                cmbSituacao.ListIndex = I
                Exit For
            End If
        Next I
    End If
    rsSituacaoCliente.Close
End Function

Private Sub txtRazaoSocial_LostFocus()
    txtRazaoSocial.Text = UCase(txtRazaoSocial.Text)
    If txtRazaoSocial.Text = Empty Then lblNome.ForeColor = vbRed
End Sub



