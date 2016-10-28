VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCliente 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cliente"
   ClientHeight    =   5385
   ClientLeft      =   90
   ClientTop       =   1935
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTroca 
      Left            =   14700
      Top             =   210
   End
   Begin VB.TextBox txtRazaoSocial 
      BackColor       =   &H00A3A3A3&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2370
      MaxLength       =   40
      TabIndex        =   2
      Top             =   705
      Width           =   6450
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   60
      ScaleHeight     =   45
      ScaleWidth      =   15000
      TabIndex        =   54
      Top             =   4665
      Width           =   15000
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Dados do Cliente"
      ForeColor       =   &H00E0E0E0&
      Height          =   4440
      Left            =   90
      TabIndex        =   32
      Top             =   105
      Width           =   15030
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         Caption         =   "Dados de Cobrança"
         ForeColor       =   &H00E0E0E0&
         Height          =   1140
         Left            =   105
         TabIndex        =   56
         Top             =   2310
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
            TabIndex        =   24
            Top             =   585
            Width           =   1710
         End
         Begin VB.TextBox txtBairroCobranca 
            BackColor       =   &H00A3A3A3&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11865
            MaxLength       =   15
            TabIndex        =   25
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
            Top             =   600
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
            Caption         =   "Município"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   17
            Left            =   6420
            TabIndex        =   63
            Top             =   315
            Width           =   705
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Endereço"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   15
            Left            =   1155
            TabIndex        =   62
            Top             =   315
            Width           =   690
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Bairro"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   16
            Left            =   11865
            TabIndex        =   61
            Top             =   315
            Width           =   405
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Estado"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   19
            Left            =   9435
            TabIndex        =   60
            Top             =   315
            Width           =   495
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "CEP"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   20
            Left            =   105
            TabIndex        =   59
            Top             =   315
            Width           =   315
         End
         Begin VB.Label lblNumeroCobranca 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "N.º"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   5595
            TabIndex        =   58
            Top             =   315
            Width           =   225
         End
         Begin VB.Label lblClite 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Complemento"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   12
            Left            =   10080
            TabIndex        =   57
            Top             =   315
            Width           =   960
         End
      End
      Begin VB.ComboBox cmbSituacao 
         BackColor       =   &H00A3A3A3&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12120
         TabIndex        =   5
         Top             =   585
         Width           =   2730
      End
      Begin VB.TextBox mskFax 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11520
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1890
         Width           =   1485
      End
      Begin VB.TextBox mskCelular 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9915
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1890
         Width           =   1530
      End
      Begin VB.TextBox mskTelefone 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8310
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   840
         MaxLength       =   14
         TabIndex        =   1
         Top             =   585
         Width           =   1395
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5820
         TabIndex        =   28
         Top             =   3825
         Width           =   6270
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
         Left            =   10575
         MaxLength       =   15
         TabIndex        =   4
         Top             =   585
         Width           =   1470
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         Top             =   585
         Width           =   660
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdMunicipio 
         Height          =   135
         Left            =   105
         TabIndex        =   31
         Top             =   2235
         Visible         =   0   'False
         Width           =   4530
         _cx             =   7990
         _cy             =   238
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
         Width           =   1305
      End
      Begin VB.ComboBox cmbPessoa 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8835
         TabIndex        =   3
         Top             =   585
         Width           =   1665
      End
      Begin VB.ComboBox cmbSegmento 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3825
         Width           =   2565
      End
      Begin VB.ComboBox cmbRamoAtiv 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3825
         Width           =   2985
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
         Left            =   6105
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1890
         Width           =   750
      End
      Begin VB.ComboBox cmbPraca 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6930
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1890
         Width           =   1305
      End
      Begin VB.CheckBox chkCart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Pagamento Carteira"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   13095
         TabIndex        =   18
         Tag             =   "Cart"
         Top             =   1890
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskDataCadastro 
         Height          =   315
         Left            =   13545
         TabIndex        =   30
         Tag             =   "DataCadastro"
         Top             =   3825
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataNascimento 
         Height          =   315
         Left            =   12165
         TabIndex        =   29
         Tag             =   "DataCadastro"
         Top             =   3825
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   12120
         TabIndex        =   55
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Código Município"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4725
         TabIndex        =   53
         Top             =   1635
         Width           =   1245
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Segmento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   36
         Left            =   3210
         TabIndex        =   52
         Top             =   3570
         Width           =   720
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Ramo de Atividade"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   35
         Left            =   105
         TabIndex        =   51
         Top             =   3570
         Width           =   1350
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Data Nascimento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   30
         Left            =   12165
         TabIndex        =   50
         Top             =   3570
         Width           =   1230
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Celular"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   18
         Left            =   9915
         TabIndex        =   49
         Tag             =   "Cel"
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Complemento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   7035
         TabIndex        =   48
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "N.º"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6270
         TabIndex        =   47
         Top             =   1005
         Width           =   225
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "E-mail"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   5820
         TabIndex        =   46
         Top             =   3570
         Width           =   420
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Bairro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   8
         Left            =   10530
         TabIndex        =   45
         Top             =   1020
         Width           =   405
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Endereço"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   7
         Left            =   1485
         TabIndex        =   44
         Top             =   1005
         Width           =   690
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Fax"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   14
         Left            =   11520
         TabIndex        =   43
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblCep 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "CEP"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   105
         TabIndex        =   42
         ToolTipText     =   "Click para consultar o cep da rua "
         Top             =   1005
         Width           =   315
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "UF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   11
         Left            =   6105
         TabIndex        =   41
         Top             =   1635
         Width           =   210
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Praça"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   10
         Left            =   6930
         TabIndex        =   40
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Inscr. Est."
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   10575
         TabIndex        =   39
         Top             =   300
         Width           =   705
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Telefone"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   13
         Left            =   8310
         TabIndex        =   38
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Municipio"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   37
         Top             =   1635
         Width           =   675
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Pessoa"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   8835
         TabIndex        =   36
         Top             =   315
         Width           =   525
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Data de Cadastro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   13560
         TabIndex        =   35
         Top             =   3570
         Width           =   1245
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "CNPJ/CPF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   34
         Top             =   315
         Width           =   780
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Código"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   33
         Top             =   315
         Width           =   645
      End
   End
   Begin Project1.chameleonButton cmdRetornar 
      Height          =   405
      Left            =   13920
      TabIndex        =   64
      Top             =   4875
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
      MICON           =   "frmCliente.frx":0052
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdGravar 
      Height          =   405
      Left            =   12765
      TabIndex        =   65
      Top             =   4890
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
      MICON           =   "frmCliente.frx":006E
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
      Left            =   11610
      TabIndex        =   66
      Top             =   4890
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
      MICON           =   "frmCliente.frx":008A
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
      Left            =   10035
      TabIndex        =   67
      Top             =   4890
      Width           =   1515
      _ExtentX        =   2672
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
      MICON           =   "frmCliente.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
Dim adoCliente As New ADODB.Recordset
Dim adoClienteCNPJ As New ADODB.Recordset
Dim adoMunicipio As New ADODB.Recordset
Dim rsSituacaoCliente As New ADODB.Recordset
Dim wPreencherCliente As Boolean
'Dim wNumeroClientePedido As Integer
Dim novoCodigo As Integer
Dim i As Integer
Dim pagamentoCarteira As String
Dim dataNascimento As String
Dim wLimpar As Boolean
Dim SQL As String * 1000


Function FU_ValidaCPF(CPF As String) As Integer

    Dim Soma As Integer
    Dim Resto As Integer
    Dim i As Integer

    If Len(CPF) <> 11 Then
        FU_ValidaCPF = False
        Exit Function
    End If

    Soma = 0
    For i = 1 To 9
        Soma = Soma + Val(Mid$(CPF, i, 1)) * (11 - i)
    Next i
    Resto = 11 - (Soma - (Int(Soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
        
    Soma = 0
    For i = 1 To 10
        Soma = Soma + Val(Mid$(CPF, i, 1)) * (12 - i)
    Next i
    Resto = 11 - (Soma - (Int(Soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
    
    FU_ValidaCPF = True

End Function

Function FU_ValidaCGC(CGC As String) As Integer
        Dim retorno, a, j, i, d1, d2
        If Len(CGC) = 8 And Val(CGC) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For i = 1 To 7
               a = Val(Mid(CGC, i, 1))
               If (i Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next i
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(CGC, 8, 1)) Then
              FU_ValidaCGC = True
           Else
              FU_ValidaCGC = False
           End If
        Else
           If Len(CGC) = 14 And Val(CGC) > 0 Then
              a = 0
              i = 0
              d1 = 0
              d2 = 0
              j = 5
              For i = 1 To 12 Step 1
                  a = a + (Val(Mid(CGC, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              i = 0
              j = 6
              For i = 1 To 13 Step 1
                  a = a + (Val(Mid(CGC, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
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
            ProximoCampo cmbUf
        End If
    ElseIf KeyAscii = 27 Then
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
            ProximoCampo txtEmail
            SelecionaCampo txtEmail
        End If
    ElseIf KeyAscii = 27 Then
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

Private Sub cmbUf_Change()
    txtEstadoCobranca.Text = cmbUf.Text
    
End Sub

Private Sub cmbUf_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If cmbUf.Text <> "" Then
            cmbUf.Text = UCase(cmbUf.Text)
            ProximoCampo mskTelefone
            SelecionaCampo mskTelefone
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo cmbPraca
    End If
End Sub

Private Sub CmdFichaFinanc_Click()
    wConexao = "Balcao"
    wCodigoCliFinan = txtCodigo.Text
       
        Cliente = txtCodigo.Text
     
        SQL = " Exec SP_FIN_Pesquisa_Cliente_Ficha_Financeira_Por_Codigo '" & Cliente & "'"
        adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

   If adoCliente.EOF = True Then
         MsgBox "Este cliente não possue ficha finaceira", vbCritical, "Atenção"
         adoCliente.Close
         Exit Sub
   End If

    adoCliente.Close

End Sub

Private Sub CmdGravar_Click()

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
    cmbUf.ListIndex = -1
    txtComplemento.Text = ""
    txtBairro.Text = ""
    cmbPraca.ListIndex = -1
    mskTelefone.Text = ""
    mskCelular.Text = ""
    mskFax.Text = ""
    mskDataNascimento.Text = ""
    cmbRamoAtiv.ListIndex = -1
    txtEmail.Text = ""
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


Private Sub CmdRetornar_Click()
    Call Limpar
    Unload Me
    
End Sub

Private Sub Form_Activate()
 tmrTroca.Interval = 1
End Sub

Private Sub Form_Load()

mskDataCadastro = Date

  frmCliente.Top = 4750
  frmCliente.Left = 90
  frmCliente.Width = 15190
  frmCliente.Height = 5800

  cmbPessoa.AddItem "FÍSICA"
  cmbPessoa.AddItem "JURÍDICA"
  cmbPessoa.AddItem "FUNCIONÁRIO"
  cmbPessoa.AddItem "ÓRGÃO PÚBLICO"
  grdMunicipio.Height = 990

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
            cmbUf.AddItem UCase(adoCliente("UF_Estado"))
            adoCliente.MoveNext
        Loop
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = "SP" Then
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
    End If
      
    If preencheUF = True Then
        Do While Not adoCliente.EOF
            cmbUf.AddItem UCase(adoCliente("UF_Estado"))
            adoCliente.MoveNext
        Loop
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = "SP" Then
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
        
    End If

    
       adoCliente.Close

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
            cmbUf.Enabled = True
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

Private Sub grdMunicipio_EnterCell()
    txtEstadoCobranca.Text = cmbUf.Text
    txtCodMun.Locked = True
    cmbUf.Locked = True
    cmbPraca.Locked = True
    mskTelefone.SetFocus
    
End Sub

Private Sub grdMunicipio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Or KeyAscii = 27 Then
        grdMunicipio.Visible = False
    End If
    
End Sub

Private Sub grdMunicipio_LostFocus()

'If grdMunicipio.Row < 0 Then
' Exit Sub
'Else
    txtMunicipio.Text = UCase(grdMunicipio.TextMatrix(grdMunicipio.Row, 0))
    txtMunicipioCobranca.Text = txtMunicipio.Text
    grdMunicipio.Visible = False


    If txtMunicipio.Text = "SÃO PAULO" Then
        cmbPraca.Text = "1 - SP"
        ElseIf txtMunicipio.Text <> "SÃO PAULO" Then

        cmbPraca.Text = "2 - Outro"
    End If
 ' End If
End Sub
Private Sub grdMunicipio_RowColChange()
   On Error GoTo SaidaRotina

    cmbUf.Text = UCase(grdMunicipio.TextMatrix(grdMunicipio.Row, 2))
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
            mskCep.SetFocus
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
        ProximoCampo txtEmail
        SelecionaCampo txtEmail
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
        ProximoCampo cmbUf
    End If
    
End Sub

Private Sub tmrTroca_Timer()
    Call TrocaBannerTopo1
End Sub

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
        ProximoCampo mskCepCobranca
        SelecionaCampo mskCepCobranca
    End If
End Sub

Private Sub txtBairroCobranca_LostFocus()
     txtBairroCobranca.Text = UCase(txtBairroCobranca.Text)
End Sub

Private Sub txtCNPJ_Change()
       Numeros (txtCnpj.Text)
    txtCnpj.Text = wNumeros
End Sub

Private Sub txtCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCnpj.Text <> "" And IsNumeric(txtCnpj) = True Then
            ProximoCampo txtRazaoSocial
        Else
            txtCnpj.SetFocus
            txtCnpj.SelStart = 0
            txtCnpj.SelLength = Len(txtCnpj.Text)
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo txtRazaoSocial
        SelecionaCampo txtRazaoSocial
    End If
End Sub

Private Sub txtCNPJ_LostFocus()
        If cmbPessoa.Text = "FÍSICA" Or cmbPessoa.Text = "FUNCIONÁRIO" Then
            If Len(txtCnpj.Text) < 11 And txtCnpj.Text <> "" Then
                MsgBox "Digite todos os dígitos do CPF!", vbCritical, "ATENÇÃO"
                txtCnpj.SetFocus
                txtCnpj.SelStart = 0
                txtCnpj.SelLength = Len(txtCnpj.Text)
                Screen.MousePointer = 0
                Exit Sub
            End If
        ElseIf cmbPessoa.Text = "JURÍDICA" Or cmbPessoa.Text = "ÓRGÃO PÚBLICO" Then
            If Len(txtCnpj.Text) < 14 And txtCnpj.Text <> "" Then
                MsgBox "Digite todos os dígitos do CNPJ!", vbCritical, "ATENÇÃO"
                txtCnpj.SetFocus
                txtCnpj.SelStart = 0
                txtCnpj.SelLength = Len(txtCnpj.Text)
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        
        If txtCnpj.Text <> "" Then
            If Len(txtCnpj.Text) = 14 Then
                If FU_ValidaCGC(txtCnpj.Text) = False Then
                    MsgBox "CNPJ INVÁLIDO", vbCritical, "Atenção"
                    txtCnpj.SetFocus
                    txtCnpj.SelStart = 0
                    txtCnpj.SelLength = Len(txtCnpj.Text)
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            If Len(txtCnpj.Text) = 11 Then
                If FU_ValidaCPF(txtCnpj.Text) = False Then
                    MsgBox "CPF INVÁLIDO", vbCritical, "Atenção"
                    txtCnpj.SetFocus
                    txtCnpj.SelStart = 0
                    txtCnpj.SelLength = Len(txtCnpj.Text)
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        
        If txtCnpj.Text <> "" Then
            If Not IsNumeric(txtCnpj.Text) Then
                MsgBox "Digite apenas números!", vbCritical, "ATENÇÃO"
                txtCnpj.Text = ""
                txtCnpj.SetFocus
                Exit Sub
            End If
            If VerificaClienteExisteCNPJ(txtCnpj.Text) = True Then
    '            PreencheDadosClienteCNPJ txtCnpj.Text
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
    If txtEmail.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtEmail.Text = ""
        txtEmail.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskDataCadastro.Text = Format(Date, "yyyy/mm/dd")
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    ElseIf KeyAscii = 27 Then
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
            txtEstadoCobranca.Text = UCase(cmbUf.Text)
        End If
    ElseIf KeyAscii = 27 Then
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
        ProximoCampo chkCart
    End If
End Sub

Private Sub txtInscricaoEstadual_LostFocus()
    If Mid(UCase(cmbPessoa.Text), 1, 1) = "F" Then
       txtInscricaoEstadual.Text = "ISENTO"
    End If
    

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
    grdMunicipio.ZOrder
    grdMunicipio.Visible = True
End Sub

Private Sub txtMunicipio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        grdMunicipio.SetFocus
    End If
End Sub

Private Sub txtMunicipio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Exit Sub
    ElseIf KeyAscii = 27 Then
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
        ProximoCampo txtEndereco
        SelecionaCampo txtEndereco
    End If
End Sub

Private Sub txtNumero_LostFocus()
    txtNumCobranca.Text = txtNumero.Text
    If Not IsNumeric(txtNumero.Text) Then
        MsgBox "Digite somente números!", vbCritical, "ATENÇÃO"
        txtNumero.Text = ""
        txtNumero.SetFocus
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
        txtEmail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
       
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
        For i = 0 To cmbPraca.ListCount - 1
            cmbPraca.ListIndex = i
            If Val(Mid(cmbPraca.Text, 1, 1)) = adoCliente("CE_Praca") Then
                cmbPraca.ListIndex = i
                Exit For
            End If
        Next i
        mskCep.Text = adoCliente("CE_Cep")
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = UCase(adoCliente("CE_Estado")) Then
                auxCMBUF = UCase(adoCliente("CE_Estado"))
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
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
 
        txtEmail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
        mskDataNascimento.Enabled = True
        mskDataNascimento.Text = IIf(IsNull(adoCliente("CE_DataNasc")), "01/01/1900", adoCliente("CE_DataNasc"))
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
        txtEmail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
       
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
        For i = 0 To cmbPraca.ListCount - 1
            cmbPraca.ListIndex = i
            If Val(Mid(cmbPraca.Text, 1, 1)) = adoCliente("CE_Praca") Then
                cmbPraca.ListIndex = i
                Exit For
            End If
        Next i
        mskCep.Text = adoCliente("CE_Cep")
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = UCase(adoCliente("CE_Estado")) Then
                auxCMBUF = UCase(adoCliente("CE_Estado"))
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
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
 
        txtEmail.Text = IIf(IsNull(adoCliente("CE_Email")), "", adoCliente("CE_Email"))
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
    Dim index As Integer
    index = 0

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
                   cmbRamoAtiv.ListIndex = index
                End If
              End If
              If cmbRamoAtiv.Text = "" Then
                  If RTrim(LTrim(adoRamo("RMO_DescricaoRamo"))) = "Indefinido" Or RTrim(LTrim(adoRamo("RMO_DescricaoRamo"))) = "Orgao Publico" Then
                     cmbRamoAtiv.ListIndex = index
                  End If
              End If
              index = index + 1
              adoRamo.MoveNext
        Loop

    Call carregaSegmento(Mid(cmbRamoAtiv.Text, 1, 2), txtCodigo.Text)
    End If
    adoRamo.Close
    adoRamo2.Close
    
End Function

Function carregaSegmento(ByVal RamoAtividade As String, ByVal CodigoCliente As String)
    Dim adoSegmento As New ADODB.Recordset
    Dim adoSegmento2 As New ADODB.Recordset
    Dim index As Integer
    index = 0
    
    cmbSegmento.Clear
    
    SQL = "SP_FIN_Pesquisa_Segmento_Por_Ramo_Atividade '" & RamoAtividade & "'"
    
    adoSegmento.CursorLocation = adUseClient
    adoSegmento.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
  
    SQL = "SP_FIN_Pesquisa_Segmento_Por_Codigo_Cliente '" & CodigoCliente & "'"
    
    adoSegmento2.CursorLocation = adUseClient
    adoSegmento2.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not adoSegmento.EOF Then
        Do While Not adoSegmento.EOF
            cmbSegmento.AddItem adoSegmento("SEG_CodigoSegmento") & " - " & adoSegmento("SEG_Descricao")
            If Not adoSegmento2.EOF Then
               If adoSegmento("SEG_CodigoSegmento") = adoSegmento2("ce_Segmento") And _
                  adoSegmento("SEG_RamoAtividade") = adoSegmento2("ce_RamoAtividade") Then
                    cmbSegmento.ListIndex = index
               End If
            End If
            If cmbSegmento.Text = "" Then
               If RTrim(LTrim(adoSegmento("SEG_descricao"))) = "Indefinido" Then
                 cmbSegmento.ListIndex = index
               End If
            End If
            index = index + 1
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
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = UCase(adoCliente("UF")) Then
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
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
        cmbUf.Text = adoCodigo("Mun_UF")
        
        mskTelefone.SetFocus
        adoCodigo.Close
        
     End If
    
    'adoCodigo.Close
    
    
    adoCliente.Close
End Function

Function verificaCamposNulos() As Boolean
      verificaCamposNulos = True
    If txtRazaoSocial.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtRazaoSocial
        SelecionaCampo txtRazaoSocial
    ElseIf cmbPessoa.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo cmbPessoa
        SelecionaCampo cmbPessoa
    ElseIf txtEndereco.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtEndereco
        SelecionaCampo txtEndereco
    ElseIf txtNumero.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtNumero
        SelecionaCampo txtNumero
    ElseIf txtBairro.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtBairro
        SelecionaCampo txtBairro
    ElseIf txtCnpj.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtCnpj
        SelecionaCampo txtCnpj
    ElseIf txtInscricaoEstadual.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtInscricaoEstadual
        SelecionaCampo txtInscricaoEstadual
    ElseIf txtMunicipio.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo txtMunicipio
        SelecionaCampo txtMunicipio
    ElseIf cmbPraca.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo cmbPraca
    ElseIf cmbUf.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo cmbUf
    ElseIf mskCep.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo mskCep
        SelecionaCampo mskCep
    ElseIf mskTelefone.Text = "" Then
        verificaCamposNulos = False
        ProximoCampo mskTelefone
        SelecionaCampo mskTelefone
    ElseIf IsDate(mskDataCadastro.Text) = False Then
        verificaCamposNulos = False
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    ElseIf txtEstadoCobranca.Text = "" Then
        txtEstadoCobranca.Text = cmbUf.Text
        verificaCamposNulos = True
    ElseIf mskCepCobranca.Text = "" Then
        mskCepCobranca.Text = mskCep.Text
        verificaCamposNulos = True
    ElseIf IsDate(mskDataNascimento.Text) = False Then
        verificaCamposNulos = True
        ProximoCampo mskDataNascimento
    End If
    If verificaCamposNulos = False Then
        MsgBox "Preencha todos os campos obrigatórios", vbCritical, "Atenção"
        Exit Function
    End If
End Function

Function AtualizaCliente(ByVal codigo As Double) As Boolean
    Dim Pessoa As String
        
    AtualizaCliente = False

    If Trim(mskCep.Text) = "0" Or Trim(mskCep.Text) = "" Then
        MsgBox "CEP Inválido", vbCritical, "Atenção"
        mskCep.SelStart = 0
        mskCep.SelLength = Len(mskCep.Text)
        Screen.MousePointer = 0
        Exit Function
    End If

    If Trim(txtCnpj.Text) <> "" Then

        If Len(txtCnpj.Text) = 11 And UCase(Mid(cmbPessoa.Text, 1, 1)) <> "F" Then
              MsgBox "CNPJ INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              Exit Function
        End If
        
        If Len(txtCnpj.Text) = 14 And UCase(Mid(cmbPessoa.Text, 1, 1)) = "F" Then
              MsgBox "CPF INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.Locked = True
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              Exit Function
        End If
    
        If Len(txtCnpj.Text) = 11 Then
           If FU_ValidaCPF(txtCnpj.Text) = False Then
              MsgBox "CPF INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.Locked = True
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              Exit Function
           End If
        End If

        If Len(txtCnpj.Text) = 14 Then
           If FU_ValidaCGC(txtCnpj.Text) = False Then
              MsgBox "CNPJ INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.Locked = True
              txtCnpj.SetFocus

              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
           Exit Function
           End If
        Else
           'If Len(txtCnpj.Text) <> 11 Then
           '   MsgBox "CNPJ/CPF INVALIDO", vbCritical, "Atenção"
           '   txtCnpj.SetFocus
           '   txtCnpj.SelStart = 0
           '   txtCnpj.SelLength = Len(txtCnpj.Text)
           '   Screen.MousePointer = 0
           'Exit Function
           'End If
        End If
      End If

      If Trim(txtCnpj.Text) = "" Then
         txtCnpj.Text = Right(String(14, "0") & txtCnpj.Text, 15)
      End If
      If txtCnpj.Text = 0 Then
         txtCnpj.Text = Right(String(14, "0") & txtCnpj.Text, 15)
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
             Exit Function
          End If
       End If
    End If

          If Trim(txtNumero.Text) = "" Or Trim(txtNumero.Text) = 0 Then
             MsgBox "Número do endereço inválido!", vbCritical, "Atenção"
             txtNumero.SetFocus
             txtNumero.SelStart = 0
             txtNumero.SelLength = Len(txtNumero.Text)
             Screen.MousePointer = 0
             Exit Function
          End If

    If Trim(txtCodMun.Text) = "" Then
       MsgBox "Muncípio ou Código do Município inválido", vbCritical, "Atenção"
       txtNumero.SetFocus
       txtCodMun.SelStart = 0
       txtCodMun.SelLength = Len(mskCep.Text)
       Screen.MousePointer = 0
       Exit Function
    End If
    
    If Pessoa = "O" Or cmbUf.Text <> "SP" Then
       If Trim(txtEmail.Text) = "" Then
          MsgBox "Obrigatório informar um E-mail para envio de informações da NF Eletrônica", vbCritical, "Atenção"
          txtEmail.SetFocus
          txtEmail.SelStart = 0
          txtEmail.SelLength = Len(txtEmail.Text)
          Screen.MousePointer = 0
          Exit Function
       End If
    End If

    If IsNumeric(mskTelefone.Text) = False Then
       MsgBox "Telefone Inválido", vbCritical, "Atenção"
       mskTelefone.SetFocus
       Screen.MousePointer = 0
       Exit Function
    End If
    
    If Trim(mskTelefone.Text) = "" Then
       MsgBox "Telefone Inválido", vbCritical, "Atenção"
       mskTelefone.SetFocus
       Screen.MousePointer = 0
       Exit Function
    End If

    Numeros (mskFax.Text)
    mskFax.Text = wNumeros
    
    dataNascimento = Format(mskDataNascimento.Text, "yyyy/mm/dd")

     adoCNLoja.BeginTrans
        
        SQL = "SP_FIN_Altera_Cliente " & codigo & ",'" & txtRazaoSocial.Text & "','" & txtCnpj.Text & "','" _
                                        & Mid(Pessoa, 1, 2) & "', " & Val(cmbSituacao.Text) & ",'" _
                                        & pagamentoCarteira & "', '" _
                                        & txtInscricaoEstadual.Text & "','" & mskCep.Text & "','" _
                                        & txtEndereco.Text & "','" & txtNumero.Text & "','" _
                                        & txtMunicipio.Text & "'," & txtCodMun.Text & ",'" _
                                        & cmbUf.Text & "','" & txtComplemento.Text & "','" _
                                        & txtBairro.Text & "'," & Mid(cmbPraca.Text, 1, 1) & ",'" _
                                        & mskTelefone.Text & "','" & mskCelular.Text & "','" _
                                        & mskFax.Text & "','" & dataNascimento & "'," _
                                        & Val(cmbRamoAtiv.Text) & ",'" & txtEmail.Text & "'," _
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
'    adoCliente.Close
End Function

Function GravaCliente() As Boolean
    Dim Pessoa As String
        
    GravaCliente = False
    
    cmbPessoa.Enabled = True
    
        
     
    If Trim(mskCep.Text) = "0" Or Trim(mskCep.Text) = "" Then
        MsgBox "CEP Inválido", vbCritical, "Atenção"
        mskCep.SelStart = 0
        mskCep.SelLength = Len(mskCep.Text)
        Screen.MousePointer = 0
        Exit Function
    End If



    If Trim(txtCnpj.Text) <> "" Then
 
'        SQL = " exec SP_FIN_Ler_Clientes_Por_Parametro_Cnpj '" & txtCnpj.Text & "'"
'          adoClienteCNPJ.CursorLocation = adUseClient
'          adoClienteCNPJ.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'        If Not adoClienteCNPJ.EOF Then
'           MsgBox "Já existe cadastro com esse CNPJ/CPF"
'           txtCnpj.SetFocus
'           txtCnpj.SelStart = 0
'           txtCnpj.SelLength = Len(txtCnpj.Text)
'           Screen.MousePointer = 0
'           adoClienteCNPJ.Close
'           Exit Function
'        End If
'        adoClienteCNPJ.Close

        If Len(txtCnpj.Text) = 11 And UCase(Mid(cmbPessoa.Text, 1, 1)) <> "F" Then
              MsgBox "CNPJ INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              Exit Function
        End If
        
        If Len(txtCnpj.Text) = 14 And UCase(Mid(cmbPessoa.Text, 1, 1)) = "F" Then
              MsgBox "CPF INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              Exit Function
        End If
    
        If Len(txtCnpj.Text) = 11 Then
           If FU_ValidaCPF(txtCnpj.Text) = False Then
              MsgBox "CPF INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
              Exit Function
           End If
        End If

        If Len(txtCnpj.Text) = 14 Then
           If FU_ValidaCGC(txtCnpj.Text) = False Then
              MsgBox "CNPJ INVÁLIDO", vbCritical, "Atenção"
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
           Exit Function
           End If
        Else
           If Len(txtCnpj.Text) <> 11 Then
              MsgBox "CNPJ/CPF INVALIDO", vbCritical, "Atenção"
              txtCnpj.SetFocus
              txtCnpj.SelStart = 0
              txtCnpj.SelLength = Len(txtCnpj.Text)
              Screen.MousePointer = 0
           Exit Function
           End If
        End If
      End If

      If Trim(txtCnpj.Text) = "" Then
         txtCnpj.Text = Right(String(14, "0") & txtCnpj.Text, 15)
      End If
      If txtCnpj.Text = 0 Then
         txtCnpj.Text = Right(String(14, "0") & txtCnpj.Text, 15)
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
             MsgBox "Inscrição Estadual Invalida", vbCritical, "Atenção"
             txtInscricaoEstadual.SetFocus
             Screen.MousePointer = 0
             Exit Function
          End If
       End If
    End If

          If Trim(txtNumero.Text) = "" Or Trim(txtNumero.Text) = 0 Then
             MsgBox "Numero do endereço Inválido", vbCritical, "Atenção"
             txtNumero.SetFocus
             txtNumero.SelStart = 0
             txtNumero.SelLength = Len(txtNumero.Text)
             Screen.MousePointer = 0
             Exit Function
          End If

    If Trim(txtCodMun.Text) = "" Then
       MsgBox "Muncípio ou código do Município Inválido", vbCritical, "Atenção"
       txtNumero.SetFocus
       txtCodMun.SelStart = 0
       txtCodMun.SelLength = Len(mskCep.Text)
       Screen.MousePointer = 0
       Exit Function
    End If
    
    If Pessoa = "O" Or cmbUf.Text <> "SP" Then
       If Trim(txtEmail.Text) = "" Then
          MsgBox "Obrigatório informar um E-mail para envio de informações da NF Eletrônica", vbCritical, "Atenção"
          txtEmail.SetFocus
          txtEmail.SelStart = 0
          txtEmail.SelLength = Len(txtEmail.Text)
          Screen.MousePointer = 0
          Exit Function
       End If
    End If

    If IsNumeric(mskTelefone.Text) = False Then
       MsgBox "Telefone Inválido", vbCritical, "Atenção"
       mskTelefone.SetFocus
       Screen.MousePointer = 0
       Exit Function
    End If
    
    If Trim(mskTelefone.Text) = "" Then
       MsgBox "Telefone Inválido", vbCritical, "Atenção"
       mskTelefone.SetFocus
       Screen.MousePointer = 0
       Exit Function
    End If

    Numeros (mskFax.Text)
    mskFax.Text = wNumeros
    
    dataNascimento = Format(mskDataNascimento.Text, "yyyy/mm/dd")

    On Error Resume Next
    adoCNLoja.BeginTrans
    SQL = ""
    SQL = "SP_FIN_Grava_Cliente_Loja '" & txtCodigo.Text & "','" & txtRazaoSocial.Text & "','" & txtCnpj.Text & "','" _
                                        & Pessoa & "'," & Mid(cmbSituacao.Text, 1, 2) & ",'" & pagamentoCarteira _
                                        & "','" _
                                        & txtInscricaoEstadual.Text & "','" & mskCep.Text & "','" _
                                        & txtEndereco.Text & "','" & txtNumero.Text & "','" _
                                        & txtMunicipio.Text & "','" & txtCodMun.Text & "','" _
                                        & cmbUf.Text & "','" & txtComplemento.Text & "','" _
                                        & txtBairro.Text & "'," & Mid(cmbPraca.Text, 1, 1) & ",'" _
                                        & mskTelefone.Text & "','" & mskCelular.Text & "','" _
                                        & mskFax.Text & "','" _
                                        & dataNascimento & "'," _
                                        & Mid(cmbRamoAtiv.Text, 1, 2) & ",'" & txtEmail.Text & "'," _
                                        & Mid(cmbSegmento.Text, 1, 2) & ", '" & txtEnderecoCobranca.Text & "','" _
                                        & txtNumCobranca.Text & "','" & txtComplCobranca.Text & "','" _
                                        & mskCepCobranca.Text & "','" & txtBairroCobranca.Text & "','" _
                                        & txtMunicipioCobranca.Text & "','" & txtEstadoCobranca.Text & "',0.00"

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
        For i = 0 To cmbSituacao.ListCount
            cmbSituacao.ListIndex = i
            If Mid(cmbSituacao.Text, 1, 1) = 0 Then
                cmbSituacao.ListIndex = i
                Exit For
            End If
        Next i
    End If
    rsSituacaoCliente.Close
End Function

Private Sub txtRazaoSocial_LostFocus()
    txtRazaoSocial.Text = UCase(txtRazaoSocial.Text)
End Sub

