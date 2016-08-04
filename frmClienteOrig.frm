VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCliente 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Cliente"
   ClientHeight    =   5010
   ClientLeft      =   1980
   ClientTop       =   1890
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCliente.frx":0000
   ScaleHeight     =   5010
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   15
      TabIndex        =   102
      Top             =   3405
      Width           =   6360
   End
   Begin VB.TextBox txtCobrNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4500
      TabIndex        =   87
      Top             =   3690
      Width           =   615
   End
   Begin VB.ComboBox cmbPessoa 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3750
      TabIndex        =   48
      Text            =   "cmbPessoa"
      Top             =   420
      Width           =   1170
   End
   Begin VB.TextBox txtSituacao 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5625
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   47
      Text            =   "0"
      Top             =   420
      Width           =   630
   End
   Begin VB.CheckBox chkCart 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   195
      Left            =   1650
      TabIndex        =   46
      Top             =   825
      Width           =   210
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5235
      TabIndex        =   45
      Top             =   1080
      Width           =   1020
   End
   Begin VB.ComboBox cmbUf 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5355
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   1740
      Width           =   900
   End
   Begin VB.ComboBox cmbPraca 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3675
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   1740
      Width           =   1260
   End
   Begin VB.ComboBox cmbRamoAtiv 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3675
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   2400
      Width           =   2595
   End
   Begin VB.ComboBox cmbSegmento 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   3045
      Width           =   2160
   End
   Begin VB.Frame fraCliente 
      ForeColor       =   &H00FF0000&
      Height          =   3510
      Left            =   6990
      TabIndex        =   9
      Top             =   990
      Width           =   6210
   End
   Begin VB.Frame fraCobranca 
      ForeColor       =   &H00FF0000&
      Height          =   840
      Left            =   7065
      TabIndex        =   8
      Top             =   2565
      Width           =   6210
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4245
      Left            =   7845
      TabIndex        =   7
      Top             =   255
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   7488
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dados do Cliente"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dados Cobrança"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Limite de Crédito"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   50
      Left            =   15
      TabIndex        =   6
      Top             =   4395
      Width           =   6360
   End
   Begin TabDlg.SSTab tabCliente 
      Height          =   4215
      Left            =   6825
      TabIndex        =   5
      Top             =   105
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Dados do Cliente"
      TabPicture(0)   =   "frmCliente.frx":7D722
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Dados Cobrança"
      TabPicture(1)   =   "frmCliente.frx":7D73E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Limite de Crédito"
      TabPicture(2)   =   "frmCliente.frx":7D75A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraCredito"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraCredito 
         Height          =   3810
         Left            =   -45
         TabIndex        =   10
         Top             =   900
         Width           =   6180
         Begin MSMask.MaskEdBox mskMaiorCompra 
            Height          =   315
            Left            =   1530
            TabIndex        =   11
            Tag             =   "MaiorCmp"
            Top             =   1260
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            Format          =   "###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskUltPagto 
            Height          =   315
            Left            =   1530
            TabIndex        =   12
            Tag             =   "UltimoPgto"
            Top             =   1605
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            Format          =   " ###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDupliAberto 
            Height          =   315
            Left            =   1530
            TabIndex        =   13
            Tag             =   "DupAberto"
            Top             =   570
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            Format          =   "###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskUltCompra 
            Height          =   315
            Left            =   1530
            TabIndex        =   14
            Tag             =   "UltimaCmp"
            Top             =   900
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            Format          =   "###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLimiteCred 
            Height          =   315
            Left            =   1530
            TabIndex        =   15
            Tag             =   "LimiteCre"
            Top             =   225
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            Format          =   "###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskQtdeDupliAberto 
            Height          =   315
            Left            =   1530
            TabIndex        =   16
            Tag             =   "QtdDupAberto"
            Top             =   2295
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   15
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataUltCompra 
            Height          =   315
            Left            =   1530
            TabIndex        =   17
            Tag             =   "DataUltimaCmp"
            Top             =   2640
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataLimite 
            Height          =   315
            Left            =   1530
            TabIndex        =   18
            Tag             =   "DataLimiteCre"
            Top             =   1950
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataMaiorCompra 
            Height          =   315
            Left            =   4575
            TabIndex        =   19
            Tag             =   "DataMaiorCmp"
            Top             =   225
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataUltPagto 
            Height          =   315
            Left            =   4575
            TabIndex        =   20
            Tag             =   "DataUltimoPgto"
            Top             =   570
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskJurosCartorio 
            Height          =   315
            Left            =   4560
            TabIndex        =   21
            Tag             =   "JuroCartorio"
            Top             =   915
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            Format          =   "###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskQtdeCompra 
            Height          =   315
            Left            =   4560
            TabIndex        =   22
            Tag             =   "QtdCmp"
            Top             =   1260
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSaldoCompra 
            Height          =   315
            Left            =   4560
            TabIndex        =   23
            Top             =   1605
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   40
            Format          =   "###,###,###,###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskMaiorAtraso 
            Height          =   315
            Left            =   4560
            TabIndex        =   24
            Tag             =   "MaiorAtraso"
            Top             =   2295
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDupliAtraso 
            Height          =   315
            Left            =   4560
            TabIndex        =   25
            Tag             =   "QtdDupAtraso"
            Top             =   1950
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            MaxLength       =   40
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Juros de Cartório"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3195
            TabIndex        =   40
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. de Compras"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3195
            TabIndex        =   39
            Top             =   1335
            Width           =   1185
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo p/ Compra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3195
            TabIndex        =   38
            Top             =   1665
            Width           =   1200
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dupl. em Atraso"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3195
            TabIndex        =   37
            Top             =   2025
            Width           =   1125
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maior Atraso"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3195
            TabIndex        =   36
            Top             =   2370
            Width           =   885
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Maior Compra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3180
            TabIndex        =   35
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Último Pagto."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3180
            TabIndex        =   34
            Top             =   645
            Width           =   1335
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data do Limite"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   105
            TabIndex        =   33
            Top             =   2040
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. Dupl. Aberto"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   2370
            Width           =   1230
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Última Compra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   105
            TabIndex        =   31
            Top             =   2700
            Width           =   1410
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de Crédito"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   30
            Top             =   315
            Width           =   1170
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dupl. em Aberto"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   29
            Top             =   645
            Width           =   1140
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última Compra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   975
            Width           =   1020
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maior Compra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   27
            Top             =   1335
            Width           =   975
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último Pagamento"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   1680
            Width           =   1290
         End
      End
   End
   Begin VB.Frame fraBotoes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   3645
      TabIndex        =   1
      Top             =   4395
      Width           =   2625
      Begin VB.CommandButton cmdFichaFinanceira 
         Caption         =   "Ficha Financ."
         Height          =   390
         Left            =   15
         TabIndex        =   104
         Top             =   105
         Width           =   1095
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Grava"
         Enabled         =   0   'False
         Height          =   390
         Left            =   1110
         TabIndex        =   3
         Top             =   120
         Width           =   750
      End
      Begin VB.CommandButton cmdRetornar 
         Caption         =   "Retorna"
         Height          =   390
         Left            =   1860
         TabIndex        =   2
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   765
         Width           =   4410
      End
   End
   Begin VB.TextBox txtpedido 
      Height          =   300
      Left            =   900
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4635
      Visible         =   0   'False
      Width           =   1050
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2235
      OleObjectBlob   =   "frmCliente.frx":7D776
      Top             =   4470
   End
   Begin MSMask.MaskEdBox mskPessoa 
      Height          =   285
      Left            =   90
      TabIndex        =   49
      Tag             =   "TipoPessoa"
      Top             =   2895
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   16711680
      MaxLength       =   8
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskRazao 
      Height          =   315
      Left            =   1890
      TabIndex        =   50
      Tag             =   "Razao"
      Top             =   90
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCodCli 
      Height          =   315
      Left            =   990
      TabIndex        =   51
      Tag             =   "CodCli"
      Top             =   90
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483644
      ForeColor       =   16711680
      MaxLength       =   8
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCGCCPF 
      Height          =   315
      Left            =   990
      TabIndex        =   52
      Tag             =   "CGC"
      Top             =   420
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskInscEst 
      Height          =   315
      Left            =   2715
      TabIndex        =   53
      Tag             =   "IE"
      Top             =   750
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskEndereco 
      Height          =   315
      Left            =   975
      TabIndex        =   54
      Tag             =   "Endereco"
      Top             =   1080
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCEP 
      Height          =   315
      Left            =   5235
      TabIndex        =   55
      Tag             =   "CEP"
      Top             =   750
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   8
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskBairro 
      Height          =   315
      Left            =   3675
      TabIndex        =   56
      Tag             =   "Bairro"
      Top             =   1410
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483644
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCompl 
      Height          =   315
      Left            =   975
      TabIndex        =   57
      Tag             =   "Bairro"
      Top             =   1410
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskMunicipio 
      Height          =   315
      Left            =   975
      TabIndex        =   58
      Tag             =   "Municipio"
      Top             =   1740
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskTelefone 
      Height          =   315
      Left            =   975
      TabIndex        =   59
      Tag             =   "Telefone"
      Top             =   2070
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFAX 
      Height          =   315
      Left            =   4890
      TabIndex        =   60
      Tag             =   "FAX"
      Top             =   2070
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCelular 
      Height          =   315
      Left            =   2925
      TabIndex        =   61
      Tag             =   "FAX"
      Top             =   2070
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataNascimento 
      Height          =   315
      Left            =   975
      TabIndex        =   62
      Tag             =   "DataCadastro"
      Top             =   2400
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskEmail 
      Height          =   300
      Left            =   975
      TabIndex        =   63
      Tag             =   "Telefone"
      Top             =   2730
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   60
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataCadastro 
      Height          =   315
      Left            =   5070
      TabIndex        =   64
      Tag             =   "DataCadastro"
      Top             =   3045
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCobrEnd 
      Height          =   315
      Left            =   495
      TabIndex        =   88
      Tag             =   "EnderecoCobr"
      Top             =   3705
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCobrCEP 
      Height          =   330
      Left            =   495
      TabIndex        =   89
      Tag             =   "CEPCobr"
      Top             =   4035
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCobrCompl 
      Height          =   315
      Left            =   5565
      TabIndex        =   90
      Tag             =   "Bairro"
      Top             =   3675
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCobrMunicipio 
      Height          =   315
      Left            =   4125
      TabIndex        =   91
      Tag             =   "MunicipioCobr"
      Top             =   4020
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCobrBairro 
      Height          =   315
      Left            =   2085
      TabIndex        =   92
      Tag             =   "BairroCobr"
      Top             =   4035
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCobrEstado 
      Height          =   315
      Left            =   5895
      TabIndex        =   93
      Tag             =   "UFCobr"
      Top             =   4035
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      MaxLength       =   2
      PromptChar      =   "_"
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "Dados Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   9060
      TabIndex        =   103
      Top             =   2475
      Width           =   1770
   End
   Begin VB.Label lblCobranca 
      BackStyle       =   0  'Transparent
      Caption         =   "Dados Cobrança"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   135
      TabIndex        =   101
      Top             =   3465
      Width           =   1770
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5640
      TabIndex        =   100
      Top             =   4080
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1650
      TabIndex        =   99
      Top             =   4095
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Município"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3375
      TabIndex        =   98
      Top             =   4080
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   135
      TabIndex        =   97
      Top             =   4095
      Width           =   315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N.º"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4245
      TabIndex        =   96
      Top             =   3750
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compl."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5115
      TabIndex        =   95
      Top             =   3735
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   135
      TabIndex        =   94
      Top             =   3795
      Width           =   450
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   86
      Top             =   150
      Width           =   495
   End
   Begin VB.Label lblCGCCPF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGC/CPF"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   85
      Top             =   525
      Width           =   705
   End
   Begin VB.Label lblSituacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4965
      TabIndex        =   84
      Top             =   465
      Width           =   630
   End
   Begin VB.Label lblPessoa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pessoa"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3120
      TabIndex        =   83
      Top             =   465
      Width           =   525
   End
   Begin VB.Label lblInscrEst 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscr. Est."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2010
      TabIndex        =   82
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lblPagCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagamento Carteira"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   81
      Top             =   825
      Width           =   1395
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   80
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label lblNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N.º"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4935
      TabIndex        =   79
      Top             =   1170
      Width           =   225
   End
   Begin VB.Label lblCep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4875
      TabIndex        =   78
      Top             =   825
      Width           =   390
   End
   Begin VB.Label lblBairro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3150
      TabIndex        =   77
      Top             =   1470
      Width           =   450
   End
   Begin VB.Label lblCompl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compl."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   76
      Top             =   1470
      Width           =   480
   End
   Begin VB.Label lblCidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   75
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblUF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5040
      TabIndex        =   74
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label lblPraca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Praça"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3150
      TabIndex        =   73
      Top             =   1785
      Width           =   420
   End
   Begin VB.Label lblFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4530
      TabIndex        =   72
      Top             =   2145
      Width           =   255
   End
   Begin VB.Label lblCel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cel"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2565
      TabIndex        =   71
      Top             =   2085
      Width           =   225
   End
   Begin VB.Label lblTelefone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   70
      Top             =   2130
      Width           =   630
   End
   Begin VB.Label lblRamoAtiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ramo de Atividade"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2250
      TabIndex        =   69
      Top             =   2475
      Width           =   1350
   End
   Begin VB.Label lblDataNasc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Nasc."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   68
      Top             =   2430
      Width           =   810
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   67
      Top             =   2790
      Width           =   435
   End
   Begin VB.Label lblSegmento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Segmento"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   66
      Top             =   3105
      Width           =   720
   End
   Begin VB.Label lblDataCadastro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Cadastro"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3705
      TabIndex        =   65
      Top             =   3090
      Width           =   1290
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim i As Integer
Dim ExisteUF As Boolean

Private Sub chkCart_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 83 Then
        chkCart.Value = 1
    ElseIf KeyCode = 78 Then
        chkCart.Value = 0
    ElseIf KeyCode = 13 Then
        ProximoCampo mskInscEst
        SelecionaCampo mskInscEst
    ElseIf KeyCode = 27 Then
        ProximoCampo txtSituacao
    End If
    
End Sub

Private Sub cmdInfoCli_Click(Index As Integer)
    
    Unload Me
    
End Sub

Private Sub cmbPessoa_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
End Sub

Private Sub cmbPessoa_LostFocus()
mskPessoa.Text = cmbPessoa.Text
Call mskPessoa_KeyPress(13)
End Sub

Private Sub cmbPraca_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If cmbPraca.Text <> "" Then
            ProximoCampo cmbUf
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskMunicipio
        SelecionaCampo mskMunicipio
    End If

End Sub

Private Sub cmbRamoAtiv_Click()

    Call CarregaSegmento(Val(cmbRamoAtiv.Text))

End Sub

Private Sub cmbRamoAtiv_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If cmbRamoAtiv.Text <> "" Then
            ProximoCampo cmbSegmento
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskDataNascimento
        SelecionaCampo mskDataNascimento
    End If

End Sub

Private Sub cmbSegmento_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If cmbSegmento.Text <> "" Then
            ProximoCampo mskEmail
            SelecionaCampo mskEmail
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo cmbRamoAtiv
    End If

End Sub

Private Sub cmbUf_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If cmbUf.Text <> "" Then
            cmbUf.Text = UCase(cmbUf.Text)
            ProximoCampo mskTelefone
            SelecionaCampo mskTelefone
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo cmbPraca
        SelecionaCampo cmbPraca
    End If

End Sub

Private Sub cmdFichaFinanceira_Click()
    FrmFichaFinanceira.Show 1
End Sub

Private Sub cmdGravar_Click()
    mskCGCCPF.Text = Right(String(15, "0") & mskCGCCPF.Text, 15)
    
    If VerificaCamposNulos = True Then
        Screen.MousePointer = 11
        If wPreencherCliente = False Then
            If GravaCliente(mskCodCli.Text, 1) = True Then
                Unload Me
            End If
        Else
            If GravaCliente(mskCodCli.Text, 2) = True Then
                Unload Me
            End If
        End If
    End If
    
End Sub

Private Sub cmdRetornar_Click()
    Call LimpaForm
    Unload Me
    
End Sub

Private Sub Form_Load()
'  Left = (Screen.Width - Width) / 2
'  Top = (Screen.Height - Height) / 2
  
  Call AjustaTela(frmCliente)
  
  Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
  Skin1.ApplySkin Me.hwnd
  
  
  Call LimpaForm
  
  cmbPessoa.AddItem "Juridica"
  cmbPessoa.AddItem "Func"
  cmbPessoa.AddItem "Fisica"
  cmbPessoa.AddItem "Altarquia"
   For i = 0 To cmbPessoa.ListCount
       cmbPessoa.ListIndex = i
       If cmbPessoa.Text = "Juridica" Then
          cmbPessoa.ListIndex = i
          Exit For
       End If
   Next i
    
    SQL = ""
    PreencheUF
    If ExisteUF = True Then
        Do While Not rdoUf.EOF
            cmbUf.AddItem UCase(rdoUf("UF_Estado"))
            rdoUf.MoveNext
        Loop
        rdoUf.Close
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = "SP" Then
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    If wPreencherCliente = True Then
        DescricaoOperacao "Pesquisando Cliente"
        PreencheDadosCliente wNumeroClientePedido
        DescricaoOperacao "Pronto"
       ' fraCredito.Enabled = False
    Else
        DescricaoOperacao "Iniciando cadastro de cliente"
        SQL = ""
        SQL = "select (CTS_SequenciaCliente + 1) as UltNumCliente from ControleSistema "
              rsNumeroCliente.CursorLocation = adUseClient
              rsNumeroCliente.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set rsNumeroCliente = rdoCNLoja.OpenResultset(SQL)
        
        SQL = ""
        SQL = "Update ControleSistema Set CTS_SequenciaCliente = " & rsNumeroCliente("UltNumCliente") & ""
        rdoCNLoja.Execute (SQL)
        
        If Not rsNumeroCliente.EOF Then
   '         fraCredito.Enabled = True
            cmbPraca.AddItem "1 - SP"
            cmbPraca.AddItem "2 - Outro"
            cmbPraca.ListIndex = 0
            mskCodCli.Text = rsNumeroCliente("UltNumCliente")
            mskCodCli.Enabled = True
            mskRazao.Enabled = True
            mskPessoa.Enabled = True
            'txtSituacao.Enabled = True
            mskEndereco.Enabled = True
            mskBairro.Enabled = True
            mskCGCCPF.Enabled = True
            mskInscEst.Enabled = True
            mskMunicipio.Enabled = True
            'mskPraca.Enabled = True
            mskCEP.Enabled = True
            cmbUf.Enabled = True
            mskDataCadastro.Enabled = True
            mskTelefone.Enabled = True
            mskFAX.Enabled = True
            mskCobrEnd.Enabled = True
            mskCobrBairro.Enabled = True
            mskCobrMunicipio.Enabled = True
            mskCobrEstado.Enabled = True
            mskCobrCEP.Enabled = True
            mskLimiteCred.Enabled = True
            mskDataLimite.Enabled = True
            mskJurosCartorio.Enabled = True
            mskDupliAberto.Enabled = True
            mskQtdeDupliAberto.Enabled = True
            mskQtdeCompra.Enabled = True
            mskUltCompra.Enabled = True
            mskDataUltCompra.Enabled = True
            mskSaldoCompra.Enabled = True
            mskMaiorCompra.Enabled = True
            mskDataMaiorCompra.Enabled = True
            mskDupliAtraso.Enabled = True
            mskUltPagto.Enabled = True
            mskDataUltPagto.Enabled = True
            mskMaiorAtraso.Enabled = True
            mskDataNascimento.Enabled = True
            mskCelular.Enabled = True
            cmbRamoAtiv.Enabled = True
            cmbSegmento.Enabled = True
        End If
        rsNumeroCliente.Close
    End If
        
End Sub

Function PreencheDadosCliente(ByVal Cliente As Double)
    'Dim rsClientePedido As rdoResultset
    If PesquisaCliente(4, Cliente, rsClientePedido) = True Then
        If Trim(rsClientePedido("CE_PagamentoCarteira")) = "S" Then
            chkCart.Value = 1
        Else
            chkCart.Value = 0
        End If
        mskCodCli.Text = rsClientePedido("CE_CodigoCliente")
        mskRazao.Text = rsClientePedido("CE_Razao")
        If Trim(rsClientePedido("CE_TipoPessoa")) = "J" Then
            mskPessoa.Text = "Juridica"
        ElseIf Trim(rsClientePedido("CE_TipoPessoa")) = "F" Then
            mskPessoa.Text = "Fisica"
        ElseIf Trim(rsClientePedido("CE_TipoPessoa")) = "U" Then
            mskPessoa.Text = "Func"
        ElseIf Trim(rsClientePedido("CE_TipoPessoa")) = "A" Then
            mskPessoa.Text = "Altarquia"
        End If
        txtNumero = IIf(IsNull(rsClientePedido("CE_Numero")), "", rsClientePedido("CE_Numero"))
        mskCompl = IIf(IsNull(rsClientePedido("CE_Complemento")), "", rsClientePedido("CE_Complemento"))
        txtSituacao.Text = rsClientePedido("CE_Situacao")
        mskEndereco.Text = rsClientePedido("CE_Endereco")
        mskBairro.Text = rsClientePedido("CE_Bairro")
        mskCGCCPF.Text = rsClientePedido("CE_CGC")
        mskInscEst.Text = rsClientePedido("CE_InscricaoEstadual")
        mskMunicipio.Text = rsClientePedido("CE_Municipio")
        cmbPraca.AddItem "1 - SP"
        cmbPraca.AddItem "2 - Outro"
        For i = 0 To cmbPraca.ListCount - 1
            cmbPraca.ListIndex = i
            If Val(Mid(cmbPraca.Text, 1, 1)) = rsClientePedido("CE_Praca") Then
                cmbPraca.ListIndex = i
                Exit For
            End If
        Next i
        mskCEP.Text = rsClientePedido("CE_Cep")
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = UCase(rsClientePedido("CE_Estado")) Then
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
        mskDataCadastro.Text = Format(rsClientePedido("CE_DataCadastro"), "dd/mm/yyyy")
        mskTelefone.Text = rsClientePedido("CE_Telefone")
        mskFAX.Text = rsClientePedido("CE_Fax")
        mskCobrEnd.Text = rsClientePedido("CE_EnderecoCobranca")
        txtCobrNumero = IIf(IsNull(rsClientePedido("CE_NumeroCobranca")), "", rsClientePedido("CE_NumeroCobranca"))
        mskCobrCompl = IIf(IsNull(rsClientePedido("CE_ComplementoCobranca")), "", rsClientePedido("CE_ComplementoCobranca"))
        mskCobrBairro.Text = rsClientePedido("CE_BairroCobranca")
        mskCobrMunicipio.Text = rsClientePedido("CE_MunicipioCobranca")
        mskCobrEstado.Text = rsClientePedido("CE_EstadoCobranca")
        mskCobrCEP.Text = rsClientePedido("CE_CepCobranca")
        mskLimiteCred.Text = rsClientePedido("CE_LimiteCredito")
        mskDataLimite.Text = Format(IIf(IsNull(rsClientePedido("CE_DataLimiteCredito")), rsClientePedido("CE_DataCadastro"), rsClientePedido("CE_DataLimiteCredito")), "dd/mm/yyyy")
        mskJurosCartorio.Text = rsClientePedido("CE_JurosCartorio")
        mskDupliAberto.Text = "0,00"
        'mskQtdeDupliAberto.Text
        mskQtdeCompra.Text = rsClientePedido("CE_QuantidadeCompras")
        mskUltCompra.Text = rsClientePedido("CE_UltimaCompra")
        mskDataUltCompra.Text = Format(IIf(IsNull(rsClientePedido("CE_DataUltimaCompra")), rsClientePedido("CE_DataCadastro"), rsClientePedido("CE_DataUltimaCompra")), "dd/mm/yyyy")
        mskSaldoCompra.Text = Format(mskLimiteCred.Text - mskDupliAberto.Text, "0.00")
        mskMaiorCompra.Text = Format(rsClientePedido("CE_MaiorCompra"), "0.00")
        mskDataMaiorCompra.Text = Format(IIf(IsNull(rsClientePedido("CE_DataMaiorCompra")), rsClientePedido("CE_DataCadastro"), rsClientePedido("CE_DataMaiorCompra")), "dd/mm/yyyy")
        'mskDupliAtraso.Text=
        mskUltPagto.Text = Format(rsClientePedido("CE_UltimoPagamento"), "0.00")
        mskDataUltPagto.Text = Format(IIf(IsNull(rsClientePedido("CE_DataUltimoPagamento")), rsClientePedido("CE_DataCadastro"), rsClientePedido("CE_DataUltimoPagamento")), "dd/mm/yyyy")
        mskMaiorAtraso.Text = rsClientePedido("CE_MaiorAtraso") & " Dia(s)"
        mskEmail.Text = IIf(IsNull(rsClientePedido("CE_Email")), "", rsClientePedido("CE_Email"))
        mskDataNascimento.Enabled = True
        mskDataNascimento.Text = IIf(IsNull(rsClientePedido("CE_DataNascimento")), "01/01/1900", rsClientePedido("CE_DataNascimento"))
        mskCelular.Text = IIf(IsNull(rsClientePedido("CE_Celular")), 0, rsClientePedido("CE_Celular"))
        Call CarregaRamo(mskPessoa)
'        cmbRamoAtiv.ListIndex = IIf(IsNull(rsClientePedido("CE_RamoAtividade")), 0, rsClientePedido("CE_RamoAtividade") - 1)
        Call CarregaSegmento(Val(cmbRamoAtiv.Text))
'        cmbSegmento.ListIndex = IIf(IsNull(rsClientePedido("CE_Segmento")), 0, rsClientePedido("CE_Segmento") - 1)
        mskCodCli.Enabled = False
        mskRazao.SelStart = 0
        mskRazao.SelLength = Len(mskRazao.Text)
        
        
    End If
    rsClientePedido.Close
End Function
    
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' If Lblcep.ForeColor = &H80FF& Then
   '     Lblcep.ForeColor = &HFF0000
   ' End If
   ' Lblcep.MousePointer = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   '  If Lblcep.ForeColor = &H80FF& Then
   '     Lblcep.ForeColor = &HFF0000
   ' End If
   ' Lblcep.MousePointer = 0

End Sub


Private Sub lblCep_Click()

    Shell "start www.maplink.uol.com.br/endereco.asp", vbHide
    
End Sub

Private Sub lblCep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblCEP.ForeColor = &H80FF&
    lblCEP.MousePointer = 14

End Sub



Private Sub mskBairro_GotFocus()
  mskBairro.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskBairro_KeyPress(KeyAscii As Integer)

lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskBairro.Text <> "" Then
            mskBairro.Text = UCase(mskBairro.Text)
            ProximoCampo mskMunicipio
            SelecionaCampo mskMunicipio
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCompl
        SelecionaCampo mskCompl
    End If
        
    
End Sub

Private Sub mskBairro_LostFocus()
   mskBairro.BackColor = vbWhite
   mskBairro.Text = UCase(mskBairro.Text)

End Sub

Private Sub mskCelular_GotFocus()
    mskCelular.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCelular_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskTelefone.Text <> "" Then
            ProximoCampo mskFAX
            SelecionaCampo mskFAX
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskTelefone
    End If

End Sub

Private Sub mskCelular_LostFocus()
   mskCelular.BackColor = vbWhite
End Sub

Private Sub mskCEP_GotFocus()
   mskCEP.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCEP_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCEP.Text <> "" Then
            If IsNumeric(mskCEP.Text) = True Then
                If ConsultaCep(mskCEP.Text) = False Then
                    ProximoCampo mskEndereco
                    SelecionaCampo mskEndereco
                Else
                    ProximoCampo txtNumero
                End If
            Else
                'MsgBox "Favor use somente numeros no CEP", vbCritical, "Atenção"
                lblInfo.Caption = "Favor use somente números no CEP"
                SelecionaCampo mskCEP
            End If
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskInscEst
        SelecionaCampo mskInscEst
    End If
    
        
End Sub

Private Sub mskCEP_LostFocus()
    mskCEP.BackColor = vbWhite
    If mskCEP.Text <> "" Then
        If IsNumeric(mskCEP.Text) = True Then
            If ConsultaCep(mskCEP.Text) = False Then
                ProximoCampo mskEndereco
                SelecionaCampo mskEndereco
            Else
                ProximoCampo txtNumero
            End If
        Else
            'MsgBox "Favor use somente numeros", vbCritical, "Atenção"
            lblInfo.Caption = "Favor use somente números no CEP"
            ProximoCampo mskCEP
            SelecionaCampo mskCEP
        End If
    End If
    
    If Len(mskEndereco.Text) > 40 Then                'cesar
        'MsgBox "Favor abreviar o Endereço do cliente", vbCritical, "Atenção" 'cesar
        lblInfo.Caption = "Favor abreviar o Endereço do cliente"
        mskEndereco.SetFocus                          'cesar
    End If

End Sub

Private Sub mskCGCCPF_GotFocus()
    mskCGCCPF.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCGCCPF_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCGCCPF.Text <> "" And IsNumeric(mskCGCCPF) = True Then
            ProximoCampo mskPessoa
            mskCGCCPF.Text = Right(String(15, "0") & mskCGCCPF.Text, 15)
        Else
            mskCGCCPF.SelStart = 0
            mskCGCCPF.SelLength = Len(mskCGCCPF.Text)
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskRazao
        SelecionaCampo mskRazao
    End If
    
    
End Sub



Private Sub mskCGCCPF_LostFocus()
   mskCGCCPF.BackColor = vbWhite
End Sub

Private Sub mskCobrBairro_GotFocus()
    mskCobrBairro.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCobrBairro_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCobrBairro.Text <> "" Then
            mskCobrBairro.Text = UCase(mskCobrBairro.Text)
            ProximoCampo mskCobrMunicipio
            SelecionaCampo mskCobrMunicipio
        Else
            mskCobrBairro.Text = UCase(mskBairro.Text)
            ProximoCampo mskCobrMunicipio
            SelecionaCampo mskCobrMunicipio
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCobrCEP
        SelecionaCampo mskCobrCEP
    End If
    
End Sub

Private Sub mskCobrBairro_LostFocus()
    mskCobrBairro.BackColor = vbWhite
    mskCobrBairro.Text = UCase(mskCobrBairro.Text)

End Sub

Private Sub mskCobrCEP_GotFocus()
    mskCobrCEP.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCobrCEP_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCobrCEP.Text <> "" Then
            mskCobrCEP.Text = UCase(mskCobrCEP.Text)
            ProximoCampo mskCobrBairro
            SelecionaCampo mskCobrBairro
        Else
            mskCobrCEP.Text = UCase(mskCEP.Text)
            ProximoCampo mskCobrBairro
            SelecionaCampo mskCobrBairro
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCobrCompl
        SelecionaCampo mskCobrCompl
    End If
    
End Sub

Private Sub mskCobrCEP_LostFocus()
   mskCobrCEP.BackColor = vbWhite
End Sub

Private Sub mskCobrCompl_GotFocus()
    mskCobrCompl.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCobrCompl_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCobrCompl.Text <> "" Then
            mskCobrCompl.Text = UCase(mskCobrCompl.Text)
            ProximoCampo mskCobrCEP
            SelecionaCampo mskCobrCEP
        Else
            mskCobrCompl.Text = UCase(mskCompl.Text)
            ProximoCampo mskCobrCEP
            SelecionaCampo mskCobrCEP
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo txtCobrNumero
        SelecionaCampo txtCobrNumero
    End If

End Sub

Private Sub mskCobrCompl_LostFocus()
    mskCobrCompl.BackColor = vbWhite
    mskCobrCompl.Text = UCase(mskCobrCompl.Text)

End Sub

Private Sub mskCobrEnd_GotFocus()
    mskCobrEnd.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCobrEnd_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCobrEnd.Text <> "" Then
            mskCobrEnd.Text = UCase(mskCobrEnd.Text)
            ProximoCampo txtCobrNumero
            SelecionaCampo txtCobrNumero
        Else
            mskCobrEnd.Text = UCase(mskEndereco.Text)
            ProximoCampo txtCobrNumero
            SelecionaCampo txtCobrNumero
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    End If
    
End Sub

Private Sub mskCobrEnd_LostFocus()
    mskCobrEnd.BackColor = vbWhite
    mskCobrEnd.Text = UCase(mskCobrEnd.Text)

End Sub

Private Sub mskCobrEstado_GotFocus()
    mskCobrEstado.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCobrEstado_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCobrEstado.Text <> "" Then
            mskCobrEstado.Text = UCase(mskCobrEstado.Text)
            ProximoCampo mskLimiteCred
            SelecionaCampo mskLimiteCred
        Else
            mskCobrEstado.Text = UCase(cmbUf.Text)
            ProximoCampo mskLimiteCred
            SelecionaCampo mskLimiteCred
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCobrMunicipio
        SelecionaCampo mskCobrMunicipio
    End If
    
End Sub

Private Sub mskCobrEstado_LostFocus()

    mskCobrEstado.BackColor = vbWhite
    mskCobrEstado.Text = UCase(mskCobrEstado.Text)

End Sub

Private Sub mskCobrMunicipio_GotFocus()
    mskCobrMunicipio.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskCobrMunicipio_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCobrMunicipio.Text <> "" Then
            mskCobrMunicipio.Text = UCase(mskCobrMunicipio.Text)
            ProximoCampo mskCobrEstado
            SelecionaCampo mskCobrEstado
        Else
            mskCobrMunicipio.Text = UCase(mskMunicipio.Text)
            ProximoCampo mskCobrEstado
            SelecionaCampo mskCobrEstado
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCobrBairro
        SelecionaCampo mskCobrBairro
    End If
    
End Sub

Private Sub mskCobrMunicipio_LostFocus()

    mskCobrMunicipio.BackColor = vbWhite
    mskCobrMunicipio.Text = UCase(mskCobrMunicipio.Text)

End Sub

Private Sub mskCodCli_GotFocus()

    mskCodCli.SelStart = 0
    mskCodCli.SelLength = Len(mskCodCli.Text)

End Sub

Private Sub mskCodCli_KeyPress(KeyAscii As Integer)
    lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCodCli.Text <> "" And IsNumeric(mskCodCli.Text) = True Then
            ProximoCampo mskRazao
            SelecionaCampo mskRazao
        Else
            ProximoCampo mskCodCli
            SelecionaCampo mskCodCli
        End If
    ElseIf KeyAscii = 27 Then
        Unload Me
        frmConsCliente.txtPesquisaCliente.SetFocus
    End If
            
End Sub

Private Sub mskCodCli_LostFocus()
    
    DescricaoOperacao "Pesquisando Cliente"
    If VerificaClienteExite(mskCodCli.Text) = True Then
        PreencheDadosCliente mskCodCli.Text
        DescricaoOperacao "Pronto"
    End If
    
End Sub

Private Sub mskCompl_KeyPress(KeyAscii As Integer)
    lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskCompl.Text <> "" Then
            mskCompl.Text = UCase(mskCompl.Text)
            ProximoCampo mskBairro
            SelecionaCampo mskBairro
        Else
            mskCompl.Text = UCase(mskCompl.Text)
            ProximoCampo mskCompl
            SelecionaCampo mskCompl
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo txtNumero
        SelecionaCampo txtNumero
    End If

End Sub

Private Sub mskCompl_LostFocus()
    If Len(mskEndereco.Text & mskCompl.Text) > 40 Then
        'MsgBox "Favor abreviar o Complemento ou Endereço do cliente", vbCritical, "Atenção" 'cesar
        lblInfo.Caption = "Favor abreviar o Complemento ou Endereço do cliente"
        mskCompl.SetFocus
    Else
        SelecionaCampo txtNumero
    End If
    
    mskCobrCompl.Text = UCase(mskCobrCompl.Text)

End Sub
Private Sub mskDataCadastro_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskDataCadastro.Text <> "" Then
            ProximoCampo mskCobrEnd
            SelecionaCampo mskCobrEnd
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskEmail
        SelecionaCampo mskEmail
    End If
    
End Sub

Private Sub mskDataNascimento_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        ProximoCampo cmbRamoAtiv
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskFAX
        SelecionaCampo mskFAX
    End If

End Sub

Private Sub mskEmail_KeyPress(KeyAscii As Integer)
    lblInfo.Caption = ""
    If KeyAscii = 13 Then
        mskDataCadastro.Text = Format(Date, "dd/mm/yyyy")
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    ElseIf KeyAscii = 27 Then
        ProximoCampo cmbSegmento
        SelecionaCampo cmbSegmento
    End If
    
End Sub

Private Sub mskEmail_LostFocus()

    'mskEmail.Text = UCase(mskEmail.Text)

End Sub

Private Sub mskEndereco_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskEndereco.Text <> "" Then
            mskEndereco.Text = UCase(mskEndereco.Text)
            ProximoCampo txtNumero
            SelecionaCampo txtNumero
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCEP
        SelecionaCampo mskCEP
    End If
    
End Sub

Private Sub mskEndereco_LostFocus()
    
    mskEndereco.Text = UCase(mskEndereco.Text)

End Sub



Private Sub mskFAX_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskFAX.Text <> "" Then
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

Private Sub mskInscEst_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskInscEst.Text <> "" Then
            If IsNumeric(mskInscEst.Text) = True Then
                ProximoCampo mskCEP
                SelecionaCampo mskCEP
                mskInscEst.Text = Right(String(15, "0") & mskInscEst.Text, 15)
            Else
                mskInscEst.SelStart = 0
                mskInscEst.SelLength = Len(mskInscEst.Text)
            End If
        Else
            mskInscEst.Text = "00000000000000"
            ProximoCampo mskCEP
            SelecionaCampo mskCEP
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo chkCart
    End If
    
End Sub

Private Sub mskInscEst_LostFocus()

    mskInscEst.Text = Right(String(15, "0") & mskInscEst.Text, 15)

End Sub

Private Sub mskLimiteCred_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskLimiteCred.Text <> "" Then
            mskDataLimite.Text = Format(Date, "dd/mm/yyyy")
            mskJurosCartorio.Text = "0,00"
            mskDupliAberto.Text = "0,00"
            mskQtdeDupliAberto.Text = 0
            mskQtdeCompra.Text = 0
            mskUltCompra.Text = "0,00"
            mskDataUltCompra.Text = Format(Date, "dd/mm/yyyy")
            mskSaldoCompra.Text = Format(mskLimiteCred.Text, "0.00")
            mskMaiorCompra.Text = "0,00"
            mskDataMaiorCompra.Text = Format(Date, "dd/mm/yyyy")
            mskDupliAtraso.Text = 0
            mskUltPagto.Text = "0,00"
            mskDataUltPagto.Text = Format(Date, "dd/mm/yyyy")
            mskMaiorAtraso.Text = 0
            ProximoCampo mskDataLimite
            SelecionaCampo mskDataLimite
            ProximoCampo cmdGravar
            cmdGravar.Enabled = True
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCobrCEP
        SelecionaCampo mskCobrCEP
    End If
    
End Sub

Private Sub mskMunicipio_KeyPress(KeyAscii As Integer)
    lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskMunicipio.Text <> "" Then
            mskMunicipio.Text = UCase(mskMunicipio.Text)
            ProximoCampo cmbPraca
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskBairro
        SelecionaCampo mskBairro
    End If
    
End Sub

Private Sub mskMunicipio_LostFocus()

    mskMunicipio.Text = UCase(mskMunicipio.Text)

End Sub

Private Sub mskPessoa_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If mskPessoa.Text <> "" Then
            mskPessoa.Text = UCase(mskPessoa.Text)
            If Mid(UCase(mskPessoa.Text), 1, 1) = "J" Then
                mskPessoa.Text = "Juridica"
                Call CarregaRamo("Juridica")
            ElseIf Mid(UCase(mskPessoa.Text), 1, 1) = "U" Or Mid(UCase(mskPessoa), 1, 2) = "FU" Then
                mskPessoa.Text = "Func"
                Call CarregaRamo("Func")
            ElseIf Mid(UCase(mskPessoa.Text), 1, 1) = "F" Then
                mskPessoa.Text = "Fisica"
                Call CarregaRamo("Fisica")
            ElseIf Mid(UCase(mskPessoa.Text), 1, 1) = "A" Then
                mskPessoa.Text = "Altarquia"
                Call CarregaRamo("Altarquia")
            Else
                mskPessoa.SelStart = 0
                mskPessoa.SelLength = Len(mskPessoa.Text)
            End If
            If txtSituacao.Enabled = True Then
                ProximoCampo txtSituacao
                SelecionaCampo txtSituacao
            Else
                ProximoCampo chkCart
            End If
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCGCCPF
        SelecionaCampo mskCGCCPF
    End If
    
End Sub

Private Sub mskPessoa_LostFocus()

    If mskPessoa.Text <> "" Then
        mskPessoa.Text = UCase(mskPessoa.Text)
        If Mid(UCase(mskPessoa.Text), 1, 1) = "J" Then
            mskPessoa.Text = "Juridica"
        ElseIf Mid(UCase(mskPessoa.Text), 1, 1) = "U" Or Mid(UCase(mskPessoa), 1, 2) = "FU" Then
            mskPessoa.Text = "Func"
        ElseIf Mid(UCase(mskPessoa.Text), 1, 1) = "F" Then
            mskPessoa.Text = "Fisica"
        ElseIf Mid(UCase(mskPessoa.Text), 1, 1) = "A" Then
            mskPessoa.Text = "Altarquia"
        Else
            mskPessoa.SelStart = 0
            mskPessoa.SelLength = Len(mskPessoa.Text)
        End If
    End If

End Sub



Private Sub mskPraca_LostFocus()

    'mskPraca.Text = UCase(mskPraca.Text)

End Sub

Private Sub mskRazao_GotFocus()
  mskRazao.BackColor = RGB(220, 235, 255)
End Sub

Private Sub mskRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
        If mskCodCli.Enabled = False Then
            Unload Me
            frmConsCliente.txtPesquisaCliente.SetFocus
        Else
            ProximoCampo mskCodCli
            SelecionaCampo mskCodCli
        End If
    ElseIf KeyCode = 192 Then
        If Len(mskRazao.Text) = 0 Then
            mskRazao.Text = ""
        Else
            mskRazao.Text = Mid(mskRazao.Text, 1, Len(mskRazao.Text) - 1)
        End If
    End If
    
End Sub

Private Sub mskRazao_KeyPress(KeyAscii As Integer)
    lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskRazao.Text <> "" Then
            ProximoCampo mskCGCCPF
            SelecionaCampo mskCGCCPF
            mskRazao.Text = UCase(mskRazao.Text)
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

Private Sub mskRazao_LostFocus()
    mskRazao.BackColor = vbWhite
    mskRazao.Text = UCase(mskRazao.Text)
    
End Sub

Private Sub mskTelefone_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If mskTelefone.Text <> "" Then
            ProximoCampo mskCelular
            SelecionaCampo mskCelular
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo cmbUf
    End If
    
End Sub

Private Sub Picture1_Click()

    Shell "START WWW.MAPLINK.UOL.COM.BR/endereco.asp", vbHide

End Sub




Private Sub TabStrip1_Click()
Dim i As Integer
fraCliente.ZOrder
i = TabStrip1.SelectedItem.Index
If i = 1 Then
   fraCliente.ZOrder
Else
  If i = 2 Then
   fraCobranca.ZOrder
  Else
   fraCredito.ZOrder
  End If
End If

End Sub

Private Sub txtNumero_GotFocus()
    
    If Len(mskEndereco.Text) > 40 Then                'cesar
        'MsgBox "Favor abreviar o Endereço do cliente", vbCritical, "Atenção" 'cesar
        lblInfo.Caption = "Favor abreviar o Endereço do cliente"
        mskEndereco.SetFocus                          'cesar
    End If
    
    SelecionaCampo txtNumero
    
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If txtNumero.Text <> "" Then
            txtNumero.Text = UCase(txtNumero.Text)
            ProximoCampo mskCompl
            SelecionaCampo mskCompl
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskEndereco
        SelecionaCampo mskEndereco
    End If
    
End Sub

Private Sub txtCobrNumero_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        If txtCobrNumero.Text <> "" Then
            txtCobrNumero.Text = UCase(txtCobrNumero.Text)
            ProximoCampo mskCobrCompl
            SelecionaCampo mskCobrCompl
        Else
            txtCobrNumero.Text = UCase(txtNumero.Text)
            ProximoCampo mskCobrCompl
            SelecionaCampo mskCobrCompl
        End If
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskCobrEnd
        SelecionaCampo mskCobrEnd
    End If

End Sub

Private Sub txtCobrNumero_LostFocus()

    txtCobrNumero.Text = UCase(txtCobrNumero.Text)

End Sub

Private Sub txtSituacao_KeyPress(KeyAscii As Integer)
lblInfo.Caption = ""
    If KeyAscii = 13 Then
        ProximoCampo chkCart
        txtSituacao.Text = UCase(txtSituacao.Text)
        txtSituacao.Text = 0
    ElseIf KeyAscii = 27 Then
        ProximoCampo mskPessoa
        SelecionaCampo mskPessoa
    End If
    
End Sub

Function VerificaCamposNulos() As Boolean
        
    If mskCodCli.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskCodCli
        SelecionaCampo mskCodCli
    ElseIf mskRazao.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskRazao
        SelecionaCampo mskRazao
    ElseIf mskPessoa.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskPessoa
        SelecionaCampo mskPessoa
    ElseIf mskEndereco.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskEndereco
        SelecionaCampo mskEndereco
    ElseIf txtNumero.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo txtNumero
        SelecionaCampo txtNumero
    ElseIf mskBairro.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskBairro
        SelecionaCampo mskBairro
    ElseIf mskCGCCPF.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskCGCCPF
        SelecionaCampo mskCGCCPF
    ElseIf mskInscEst.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskInscEst
        SelecionaCampo mskInscEst
    ElseIf mskMunicipio.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskMunicipio
        SelecionaCampo mskMunicipio
    ElseIf cmbPraca.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo cmbPraca
        SelecionaCampo cmbPraca
    ElseIf cmbUf.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo cmbUf
    ElseIf mskCEP.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskCEP
        SelecionaCampo mskCEP
    ElseIf mskTelefone.Text = "" Then
        VerificaCamposNulos = False
        ProximoCampo mskTelefone
        SelecionaCampo mskTelefone
    ElseIf IsDate(mskDataCadastro.Text) = False Then
        VerificaCamposNulos = False
        ProximoCampo mskDataCadastro
        SelecionaCampo mskDataCadastro
    ElseIf mskCobrEstado.Text = "" Then
        mskCobrEstado.Text = cmbUf.Text
        VerificaCamposNulos = True
    ElseIf mskCobrCEP.Text = "" Then
        mskCobrCEP.Text = mskCEP.Text
        VerificaCamposNulos = True
    ElseIf IsDate(mskDataNascimento.Text) = False Then
        VerificaCamposNulos = True
        ProximoCampo mskDataNascimento
    ElseIf mskLimiteCred.Text < 0 Then
        VerificaCamposNulos = False
    Else
        VerificaCamposNulos = True
    End If
    mskCobrEnd.Text = mskEndereco.Text
    txtCobrNumero.Text = txtNumero.Text
    mskCobrBairro.Text = mskBairro.Text
    mskCobrMunicipio.Text = mskMunicipio.Text
    mskCobrCompl.Text = mskCompl.Text
    If VerificaCamposNulos = False Then
       ' MsgBox "Preencha todos os campos obrigatorios", vbCritical, "Atenção"
        lblInfo.Caption = "Preencha todos os campos obrigatorios"
        Exit Function
    End If
'    mskMunicipio.Enabled = True
'    mskPraca.Enabled = True
'    mskCEP.Enabled = True
'    mskEstado.Enabled = True
'    mskDataCadastro.Enabled = True
'    mskTelefone.Enabled = True
'    mskFAX.Enabled = True
'    mskCobrEnd.Enabled = True
'    mskCobrBairro.Enabled = True
'    mskCobrMunicipio.Enabled = True
'    mskCobrEstado.Enabled = True
'    mskCobrCEP.Enabled = True
'    mskLimiteCred.Enabled = True
'    mskDataLimite.Enabled = True
'    mskJurosCartorio.Enabled = True
'    mskDupliAberto.Enabled = True
'    mskQtdeDupliAberto.Enabled = True
'    mskQtdeCompra.Enabled = True
'    mskUltCompra.Enabled = True
'    mskDataUltCompra.Enabled = True
'    mskSaldoCompra.Enabled = True
'    mskMaiorCompra.Enabled = True
'    mskDataMaiorCompra.Enabled = True
'    mskDupliAtraso.Enabled = True
'    mskUltPagto.Enabled = True
'    mskDataUltPagto.Enabled = True
'    mskMaiorAtraso.Enabled = True
'    mskRazao.TabIndex = 0

    
End Function


Function GravaCliente(ByVal Codigo As Double, ByVal Tipo As Integer) As Boolean
    Dim Pessoa As String
        
    GravaCliente = False
    DescricaoOperacao "Gravando Cliente"
    If Trim(UCase(mskPessoa.Text)) = "FISICA" Then
        Pessoa = "F"
    ElseIf Trim(UCase(mskPessoa.Text)) = "JURIDICA" Then
        Pessoa = "J"
    ElseIf Trim(UCase(mskPessoa.Text)) = "FUNC" Then
        Pessoa = "U"
    ElseIf Trim(UCase(mskPessoa.Text)) = "ALTARQUIA" Then
        Pessoa = "A"
    End If
    If txtNumero.Text = "" Then
        txtNumero.Text = 0
    End If
    If txtCobrNumero.Text = "" Then
        txtCobrNumero.Text = 0
    End If
    If Tipo = 1 Then 'Novo Cliente
        On Error Resume Next
        rdoCNLoja.BeginTrans
        SQL = ""
        SQL = "Insert into Cliente (CE_CodigoCliente,CE_CGC,CE_InscricaoEstadual,CE_Razao,CE_Endereco,CE_Bairro,CE_Municipio, " _
            & "CE_Estado, CE_CEP,CE_Telefone,CE_Fax,CE_EMail,CE_TipoPessoa,CE_Praca,CE_PagamentoCarteira,CE_EnderecoCobranca,CE_BairroCobranca, " _
            & "CE_MunicipioCobranca,CE_EstadoCobranca,CE_CEPCobranca,CE_LimiteCredito,CE_DataLimiteCredito,CE_MaiorCompra,CE_DataMaiorCompra,CE_UltimaCompra, " _
            & "CE_DataUltimaCompra,CE_UltimoPagamento,CE_DataUltimoPagamento,CE_MaiorAtraso,CE_QuantidadeCompras,CE_JurosCartorio,CE_DataCadastro,CE_DataCancelamento,CE_Alteracao,CE_Situacao,CE_HoraManutencao,CE_Numero,CE_NumeroCobranca,CE_Complemento,CE_ComplementoCobranca,CE_Celular,CE_DataNascimento,CE_Segmento,CE_RamoAtividade,CE_Loja) " _
            & "Values (" & Codigo & ",'" & mskCGCCPF.Text & "','" & mskInscEst.Text & "','" & mskRazao.Text & "','" & mskEndereco.Text & "','" & mskBairro.Text & "','" & mskMunicipio.Text & "', " _
            & "'" & cmbUf.Text & "','" & mskCEP.Text & "','" & mskTelefone.Text & "','" & mskFAX.Text & "','" & mskEmail.Text & "','" & Mid(Pessoa, 1, 1) & "'," & Mid(cmbPraca.Text, 1, 1) & ", 0,'" & mskCobrEnd.Text & "','" & mskCobrBairro.Text & "', " _
            & "'" & mskCobrMunicipio.Text & "','" & mskCobrEstado.Text & "','" & mskCobrCEP.Text & "'," & ConverteVirgula(Format(mskLimiteCred.Text, "0.00")) & ",'" & Format(mskDataLimite.Text, "mm/dd/yyyy") & "',0,'" & Format(mskDataMaiorCompra.Text, "mm/dd/yyyy") & "',0, " _
            & "'" & Format(mskDataUltCompra.Text, "mm/dd/yyyy") & "',0,'" & Format(mskDataUltPagto.Text, "mm/dd/yyyy") & "',0,0,0,'" & Format(Date, "mm/dd/yyyy") & "','00:00:00','A',0,getdate(),'" & txtNumero & "','" & txtCobrNumero & "', '" & mskCompl.Text & "', '" & mskCobrCompl.Text & "','" & mskCelular & "', '" & Format(IIf(mskDataNascimento = "__/__/____", "01/01/1900", mskDataNascimento), "mm/dd/yyyy") & "', " & Val(cmbSegmento.Text) & ", " & Val(cmbRamoAtiv.Text) & ", '" & AchaLojaControle & "') "
            rdoCNLoja.Execute (SQL)
        If Err.Number = 0 Then
            rdoCNLoja.CommitTrans
            rdoCNLoja.BeginTrans
            SQL = ""
            SQL = "update Controle set CT_SeqCliente = CT_SeqCliente + 1"
                rdoCNLoja.Execute (SQL)
            rdoCNLoja.CommitTrans
            frmConsCliente.txtPesquisaCliente.Text = mskCodCli.Text
            Screen.MousePointer = 0
            'MsgBox "Cliente cadastrado com sucesso", vbInformation, "Sucesso"
            lblInfo.Caption = "Cliente cadastrado com sucesso"
            GravaCliente = True
        Else
            rdoCNLoja.RollbackTrans
            Screen.MousePointer = 0
            'MsgBox "Erro cadastrando cliente", vbCritical, "ERRO"
            lblInfo.Caption = "Erro cadastrando cliente"
            GravaCliente = False
        End If
    ElseIf Tipo = 2 Then 'Alteracao
        On Error Resume Next
        rdoCNLoja.BeginTrans
        SQL = ""
        SQL = "Update Cliente set CE_CGC='" & mskCGCCPF.Text & "',CE_InscricaoEstadual='" & mskInscEst.Text & "',CE_Razao='" & mskRazao.Text & "',CE_Endereco='" & mskEndereco.Text & "',CE_Bairro='" & mskBairro.Text & "',CE_Municipio='" & mskMunicipio.Text & "', " _
            & "CE_Estado='" & cmbUf.Text & "' , CE_CEP='" & mskCEP.Text & "',CE_Telefone='" & mskTelefone.Text & "',CE_Fax='" & mskFAX.Text & "',CE_EMail='" & mskEmail.Text & "',CE_TipoPessoa='" & Pessoa & "',CE_Praca=" & Mid(cmbPraca.Text, 1, 1) & ",CE_EnderecoCobranca='" & mskCobrEnd.Text & "',CE_BairroCobranca='" & mskCobrBairro.Text & "', " _
            & "CE_MunicipioCobranca='" & mskCobrMunicipio.Text & "',CE_EstadoCobranca='" & mskCobrEstado.Text & "',CE_CEPCobranca='" & mskCobrCEP.Text & "',CE_Alteracao='S',CE_HoraManutencao=getdate(), CE_Numero='" & txtNumero.Text & "', CE_Celular = '" & mskCelular & "', CE_DataNascimento = '" & Format(mskDataNascimento, "mm/dd/yyyy") & "', CE_Segmento = " & Val(cmbSegmento.Text) & ", CE_RamoAtividade = " & Val(cmbRamoAtiv.Text) & ", CE_Loja = '" & AchaLojaControle & "' " _
            & "where CE_CodigoCliente=" & Codigo & ""
            rdoCNLoja.Execute (SQL)
        If Err.Number = 0 Then
            rdoCNLoja.CommitTrans
            Screen.MousePointer = 0
            'MsgBox "Cliente alterado com sucesso", vbInformation, "Sucesso"
            lblInfo.Visible = True
            lblInfo.Caption = "Cliente alterado com sucesso"
            GravaCliente = True
        Else
            rdoCNLoja.RollbackTrans
            Screen.MousePointer = 0
            'MsgBox "Erro alterando dados do cliente", vbCritical, "ERRO"
            lblInfo.Caption = "Cliente alterado com sucesso"
        End If
    End If
    DescricaoOperacao "Pronto"
        
End Function


Function VerificaClienteExite(ByVal Cliente As Double) As Boolean
   ' Dim rdoCliente As rdoResultset

    SQL = ""
    SQL = "Select CE_CodigoCliente from Cliente " _
        & "where CE_CodigoCliente=" & Cliente & " "
        rdoCliente.CursorLocation = adUseClient
        rdoCliente.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    'Set rdoCliente = rdoCNLoja.OpenResultset(SQL)
    If Not rdoCliente.EOF Then
        VerificaClienteExite = True
    Else
        VerificaClienteExite = False
    End If
    rdoCliente.Close
End Function

Function PreencheUF()
    
    SQL = ""
    SQL = "Select UF_Estado from Estados " _
        & "Order by UF_Estado"
    rdoUf.CursorLocation = adUseClient
    rdoUf.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    'Set rdoUf = rdoCNLoja.OpenResultset(SQL)
    If Not rdoUf.EOF Then
        ExisteUF = True
    Else
        ExisteUF = False
    End If
    
End Function

Function ConsultaCep(ByVal Cep As String) As Boolean
    
    
    mskEndereco.Text = ""
    mskMunicipio.Text = ""
    mskBairro.Text = ""
    SQL = ""
    SQL = "Select * from Cep " _
        & "where Cep= '" & Cep & "'"
     rdoCep.CursorLocation = adUseClient
     rdoCep.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    'Set rdoCep = rdoCnLojaBach.OpenResultset(SQL)
    If Not rdoCep.EOF Then
        ConsultaCep = True
        mskEndereco.Text = rdoCep("Logradouro")
        mskMunicipio.Text = rdoCep("Municipio")
        mskBairro.Text = rdoCep("Bairroa")
        For i = 0 To cmbUf.ListCount
            cmbUf.ListIndex = i
            If cmbUf.Text = UCase(rdoCep("UF")) Then
                cmbUf.ListIndex = i
                Exit For
            End If
        Next i
    Else
        ConsultaCep = False
    End If
    rdoCep.Close
End Function

Function CarregaRamo(ByVal Pessoa As String)
    
    
    cmbRamoAtiv.Clear
    
    If Pessoa = "Juridica" Or Pessoa = "JURIDICA" Then
        Pessoa = "J"
    ElseIf Pessoa = "Fisica" Or Pessoa = "FISICA" Then
        Pessoa = "F"
    ElseIf Pessoa = "Func" Or Pessoa = "FUNC" Then
        Pessoa = "U"
    ElseIf Pessoa = "Altarquia" Or Pessoa = "ALTARQUIA" Then
        Pessoa = "A"
    End If
    
    SQL = ""
    SQL = "Select * From RamoAtividade Where RMO_Pessoa = '" & Pessoa & "'"
     rdoRamo.CursorLocation = adUseClient
     rdoRamo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    'Set rdoRamo = rdoCNLoja.OpenResultset(SQL)
    
    If Not rdoRamo.EOF Then
        Do While Not rdoRamo.EOF
            cmbRamoAtiv.AddItem rdoRamo("RMO_Codigo") & " - " & rdoRamo("RMO_DescricaoRamo")
            rdoRamo.MoveNext
        Loop
        cmbRamoAtiv.ListIndex = 0
        
        Call CarregaSegmento(Val(cmbRamoAtiv.Text))
    End If
    rdoRamo.Close
End Function

Function CarregaSegmento(ByVal CodigoRamo As Integer)
    
    
    cmbSegmento.Clear
    
    SQL = ""
    SQL = "Select * From Segmento Where SEG_RamoAtividade = " & CodigoRamo & ""
     rdoSegmento.CursorLocation = adUseClient
     rdoSegmento.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    'Set rdoSegmento = rdoCNLoja.OpenResultset(SQL)
    
    If Not rdoSegmento.EOF Then
        Do While Not rdoSegmento.EOF
            cmbSegmento.AddItem rdoSegmento("SEG_CodigoSegmento") & " - " & rdoSegmento("SEG_Descricao")
            rdoSegmento.MoveNext
        Loop
        cmbSegmento.ListIndex = 0
    End If
    rdoSegmento.Close
End Function
Function PesquisaCliente(ByVal TipoPesquisa As Integer, ByVal Cliente As String, ByRef NomerdoResultset) As Boolean

'
'--------------------------------Pesquisa Pelo Codigo do Cliente (1)-------------------------
'
    'DescricaoOperacao "Pesquisando Cliente"
    If TipoPesquisa = 1 Then
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from Cliente " _
            & "where CE_CodigoCliente = " & Cliente & " "
            NomerdoResultset.CursorLocation = adUseClient
            NomerdoResultset.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = rdoCNLoja.OpenResultset(SQL)

'
'-------------------------------Pesquisa por cgc ou cpf (2) ---------------------------------
'
    ElseIf TipoPesquisa = 2 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from Cliente " _
            & "where CE_Cgc = '" & Cliente & "' "
            NomerdoResultset.CursorLocation = adUseClient
            NomerdoResultset.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = rdoCNLoja.OpenResultset(SQL)
    
'
'-------------------------------Pesquisa Pelo Nome Cliente (3) ---------------------------------
'
    ElseIf TipoPesquisa = 3 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_razao,CE_CodigoCliente from Cliente " _
            & "where CE_Razao like '" & UCase(Cliente) & "%' order by CE_Razao"
            NomerdoResultset.CursorLocation = adUseClient
            NomerdoResultset.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = rdoCNLoja.OpenResultset(SQL)
    
'
'-------------------------------Pesquisa Cliente Tela frmCadCliente(4) --------------------------
'
    ElseIf TipoPesquisa = 4 Then
        SQL = ""
        SQL = ""
        SQL = "Select * from Cliente " _
            & "where CE_CodigoCliente = " & Cliente & " order by CE_CodigoCliente"
            NomerdoResultset.CursorLocation = adUseClient
            NomerdoResultset.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            'Set NomerdoResultset = rdoCNLoja.OpenResultset(SQL)
    
    Else
        Exit Function
    End If
    If Not NomerdoResultset.EOF Then
        PesquisaCliente = True
    Else
        PesquisaCliente = False
    End If
    'DescricaoOperacao "Pronto"
    
End Function

Private Sub LimpaForm()
Dim cObjeto As Control
  
  For Each cObjeto In Me.Controls
      If (TypeOf cObjeto Is TextBox) Then
        cObjeto.Text = ""
      End If
  Next
  
  For Each cObjeto In Me.Controls
      If (TypeOf cObjeto Is ComboBox) Then
        cObjeto.Clear
      End If
  Next

End Sub


