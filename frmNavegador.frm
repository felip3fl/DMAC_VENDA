VERSION 5.00
Begin VB.Form frmNavegador 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   5895
   ClientLeft      =   2655
   ClientTop       =   3540
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmrDadosCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   4830
      Visible         =   0   'False
      Width           =   10620
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   160
         MaxLength       =   14
         TabIndex        =   6
         Top             =   360
         Width           =   1635
      End
      Begin VB.TextBox txtRazaoSocial 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1875
         MaxLength       =   40
         TabIndex        =   5
         Top             =   360
         Width           =   6450
      End
      Begin VB.TextBox txtInscricaoEstadual 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   8400
         MaxLength       =   15
         TabIndex        =   4
         Top             =   360
         Width           =   1470
      End
      Begin VB.TextBox txtUf 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   9945
         MaxLength       =   15
         TabIndex        =   3
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   10
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Nome"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   1875
         TabIndex        =   9
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Inscr. Est."
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   8400
         TabIndex        =   8
         Top             =   120
         Width           =   705
      End
      Begin VB.Label lblClite 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "UF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   11
         Left            =   9945
         TabIndex        =   7
         Top             =   120
         Width           =   210
      End
   End
   Begin VB.Frame FrameNavegador 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   15015
      Begin Project1.IEalt IEalt1 
         Left            =   120
         Top             =   120
         _ExtentX        =   26061
         _ExtentY        =   7435
      End
   End
   Begin Project1.chameleonButton cmdRetornar 
      Height          =   405
      Left            =   13920
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Retornar"
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
      MICON           =   "frmNavegador.frx":0000
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
Attribute VB_Name = "frmNavegador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRetornar_Click()
TerminateProcess ("iexplore.exe")
Unload Me
End Sub

Private Sub Form_Load()
  frmNavegador.top = 4680
  frmNavegador.left = 90
  frmNavegador.Width = 15180
  frmNavegador.Height = 5790
  
  OpenIE (txt_url)
End Sub

Private Sub OpenIE(url As String)
    IEalt1.Nav (url)
    IEalt1.EmbedIE FrameNavegador.hwnd
    txt_url = ""
End Sub

Private Sub TerminateProcess(app_exe As String)
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & app_exe & "'")
        Process.Terminate
    Next
End Sub
