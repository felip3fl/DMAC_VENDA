VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   5895
   ClientLeft      =   2520
   ClientTop       =   3000
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameNavegador 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   15015
      Begin Project1.IEalt IEalt1 
         Left            =   120
         Top             =   120
         _ExtentX        =   26061
         _ExtentY        =   8705
      End
   End
   Begin Project1.chameleonButton Retorna 
      Height          =   405
      Left            =   13980
      TabIndex        =   0
      Top             =   5340
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
      MICON           =   "frmMapa.frx":0000
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
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
IEalt1.Nav ("https://www.google.com.br/maps/search/" & frmCliente.txtEndereco.Text)
IEalt1.EmbedIE FrameNavegador.hwnd
End Sub

Private Sub Form_Load()
  frmMapa.top = 4680
  frmMapa.left = 90
  frmMapa.Width = 15180
  frmMapa.Height = 5790
End Sub

Private Sub Retorna_Click()
Unload Me
End Sub
