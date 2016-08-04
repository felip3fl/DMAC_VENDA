VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmCotacao 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   ClientHeight    =   7095
   ClientLeft      =   4800
   ClientTop       =   2070
   ClientWidth     =   13650
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   7095
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerImpressao 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   5055
   End
   Begin SHDocVwCtl.WebBrowser WebNavegador 
      Height          =   4740
      Left            =   11475
      TabIndex        =   3
      Top             =   915
      Width           =   6360
      ExtentX         =   11218
      ExtentY         =   8361
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   1
      Top             =   4875
      Width           =   6165
   End
   Begin VB.TextBox txtPedido 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5970
      Visible         =   0   'False
      Width           =   450
   End
   Begin Project1.chameleonButton cmdOk 
      Height          =   405
      Left            =   5280
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
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
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "FrmCotacao.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTipoTransporte 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Transporte"
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
      Left            =   180
      TabIndex        =   4
      Top             =   4320
      Width           =   2325
   End
End
Attribute VB_Name = "FrmCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()
    
    'timerImpressao.Enabled = True
   
    'Unload Me
    'timerImpressao_Timer
    
'    If OptEnviaCotacao.Value = True Then
'        CriaCotacaoHtml txtpedido.Text
'        frmEnviaEmail.Show 1
'    ElseIf OptImprimeCotacao.Value = True Then
'        'ImprimirCotacaoBola frmPedido.txtPedido.Text
        
'    ElseIf OptVisualizar.Value = True Then
'        CriaCotacaoHtml frmPedido.txtpedido.Text
'        ShellExecute &O0, "Open", GLB_Cotacao & "Cot" & txtpedido.Text & ".html", &O0, &O0, 1
'    ElseIf OptFechaCotacao.Value = True Then
'    End If
'    OptFechaCotacao.Value = True

    FrmCotacao.WebNavegador.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
    
    Call FechaCotacao
    
    Screen.MousePointer = 0
    
End Sub

Private Sub CmdCancela_Click()
    Unload Me
    frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 11
    dados frmPedido.txtpedido.Text
    Me.Visible = False
    txtpedido.Text = frmPedido.txtpedido.Text
    timerImpressao.Enabled = True
End Sub

Private Sub Form_Load()
    Call AjustaTela(FrmCotacao)
End Sub
Private Sub SkinLabel1_DragDrop(Source As Control, X As Single, Y As Single)
    
End Sub

Private Sub SSOption1_Click(Value As Integer)
    
End Sub

Private Sub Label1_Click()
    
End Sub

Private Sub FechaCotacao()
    
    Dim SQL As String
    
    SQL = "Select sum(vltotitem) as vlrmercadoria, sum(vltotitem - desconto) as totalnota " & _
    "from nfitens where numeroped = " & txtpedido.Text
    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    SQL = ""
    SQL = "Update NFCapa set TipoNota = 'PD', vlrmercadoria = " & ConverteVirgula(Format(rsComplementoVenda("vlrmercadoria"), "##0.00")) & _
    ", TotalNota = " & ConverteVirgula(Format(rsComplementoVenda("Totalnota"), "##0.00")) & _
    " Where NumeroPed = " & txtpedido.Text
    adoCNLoja.Execute SQL
    
    SQL = ""
    SQL = "Update NFItens Set TipoNota = 'PD' Where NumeroPed = " & txtpedido.Text
    adoCNLoja.Execute SQL
    
    SQL = ""
    
    Call LimpaTR
    rsComplementoVenda.Close
    Unload Me
    Call LimpaForm
    
End Sub

Private Sub timerImpressao_Timer()
    CmdOK_Click
    timerImpressao.Enabled = False
    Unload Me
End Sub
