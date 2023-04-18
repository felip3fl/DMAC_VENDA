VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmEnviaEmail 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Envia Cotação de Venda para Cliente"
   ClientHeight    =   5505
   ClientLeft      =   540
   ClientTop       =   1530
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   15
      ScaleHeight     =   45
      ScaleWidth      =   6360
      TabIndex        =   9
      Top             =   4845
      Width           =   6360
   End
   Begin VB.TextBox txtDestino 
      BackColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   510
      TabIndex        =   3
      Top             =   45
      Width           =   5850
   End
   Begin VB.TextBox txtObs 
      BackColor       =   &H00C0C0C0&
      Height          =   3330
      Left            =   495
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1020
      Width           =   5850
   End
   Begin VB.Frame fraGeral 
      BackColor       =   &H00404040&
      Height          =   525
      Left            =   525
      TabIndex        =   0
      Top             =   405
      Width           =   5820
      Begin VB.ComboBox cmbTamanho 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   135
         Width           =   915
      End
      Begin VB.ComboBox cmbFonte 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "FrmEnviaEmail.frx":0000
         Left            =   45
         List            =   "FrmEnviaEmail.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   135
         Width           =   2775
      End
      Begin Threed.SSRibbon srbSublinhado 
         Height          =   345
         Left            =   5340
         TabIndex        =   6
         Top             =   135
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   609
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "FrmEnviaEmail.frx":0004
      End
      Begin Threed.SSRibbon srbItalico 
         Height          =   345
         Left            =   4800
         TabIndex        =   7
         Top             =   135
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   609
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "FrmEnviaEmail.frx":0672
      End
      Begin Threed.SSRibbon srbNegrito 
         Height          =   345
         Left            =   4245
         TabIndex        =   8
         Top             =   135
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   609
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "FrmEnviaEmail.frx":0BCC
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   4140
         X2              =   4140
         Y1              =   135
         Y2              =   495
      End
      Begin VB.Line Line2 
         X1              =   5805
         X2              =   5805
         Y1              =   135
         Y2              =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2085
      OleObjectBlob   =   "FrmEnviaEmail.frx":1236
      Top             =   5070
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1095
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   105
      Top             =   4905
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin Project1.chameleonButton cmdEnviar 
      Height          =   405
      Left            =   4080
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Enviar"
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
      MICON           =   "FrmEnviaEmail.frx":146A
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
      Left            =   5235
      TabIndex        =   11
      Top             =   5040
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
      MICON           =   "FrmEnviaEmail.frx":1486
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPara 
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   30
      TabIndex        =   1
      Top             =   150
      Width           =   525
   End
End
Attribute VB_Name = "frmEnviaEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFonte_Click()
    
    txtObs.FontName = cmbFonte.Text
    
End Sub

Private Sub cmbTamanho_Click()

    txtObs.FontSize = cmbTamanho.Text

End Sub

Private Sub cmdEnviar_Click()
    Dim Cotacao As String
    
    Cotacao = "Cot" & frmPedido.txtPedido.Text & ".html"
    
    MAPISession1.Action = 1
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.Compose
    MAPIMessages1.RecipAddress = txtDestino.Text
    MAPIMessages1.AddressResolveUI = True
'    MAPIMessages1.ResolveName
    MAPIMessages1.AttachmentName = "Cotacao.html"
    MAPIMessages1.AttachmentPathName = GLB_Cotacao & Cotacao
    MAPIMessages1.MsgSubject = "Cotação de Venda"
    MAPIMessages1.MsgNoteText = txtObs.Text
    MAPIMessages1.Send False
    MAPISession1.Action = 2
    
    txtDestino.Text = ""
    txtObs.Text = ""
    Unload Me

End Sub

Private Sub CmdRetornar_Click()
'    frmPedido.picQuadroGeral.Width = 12200
    Unload Me
End Sub



Private Sub cmdRetorna_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  
  Call AjustaTela(frmEnviaEmail)
  
  'Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
'  Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd

    For i = 0 To Screen.FontCount
        If Screen.Fonts(i) <> "" Then
            cmbFonte.AddItem Screen.Fonts(i)
        End If
    Next i
    cmbFonte.ListIndex = 42
    
    For i = 12 To 34 Step 2
        cmbTamanho.AddItem i
    Next i
    cmbTamanho.ListIndex = 0
    
    If AchaLojaControle = "85" Then
        txtObs.Text = "Segue em anexo a cotação de venda AFG"
    ElseIf AchaLojaControle = "314" Or AchaLojaControle = "364" Then
        txtObs.Text = "Segue em anexo a cotação de venda DM Motores"
    Else
        txtObs.Text = "Segue em anexo a cotação de venda De Meo"
    End If
    txtDestino.Text = LCase(PegaEmailCliente(frmPedido.txtPedido.Text))
    txtDestino.SelStart = 0
    txtDestino.SelLength = Len(txtDestino.Text)
    
End Sub

Private Sub rctObs_Change()
    
    If srbNegrito.Value = True Then
        
    End If
    
End Sub



Private Sub srbItalico_Click(Value As Integer)

    If srbItalico.Value = True Then
        txtObs.FontItalic = True
    Else
        txtObs.FontItalic = False
    End If

End Sub

Private Sub srbNegrito_Click(Value As Integer)
    
    If srbNegrito.Value = True Then
        txtObs.FontBold = True
    Else
        txtObs.FontBold = False
    End If
    
End Sub

Private Sub srbSublinhado_Click(Value As Integer)
    
    If srbSublinhado.Value = True Then
        txtObs.FontUnderline = True
    Else
        txtObs.FontUnderline = False
    End If
    
End Sub

Private Sub txtDestino_GotFocus()
'    cmdEnviar.Enabled = False
End Sub

Private Sub txtDestino_LostFocus()
    If txtDestino.Text <> "" Then
        cmdEnviar.Enabled = True
    End If
End Sub

Private Sub txtObs_GotFocus()

    txtObs.FontName = cmbFonte.Text
    txtObs.FontSize = cmbTamanho.Text
    If srbNegrito.Value = True Then
        txtObs.FontBold = True
    Else
        txtObs.FontBold = False
    End If
    If srbItalico.Value = True Then
        txtObs.FontItalic = True
    Else
        txtObs.FontItalic = False
    End If
    If srbSublinhado.Value = True Then
        txtObs.FontUnderline = True
    Else
        txtObs.FontUnderline = False
    End If
    
End Sub

Function PegaEmailCliente(ByVal Pedido As Double) As String
    sql = ""
    sql = "Select  COV_ValorComplemento from ComplementoVenda" _
          & " where COV_NumeroPedido =" & Pedido & " and COV_SequenciaComplemento = 6"
    rdoCliente.CursorLocation = adUseClient
    rdoCliente.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
       
If Not rdoCliente.EOF Then
   wValorComplemento = rdoCliente("COV_ValorComplemento")
       
        sql1 = ""
   
        sql1 = "Select CE_Email  from FIN_ClienteWhere CE_CodigoCliente= " & wValorComplemento & ""
            rsCliente.CursorLocation = adUseClient
            rsCliente.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
                         
   
    If Not rsCliente.EOF Then
        PegaEmailCliente = IIf(IsNull(rsCliente("CE_Email")), "", rsCliente("CE_Email"))
        'PegaEmailCliente = rsCliente("CE_Email")
        rsCliente.Close
    Else
        PegaEmailCliente = ""
    End If
End If
   rdoCliente.Close

End Function
