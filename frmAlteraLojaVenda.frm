VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAlteraLojaVenda 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Altera Loja Venda (Venda a Distancia)"
   ClientHeight    =   5670
   ClientLeft      =   3540
   ClientTop       =   4380
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmAlterarNF 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   150
      TabIndex        =   20
      Top             =   2790
      Width           =   6165
      Begin VB.ComboBox cmbLoja 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   150
         TabIndex        =   25
         Text            =   "271"
         Top             =   1185
         Width           =   1515
      End
      Begin VB.TextBox txtChaveAcesso 
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
         Left            =   150
         MaxLength       =   44
         TabIndex        =   5
         ToolTipText     =   " "
         Top             =   420
         Width           =   5835
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja venda distancia"
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
         Left            =   150
         TabIndex        =   26
         Top             =   930
         Width           =   2850
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Chave de acesso da NF (Somente Números)"
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
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   150
         Width           =   6165
      End
   End
   Begin VB.Frame frmInfoNF 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   150
      TabIndex        =   13
      Top             =   1500
      Width           =   6165
      Begin VB.Label lblLojaVenda 
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
         Left            =   4290
         TabIndex        =   24
         Top             =   750
         Width           =   435
      End
      Begin VB.Label lblLojaOrigem 
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
         Left            =   1365
         TabIndex        =   23
         Top             =   750
         Width           =   435
      End
      Begin VB.Label lblVendedor 
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
         Left            =   4290
         TabIndex        =   22
         Top             =   400
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loja Venda:"
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
         Left            =   3000
         TabIndex        =   19
         Top             =   750
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loja Origem:"
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
         Left            =   0
         TabIndex        =   18
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label lblTotalVendas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   3000
         TabIndex        =   17
         Top             =   400
         Width           =   1095
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição da Nota Fiscal"
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
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2640
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Nota:"
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
         Left            =   0
         TabIndex        =   15
         Top             =   400
         Width           =   1170
      End
      Begin VB.Label lblTotalNota 
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
         Left            =   1365
         TabIndex        =   14
         Top             =   400
         Width           =   435
      End
   End
   Begin VB.Frame frmNF 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   885
      Left            =   795
      TabIndex        =   8
      Top             =   600
      Width           =   5115
      Begin MSMask.MaskEdBox mskDataEmissao 
         Height          =   360
         Left            =   2595
         TabIndex        =   3
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   3885
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   1305
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   3885
         TabIndex        =   12
         Top             =   0
         Width           =   2490
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Emissão"
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
         Left            =   2595
         TabIndex        =   11
         Top             =   0
         Width           =   2490
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
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
         Left            =   1275
         TabIndex        =   10
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
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
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   0
      Top             =   4725
      Width           =   6165
   End
   Begin Project1.chameleonButton cmdGravar 
      Height          =   405
      Left            =   5190
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
      _ExtentX        =   3493
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Gravar"
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
      MICON           =   "frmAlteraLojaVenda.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPagamento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alterar Loja Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   6165
   End
End
Attribute VB_Name = "frmAlteraLojaVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim SQL As String
Dim rdoLojas As New ADODB.Recordset
Dim rdoNotas As New ADODB.Recordset
Dim rdoGravar As New ADODB.Recordset
Dim wTotalNota As Double
Dim CHAVENF As String


Private Sub cmbLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

  If rdoCNMatriz.State = 1 Then
    rdoCNMatriz.Close
  End If
  
  ConectaODBCMatriz
  If GLB_ConectouOK = False Then
     MsgBox "Erro ao conectar-se ao Banco de Dados da Matriz", vbCritical, "Atenção"
     Unload Me
  End If
  

    Call AjustaTela(frmAlteraLojaVenda)

    frmNF.BackColor = Me.BackColor
    frmInfoNF.BackColor = Me.BackColor
    'rmAlterarNF.BackColor = Me.BackColor
    
    
    lblTotalNota.Caption = ""
    lblLojaOrigem.Caption = ""
    lblVendedor.Caption = ""
    lblLojaVenda.Caption = ""
    cmdGravar.Enabled = False
    
    frmAlterarNF.Enabled = False
    
    SQL = "Select Lo_loja from loja where lo_Regiao < 900 and lo_loja not in ('CD5') GROUP BY lo_loja"
 
        rdoLojas.CursorLocation = adUseClient
        rdoLojas.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If Not rdoLojas.EOF Then
            cmbLoja.Clear
            Do While Not rdoLojas.EOF
                cmbLoja.AddItem Trim(rdoLojas("lo_loja"))
                rdoLojas.MoveNext
            Loop
            
            Screen.MousePointer = 0
        End If

rdoLojas.Close

cmbLoja.ListIndex = 0

End Sub
Private Sub CarregaNota()
    

'SQL = "select  * " _
'    & "From NFCAPA " _
'    & "where NF = '" & txtNumero.Text & "' and SERIE = '" & txtSerie.Text & "' and DATAEMI = '" & Format(mskDataEmissao.Text, "yyyy/mm/dd") & "' and Cliente = '" & RTrim(Trim(txtCliente.Text)) & "' "

SQL = ""
SQL = "select VC_ChaveNFE,VC_TotalNota, VC_LojaOrigem, VC_VendedorLojaVenda ,VC_LojaVenda " _
    & "From CapaNFVenda " _
    & "where VC_NotaFiscal = '" & txtNumero.Text & "' and VC_Serie = '" & "NE" & "' " _
    & "and VC_DataEmissao = '" & Format(mskDataEmissao.Text, "yyyy/mm/dd") & "' " _
    & "and VC_Cliente = '" & RTrim(Trim(txtCliente.Text)) & "'" & "and VC_CodigoVendedor = '" & Mid(frmPedido.txtVendedor.Text, 1, 3) & "' " _
    & "and VC_LojaOrigem = '" & RTrim(wLoja) & "'"


        rdoNotas.CursorLocation = adUseClient
        
        rdoNotas.Open SQL, rdoCNMatriz, adOpenForwardOnly, adLockPessimistic
        
        
           If Not rdoNotas.EOF Then

                lblTotalNota.Caption = Format(rdoNotas("VC_TotalNota"), "##0.00")
                lblLojaOrigem.Caption = rdoNotas("VC_LojaOrigem")
                lblVendedor.Caption = rdoNotas("VC_VendedorLojaVenda")
                lblLojaVenda.Caption = rdoNotas("VC_LojaVenda")
                CHAVENF = rdoNotas("VC_ChaveNFE")
                
                frmAlterarNF.Enabled = True
                frmNF.Enabled = False
            
           Else
                MsgBox "Não existe informações sobre esta nota, ou série errada", vbInformation
                LimparCampos
                frmAlterarNF.Enabled = False
                frmNF.Enabled = True
                txtCliente.SetFocus
           End If
           
        rdoNotas.Close
End Sub

Private Sub mskDataEmissao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

    If Len(mskDataEmissao.Text) = 2 Then
            mskDataEmissao.Text = mskDataEmissao.Text & "/"
            mskDataEmissao.SelStart = 3
        ElseIf Len(mskDataEmissao.Text) = 5 Then
            mskDataEmissao.Text = mskDataEmissao.Text & "/"
            mskDataEmissao.SelStart = 6
        End If
        
End Sub

Private Sub txtChaveAcesso_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

If KeyAscii > 64 Then  'Não permite letras
    KeyAscii = 0
End If

    'Não permite caracteres especiais
  If (KeyAscii < 65 Or KeyAscii > 90) And _
       (KeyAscii < 97 Or KeyAscii > 122) And _
       (KeyAscii < 48 Or KeyAscii > 57) And _
       (KeyAscii <> 32) And _
       (KeyAscii <> 13) And _
       (KeyAscii > 9) Then
        KeyAscii = 0
  End If
  
  If KeyAscii = 13 Then
  
    If Len(txtChaveAcesso.Text) < 44 Then
            MsgBox "Chave de acesso Incorreta, A chave deve conter 44 números ", vbCritical, "Atenção"
            txtChaveAcesso.SetFocus
    Else
    
        If CHAVENF = txtChaveAcesso.Text Then
            cmdGravar.Enabled = True
        Else
            MsgBox "Chave de acesso incorreta", vbInformation, "Atenção"
        End If
    
    End If

    
 End If
      
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtSerie_LostFocus()
    txtSerie.Text = UCase(txtSerie.Text)
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 13 Then
        CarregaNota
        cmdGravar.Enabled = False
       
    End If
         
End Sub

Private Sub cmdGravar_Click()
'ricardo
         'Update loja
        SQL = ""
        SQL = "Update NFCAPA set LojaVenda = '" & cmbLoja.Text & "' where NF = '" & txtNumero.Text & "'" _
          & "and SERIE = '" & txtSerie.Text & "' and DATAEMI = '" & Format(mskDataEmissao.Text, "yyyy/mm/dd") & "'" _
          & "and Cliente = '" & RTrim(Trim(txtCliente.Text)) & "' and LojaOrigem = '" & lblLojaOrigem.Caption & "'"
    
        adoCNLoja.Execute (SQL)
        rdoCNMatriz.Execute (SQL)
        
        'Update Matriz
        SQL = "Update CapaNFVenda set VC_LojaVenda = '" & cmbLoja.Text & "'" _
            & "where VC_TipoNota = 'V' and vc_notafiscal = '" & txtNumero.Text & "' and vc_Serie = '" & txtSerie.Text & "' and vc_DataEmissao = '" & Format(mskDataEmissao.Text, "yyyy/mm/dd") & "'" _
            & "and VC_Cliente = '" & RTrim(Trim(txtCliente.Text)) & "' and vc_LojaOrigem = '" & lblLojaOrigem.Caption & "'"
        
        rdoCNMatriz.Execute (SQL)
        
        MsgBox "Nota fiscal alterada com sucesso", vbInformation, "Atenção"
        LimparCampos
        Unload Me
        
End Sub

Private Sub LimparCampos()
    txtNumero.Text = ""
    txtSerie.Text = ""
    mskDataEmissao.Text = ""
    txtCliente.Text = ""
    txtChaveAcesso.Text = ""
    lblTotalNota.Caption = ""
    lblLojaOrigem.Caption = ""
    lblVendedor.Caption = ""
    lblLojaVenda.Caption = ""
    txtChaveAcesso.Enabled = False
    cmdGravar.Enabled = False
    cmbLoja.ListIndex = 0
    
End Sub

Private Sub lblTotalVendas_Click()
400
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtPesquisaCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub


