VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmTransportadora 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transportadora"
   ClientHeight    =   5595
   ClientLeft      =   6780
   ClientTop       =   2700
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   ScaleHeight     =   5595
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   12
      Top             =   5355
      Width           =   6165
   End
   Begin VB.Frame fraPagamento 
      BackColor       =   &H00505050&
      ForeColor       =   &H00FFFFFF&
      Height          =   4185
      Left            =   165
      TabIndex        =   8
      Top             =   570
      Width           =   6195
      Begin VB.TextBox txtTransportadora 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   1365
         TabIndex        =   0
         Text            =   "DE MEO TRANSPORTADORA LTDA"
         ToolTipText     =   " "
         Top             =   600
         Width           =   4710
      End
      Begin VB.TextBox txtPlaca 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   4305
         TabIndex        =   7
         Text            =   "YYY-1234"
         ToolTipText     =   " "
         Top             =   3300
         Width           =   1770
      End
      Begin VB.TextBox txtInscricaoEstadual 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "SÃO PAULO"
         ToolTipText     =   " "
         Top             =   1500
         Width           =   3435
      End
      Begin VB.ComboBox cmbEstado 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   6
         Text            =   "SP"
         Top             =   3300
         Width           =   930
      End
      Begin VB.TextBox txtMunicipio 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   150
         TabIndex        =   5
         Text            =   "SÃO PAULO"
         ToolTipText     =   " "
         Top             =   3300
         Width           =   2955
      End
      Begin VB.TextBox txtEndereco 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   150
         TabIndex        =   4
         Text            =   "FLORENCIO DE ABREU 271"
         ToolTipText     =   " "
         Top             =   2400
         Width           =   5925
      End
      Begin VB.TextBox txtNumeroTransportadora 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   150
         TabIndex        =   20
         Text            =   "Codigo"
         ToolTipText     =   " "
         Top             =   600
         Width           =   1080
      End
      Begin VB.TextBox txtCNPJ 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   150
         MaxLength       =   14
         TabIndex        =   2
         Text            =   "60872124001020"
         Top             =   1500
         Width           =   2370
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdMunicipio 
         Height          =   480
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Visible         =   0   'False
         Width           =   3135
         _cx             =   5530
         _cy             =   847
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
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransportadora.frx":0000
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
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   6075
         X2              =   6075
         Y1              =   315
         Y2              =   3750
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Transportadora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   3210
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   4305
         TabIndex        =   18
         Top             =   2950
         Width           =   2760
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   2640
         TabIndex        =   17
         Top             =   1150
         Width           =   2760
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   3240
         TabIndex        =   16
         Top             =   2950
         Width           =   2760
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Municipio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   15
         Top             =   2950
         Width           =   2760
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   14
         Top             =   2050
         Width           =   2760
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   13
         Top             =   1150
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         TabIndex        =   11
         Top             =   255
         Width           =   810
      End
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5280
      TabIndex        =   9
      Top             =   4800
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
      MICON           =   "frmTransportadora.frx":003D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdRetornar 
      Height          =   405
      Left            =   4200
      TabIndex        =   10
      Top             =   4800
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
      MICON           =   "frmTransportadora.frx":0059
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
      Caption         =   "Transportadora"
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
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   6300
   End
End
Attribute VB_Name = "frmTransportadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGravaTransportadora As New ADODB.Recordset
Dim adoCliente As New ADODB.Recordset
Dim rsBuscaNumeroTransportadora As New ADODB.Recordset
Dim SQL As String

Dim wLimpar As Boolean
Dim wPreencheInicio As Boolean
Dim ln As Integer

Private Sub cmdRetornar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call AjustaTela(Me)
    LimparCampos
    
    SQL = ""
    SQL = "select CTS_NumeroTransportadora from ControleSistema "
    
            rsBuscaNumeroTransportadora.CursorLocation = adUseClient
            rsBuscaNumeroTransportadora.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            txtNumeroTransportadora.Text = rsBuscaNumeroTransportadora("CTS_NumeroTransportadora")
            
            If Not rsBuscaNumeroTransportadora.EOF Then
                adoCNLoja.BeginTrans
                Screen.MousePointer = vbHourglass
                    SQL = ""
                    SQL = "Update ControleSistema set CTS_NumeroTransportadora=(CTS_NumeroTransportadora + 1)"
                          adoCNLoja.Execute SQL
                          Screen.MousePointer = vbNormal
                          adoCNLoja.CommitTrans
                          
                rsBuscaNumeroTransportadora.Close
            End If
            
End Sub
Private Sub LimparCampos()

    txtNumeroTransportadora.Text = ""
    txtTransportadora.Text = ""
    txtCNPJ.Text = ""
    txtInscricaoEstadual.Text = ""
    txtEndereco.Text = ""
    txtMunicipio.Text = ""
    cmbEstado.Text = ""
    txtPlaca.Text = ""
    
End Sub
Function CamposVazio()
    
    'ricardo
    If txtTransportadora.Text = "" Then
        MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
        txtTransportadora.SetFocus
        Exit Function
    End If
    
     If txtCNPJ.Text = "" Then
        MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
        txtCNPJ.SetFocus
        Exit Function
     End If
    
       If txtInscricaoEstadual.Text = "" Then
        MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
        txtInscricaoEstadual.SetFocus
        Exit Function
      End If
       
         If txtEndereco.Text = "" Then
          MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
          txtEndereco.SetFocus
          Exit Function
      End If
    
             If txtMunicipio.Text = "" Then
              MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
              txtMunicipio.SetFocus
              Exit Function
           End If
             
                If cmbEstado.Text = "" Then
                 MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
                 cmbEstado.SetFocus
                 Exit Function
                End If
                
                   If txtPlaca.Text = "" Then
                    MsgBox "Este campo não pode estar vazio", vbInformation, "ATENÇÃO"
                    txtPlaca.SetFocus
                    Exit Function
                   End If
   
End Function

Private Sub cmdGrava_Click()
   
   CamposVazio
   
  SQL = ""
  SQL = "Insert Into Transportadora (Tra_CodigoTransp, Tra_NomeTransportadora , Tra_Placa , Tra_UF , " & _
  " Tra_CNPJ , Tra_IE , Tra_Endereco , Tra_Municipio ) " & _
  "Values ('" & txtNumeroTransportadora.Text & "','" & txtTransportadora.Text & "', '" & txtPlaca.Text & "', " & _
           " '" & cmbEstado.Text & "', '" & txtCNPJ.Text & "', '" & txtInscricaoEstadual.Text & "', " & _
           " '" & txtEndereco.Text & "', '" & txtMunicipio.Text & "')"
           
 adoCNLoja.Execute (SQL)
 MsgBox "Transportadora gravada com sucesso !", vbInformation, "Obrigado"
 
    SQL = ""
    SQL = "Update Transportadora set Tra_CodigoTrans=(Tra_CodigoTransp + 1)"
     adoCNLoja.Execute SQL
 
adoCNLoja.Close
     LimparCampos
     Unload Me
     
End Sub

Private Sub txtMunicipio_Change()

        If txtMunicipio.Text = "'" Then
        MsgBox "Este campo não permite caracteres especiais!", vbCritical, "ATENÇÃO"
        txtMunicipio.Text = ""
        txtMunicipio.SetFocus
        Exit Sub
    End If
    
    grdMunicipio.ZOrder
    If wPreencheInicio = False Then
       grdMunicipio.Visible = True
       PreencheGridMunicipioPesquisa
    End If
    If Trim(txtMunicipio.Text) = "" Then
       grdMunicipio.Visible = False
    End If
    
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

ln = 0

        If Len(txtMunicipio.Text) > 0 Then
           SQL = "SP_FIN_Ler_Codigo_Municipio_Por_Parametro '" & txtMunicipio.Text & "'"
           
            adoCliente.CursorLocation = adUseClient
            adoCliente.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
            
                
         Else
        
         Exit Sub
           
         End If

         Do While Not adoCliente.EOF
           
                With grdMunicipio
                     .AddItem Trim(adoCliente("Mun_nome")) & Chr(9) & Trim(adoCliente("Mun_UF"))
                     .IsSubtotal(.Rows - 1) = True
                     .RowOutlineLevel(.Rows - 1) = 3
                     .Cell(flexcpFontBold, .Rows - 1, 0) = False
                     .Redraw = flexRDBuffered
                End With

            adoCliente.MoveNext
            ln = ln + 1
         Loop
     
            ln = ln - 1
            Do While ln >= 0
                grdMunicipio.IsCollapsed(ln) = flexOutlineCollapsed
                ln = ln - 1
            Loop
            adoCliente.Close
End Sub

Private Sub grdMunicipio_RowColChange()
   On Error GoTo SaidaRotina

    txtMunicipio.Text = UCase(grdMunicipio.TextMatrix(grdMunicipio.Row, 0))
    cmbEstado.Text = UCase(grdMunicipio.TextMatrix(grdMunicipio.Row, 1))
    
SaidaRotina:

    Exit Sub
    
End Sub

Private Sub txtNumeroTransportadora_LostFocus()
    txtNumeroTransportadora.Text = UCase(txtNumeroTransportadora.Text)
End Sub
Private Sub txtTransportadora_LostFocus()
    txtTransportadora.Text = UCase(txtTransportadora.Text)
End Sub
Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
    If KeyAscii > 64 Then  'Não permite letras
    KeyAscii = 0
  End If
End Sub
Private Sub txtInscricaoEstadual_LostFocus()
    txtInscricaoEstadual.Text = UCase(txtInscricaoEstadual.Text)
End Sub
Private Sub txtEndereco_LostFocus()
    txtEndereco.Text = UCase(txtEndereco.Text)
End Sub
Private Sub txtMunicipio_LostFocus()
    txtMunicipio.Text = UCase(txtMunicipio.Text)
End Sub
Private Sub txtPlaca_LostFocus()
    txtPlaca.Text = UCase(txtPlaca.Text)
End Sub








