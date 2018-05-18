VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmLembreMe 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Lembre-me quando chegar"
   ClientHeight    =   5505
   ClientLeft      =   7770
   ClientTop       =   1740
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5505
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   60
      ScaleHeight     =   45
      ScaleWidth      =   6360
      TabIndex        =   6
      Top             =   4830
      Width           =   6360
   End
   Begin VB.TextBox txtObservacao 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3525
      Width           =   6300
   End
   Begin VB.TextBox txtReferencia 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   75
      TabIndex        =   2
      Top             =   495
      Width           =   1320
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   570
      TabIndex        =   0
      Top             =   4710
      Visible         =   0   'False
      Width           =   300
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1455
      OleObjectBlob   =   "frmLembreMe.frx":0000
      Top             =   4665
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItensLembre 
      Height          =   2265
      Left            =   75
      TabIndex        =   3
      Top             =   900
      Width           =   6330
      _cx             =   11165
      _cy             =   3995
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLembreMe.frx":0234
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5325
      TabIndex        =   7
      Top             =   5040
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
      MICON           =   "frmLembreMe.frx":0288
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblObservacao 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Observação"
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
      Left            =   90
      TabIndex        =   4
      Top             =   3255
      Width           =   1125
   End
   Begin VB.Label lblReferencia 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
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
      Left            =   105
      TabIndex        =   1
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmLembreMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim wcont As Integer
Dim wSequencia As Integer
Dim rsLembre As ADODB.Recordset



Private Sub cmdGrava_Click()

    If Len(txtObservacao.Text) > 135 Then
      MsgBox "Abreviar a observação"
      Exit Sub
    ElseIf txtObservacao.Text = "" Then
      MsgBox "Informar a observação"
      Exit Sub
    End If
    
    wcont = 0
    
    
    For Index = 1 To grdItensLembre.Rows - 1
        SQL = ""
        SQL = "Insert into LembreMe (LEM_Loja,LEM_Vendedor,LEM_Referencia,LEM_Data,LEM_Observacao,LEM_Situacao)" & _
              " values ('" & Trim(wLoja) & "','" & Mid(frmPedido.txtVendedor.Text, 1, 3) & _
              "','" & Trim(grdItensLembre.TextMatrix(Index, 0)) & "', '" & Format(Date, "yyyy/mm/dd") & "' ,'" & _
              txtObservacao.Text & "','E')"
        adoCNLoja.Execute SQL
        wcont = wcont + 1
     Next Index
     
    If wcont = 0 Then
       MsgBox "É necessário informar no minimo uma referência. "
       txtReferencia.SetFocus
       Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdRetorna_Click()
 Unload Me
 frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub Form_Load()
    Call AjustaTela(frmLembreMe)
 '  Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 '  Skin1.ApplySkin Me.hwnd
  
End Sub


Private Sub grdItensLembre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub grdItensLembre_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub

Private Sub txtObservacao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtObservacao_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub

Private Sub txtReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If txtReferencia <> "" Then
        SQL = ""
        SQL = "select pr_referencia,pr_Descricao from produtoLoja where pr_referencia = '" & txtReferencia.Text & "'"
        rsLembreMe.CursorLocation = adUseClient
        rsLembreMe.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        
        If Not rsLembreMe.EOF Then
             
             For Index = 1 To grdItensLembre.Rows - 1
                  If Trim(grdItensLembre.TextMatrix(Index, 0)) = txtReferencia.Text Then
                     MsgBox "Refêrencia já informada"
                     txtReferencia.Text = ""
                     txtReferencia.SetFocus
                     rsLembreMe.Close
                   Exit Sub
                   End If
             Next Index

             grdItensLembre.AddItem rsLembreMe("PR_referencia") & Chr(9) _
             & rsLembreMe("pr_Descricao")
             
        Else
             MsgBox "Referência inválida."
        End If
      Else
        MsgBox "Informe a Referência"
      End If
      rsLembreMe.Close
      txtReferencia.SetFocus
      txtReferencia.Text = ""
   End If
 
If KeyAscii = 27 Then
    Unload Me
   Exit Sub
End If
End Sub
