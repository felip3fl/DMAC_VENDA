VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmAgenda 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Agenda"
   ClientHeight    =   5430
   ClientLeft      =   4095
   ClientTop       =   2295
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbAno 
      BackColor       =   &H00A3A3A3&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1695
      TabIndex        =   13
      Top             =   60
      Width           =   885
   End
   Begin VB.ComboBox cmbMes 
      BackColor       =   &H00A3A3A3&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   75
      TabIndex        =   12
      Top             =   60
      Width           =   1590
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   6165
      TabIndex        =   11
      Top             =   4875
      Width           =   6165
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5175
      TabIndex        =   10
      Top             =   5055
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmAgenda.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox picAvancar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2895
      MouseIcon       =   "frmAgenda.frx":001C
      Picture         =   "frmAgenda.frx":0326
      ScaleHeight     =   375
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "Avança"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox picVoltar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2625
      MouseIcon       =   "frmAgenda.frx":05B5
      Picture         =   "frmAgenda.frx":08BF
      ScaleHeight     =   375
      ScaleWidth      =   240
      TabIndex        =   8
      ToolTipText     =   "Retorna"
      Top             =   30
      Width           =   240
   End
   Begin VB.TextBox txtVendedor 
      Height          =   285
      Left            =   210
      TabIndex        =   6
      Top             =   5085
      Visible         =   0   'False
      Width           =   500
   End
   Begin MSMask.MaskEdBox mskCadData 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   4170
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Project1.chameleonButton cmdCadastrar 
      Height          =   0
      Left            =   5085
      TabIndex        =   2
      Top             =   4185
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   14737632
      MPTR            =   1
      MICON           =   "frmAgenda.frx":0B4D
      PICN            =   "frmAgenda.frx":0B69
      PICH            =   "frmAgenda.frx":71AF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAssunto 
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
      Height          =   315
      Left            =   1185
      MaxLength       =   45
      TabIndex        =   1
      Top             =   4170
      Width           =   5160
   End
   Begin VSFlex7UCtl.VSFlexGrid grdAgenda 
      Height          =   3390
      Left            =   60
      TabIndex        =   3
      Top             =   435
      Width           =   6285
      _cx             =   11086
      _cy             =   5980
      _ConvInfo       =   1
      Appearance      =   2
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   31
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAgenda.frx":D7F5
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   2
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
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1395
      OleObjectBlob   =   "frmAgenda.frx":D88D
      Top             =   4830
   End
   Begin VB.Label lblMensagem 
      BackStyle       =   0  'Transparent
      Caption         =   "Nenhum registro encontrado."
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
      Height          =   240
      Left            =   3450
      TabIndex        =   7
      Top             =   105
      Width           =   2535
   End
   Begin VB.Label lblAssunto 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1230
      TabIndex        =   5
      Top             =   3885
      Width           =   570
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   3885
      Width           =   345
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAgenda As New ADODB.Recordset
Dim SQL As String


Private Sub cmbAno_Click()
  Call PesquisaAgenda
End Sub

Private Sub cmbAno_Scroll()
  Call PesquisaAgenda
End Sub

Private Sub cmbMes_Click()
  Call PesquisaAgenda
End Sub

Private Sub cmbMes_Scroll()
  Call PesquisaAgenda
End Sub

Private Sub cmdCadastrar_Click()
    
    If Trim(mskCadData.Text) = "" Or mskCadData.Text = "__/__/____" Then
        MsgBox "Informe uma data.", vbInformation, Me.Caption
        mskCadData.SetFocus
        Exit Sub
    ElseIf IsDate(mskCadData.Text) = False Then
        MsgBox "Data inválida.", vbCritical, Me.Caption
        mskCadData.SetFocus
        Exit Sub
    End If
    
    If Trim(txtAssunto.Text) = "" Then
        MsgBox "Informe o assunto.", vbCritical, Me.Caption
        txtAssunto.SetFocus
        Exit Sub
    End If
    
On Error GoTo ErroCadastro
    
    sql1 = ""
    sql1 = "Insert Into Agenda(AGE_Data,AGE_Loja,AGE_Vendedor,AGE_Assunto) Values " & _
           "('" & Format(mskCadData.Text, "yyyy/mm/dd") & "','" & Trim(wLoja) & "'," & txtVendedor.Text & ",'" & Trim(txtAssunto.Text) & "')"

    adoCNLoja.BeginTrans
    adoCNLoja.Execute sql1
    adoCNLoja.CommitTrans
    MsgBox "Cadastro incluído com sucesso!", vbInformation, Me.Caption
    Call LimpaCampos
    'txtData.SetFocus
    Exit Sub


ErroCadastro:
    adoCNLoja.RollbackTrans
    MsgBox Err.description & vbLf _
           & "Erro ao incluir cadastro.", vbCritical, Me.Caption
    Exit Sub
    
End Sub

Private Sub cmdGrava_Click()
    
    If Trim(mskCadData.Text) = "" Or mskCadData.Text = "__/__/____" Then
        MsgBox "Informe uma data.", vbInformation, Me.Caption
        mskCadData.SetFocus
        Exit Sub
    ElseIf IsDate(mskCadData.Text) = False Then
        MsgBox "Data inválida.", vbCritical, Me.Caption
        mskCadData.SetFocus
        Exit Sub
    ElseIf Year(mskCadData.Text) > Year(DateAdd("yyyy", 1, Date)) Then
        MsgBox "Data maior que a permitida.", vbCritical, Me.Caption
        mskCadData.SetFocus
        Exit Sub
     ElseIf Year(mskCadData.Text) < Year(DateAdd("yyyy", -1, Date)) Then
        MsgBox "Data maior que a permitida.", vbCritical, Me.Caption
        mskCadData.SetFocus
        Exit Sub
    End If
    
    If Trim(txtAssunto.Text) = "" Then
        MsgBox "Informe o assunto.", vbCritical, Me.Caption
        txtAssunto.SetFocus
        Exit Sub
    End If
    
'On Error GoTo ErroCadastro
    
    sql1 = ""
    sql1 = "Insert Into Agenda(AGE_Data,AGE_Loja,AGE_Vendedor,AGE_Assunto) Values " & _
           "('" & Format(mskCadData.Text, "yyyy/mm/dd") & "','" & Trim(wLoja) & "'," & txtVendedor.Text & ",'" & Trim(txtAssunto.Text) & "')"

    adoCNLoja.BeginTrans
    adoCNLoja.Execute sql1
    adoCNLoja.CommitTrans
    MsgBox "Cadastro incluído com sucesso!", vbInformation, Me.Caption
    Call LimpaCampos
    Call PesquisaAgenda
    Exit Sub


'ErroCadastro:
'    adoCNLoja.RollbackTrans
'    MsgBox Err.description & vbLf _
'           & "Erro ao incluir cadastro.", vbCritical, Me.Caption
'    Exit Sub
    
End Sub




Private Sub cmdRetorna_Click()
    Unload Me
    frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub Form_Load()

'  Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
'  Skin1.ApplySkin Me.hwnd

   grdAgenda.MergeRow(0) = True
   grdAgenda.MergeRow(1) = True
   grdAgenda.MergeCol(0) = True
   grdAgenda.MergeCol(1) = True
   grdAgenda.MergeCol(2) = True
   
   Call AjustaTela(frmAgenda)
   Call LimpaCampos
   Call PreencheCombos
   
End Sub

Private Sub fraAgenda_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grdAgenda_DblClick()
 On Error GoTo erronaUpdate

 If grdAgenda.Row < 2 Then
        Exit Sub
 End If

 
 If MsgBox("Deseja Excluir?", _
           vbYesNo + vbQuestion, "Atenção") = vbYes Then
    adoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    

    SQL = "Delete Agenda Where " _
          & "AGE_Sequencia = " & grdAgenda.TextMatrix(grdAgenda.Row, 2)
         
         adoCNLoja.Execute (SQL)
         Screen.MousePointer = vbNormal
         adoCNLoja.CommitTrans

    Call PesquisaAgenda
      
    Exit Sub
 Else
    Exit Sub
 End If

erronaUpdate:
MsgBox "Erro na Exclusão Item " & Err.description, vbCritical, "Aviso"
'adoCNLoja.RollbackTrans
Screen.MousePointer = vbNormal


End Sub

Private Sub grdAgenda_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub grdAgenda_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub mskCadData_GotFocus()
    mskCadData.SelStart = 0
    mskCadData.SelLength = Len(mskCadData.Text)
End Sub

Private Sub mskCadData_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub mskCadData_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
    frmPedido.txtPesquisar.SetFocus
End If

 
End Sub

Private Sub picAvancar_Click()

     If cmbAno.Text >= Year(DateAdd("yyyy", 1, Date)) And cmbMes.ListIndex = 11 Then
        MsgBox "Ano maior que o permitido."
        Exit Sub
     End If


     If cmbMes.ListIndex = 11 Then
        cmbMes.ListIndex = 0
        cmbAno.ListIndex = cmbAno.ListIndex + 1
     Else
        cmbMes.ListIndex = cmbMes.ListIndex + 1
     End If
     Call PesquisaAgenda
     
End Sub

Private Sub picVoltar_Click()

     If cmbAno.Text <= Year(DateAdd("yyyy", -1, Date)) And cmbMes.ListIndex = 0 Then
        MsgBox "Ano menor que o permitido."
        Exit Sub
     End If
     If cmbMes.ListIndex = 0 Then
        cmbMes.ListIndex = 11
        cmbAno.ListIndex = cmbAno.ListIndex - 1
        
     Else
        cmbMes.ListIndex = cmbMes.ListIndex - 1
     End If
     Call PesquisaAgenda
End Sub



Private Sub txtAssunto_GotFocus()
   txtAssunto.SelStart = 0
   txtAssunto.SelLength = Len(txtAssunto.Text)
End Sub



Private Sub txtData_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
    frmPedido.txtPesquisar.SetFocus
    
 ElseIf KeyAscii = 13 Then
     Call PesquisaAgenda
 End If
 
End Sub

Private Sub LimpaCampos()
    txtAssunto.Text = ""
    mskCadData.Text = "__/__/____"
    'txtData.Text = "__/__/____"
    grdAgenda.Rows = 2
 '   grdAgenda.TextMatrix(0, 0) = "Mês"
'    grdAgenda.TextMatrix(1, 0) = "Dia"
    
    txtVendedor.Text = ""
    txtVendedor.Text = Val(frmPedido.txtVendedor.Text)

End Sub

Function TraduzMesAgenda(ByVal Mes As String) As String
    Select Case Mes
        Case "janeiro", -11: TraduzMesAgenda = "Janeiro"
        Case "fevereiro", -10: TraduzMesAgenda = "Fevereiro"
        Case "março", -9: TraduzMesAgenda = "Março"
        Case "abril", -8: TraduzMesAgenda = "Abril"
        Case "maio", -7: TraduzMesAgenda = "Maio"
        Case "junho", -6: TraduzMesAgenda = "Junho"
        Case "julho", -5: TraduzMesAgenda = "Julho"
        Case "agosto", -4: TraduzMesAgenda = "Agosto"
        Case "setembro", -3: TraduzMesAgenda = "Setembro"
        Case "outubro", -2: TraduzMesAgenda = "Outubro"
        Case "novembro", -1: TraduzMesAgenda = "Novembro"
        Case "dezembro", 0: TraduzMesAgenda = "Dezembro"
    End Select
End Function

Private Sub PesquisaAgenda()

    
    If cmbAno = "" Or cmbMes = "" Then
        Exit Sub
    End If
        
        grdAgenda.Rows = 2
        sql1 = ""
        sql1 = "Select AGE_Data, AGE_Assunto, AGE_Sequencia From Agenda " & _
               "Where  AGE_Vendedor = " & txtVendedor.Text & " And " & _
               "year(AGE_Data) = " & cmbAno.Text & " And Month(AGE_Data) = " & cmbMes.ListIndex + 1 & " And " & _
               "AGE_Loja = '" & Trim(wLoja) & "' Order By AGE_Sequencia"
               
        rsAgenda.CursorLocation = adUseClient
        rsAgenda.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If rsAgenda.EOF = False Then
            Do While Not rsAgenda.EOF
                grdAgenda.AddItem Day(rsAgenda("Age_data")) & Chr(9) & _
                                  Trim(rsAgenda("AGE_Assunto")) & Chr(9) & (rsAgenda("AGE_sequencia"))
                rsAgenda.MoveNext
            Loop
            
            lblMensagem.Visible = False
'            grdAgenda.SetFocus

        Else
            lblMensagem.Visible = True
            grdAgenda.Rows = 2

        End If
        rsAgenda.Close



End Sub

Private Sub PreencheCombos()

  cmbMes.AddItem "01 - Janeiro"
  cmbMes.AddItem "02 - Fevereiro"
  cmbMes.AddItem "03 - Março"
  cmbMes.AddItem "04 - Abril"
  cmbMes.AddItem "05 - Maio"
  cmbMes.AddItem "06 - Junho"
  cmbMes.AddItem "07 - Julho"
  cmbMes.AddItem "08 - Agosto"
  cmbMes.AddItem "09 - Setembro"
  cmbMes.AddItem "10 - Outubro"
  cmbMes.AddItem "11 - Novembro"
  cmbMes.AddItem "12 - Dezembro"
  cmbMes.ListIndex = Month(Date) - 1
  
  cmbAno.AddItem Year(DateAdd("yyyy", -1, Date))
  cmbAno.AddItem Year(DateAdd("yyyy", 0, Date))
  cmbAno.AddItem Year(DateAdd("yyyy", 1, Date))
  cmbAno.ListIndex = 1
   
End Sub


Private Sub txtAssunto_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtAssunto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me
     ElseIf KeyAscii = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub
