VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmEmarketing 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro E-Marketing"
   ClientHeight    =   5475
   ClientLeft      =   3735
   ClientTop       =   2880
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   135
      ScaleHeight     =   45
      ScaleWidth      =   6360
      TabIndex        =   14
      Top             =   4695
      Width           =   6360
   End
   Begin VB.ComboBox cmbRamoAtividade 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmEmarketing.frx":0000
      Left            =   1635
      List            =   "frmEmarketing.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2430
      Width           =   4305
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1635
      TabIndex        =   8
      Top             =   1830
      Width           =   4290
      Begin VB.OptionButton optFisica 
         BackColor       =   &H00505050&
         Caption         =   "Fisica"
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
         Left            =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optJuridica 
         BackColor       =   &H00505050&
         Caption         =   "Jurídica"
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
         Left            =   1125
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   1155
      End
      Begin VB.OptionButton optOrgaoPublico 
         BackColor       =   &H00505050&
         Caption         =   "Orgão Púbico"
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
         Left            =   2430
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   1845
      End
   End
   Begin VB.TextBox txtCEP 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1635
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2895
      Width           =   4290
   End
   Begin VB.TextBox txtNomeContato 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1635
      TabIndex        =   2
      Top             =   1425
      Width           =   4290
   End
   Begin VB.TextBox txtEmarketing 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1635
      TabIndex        =   1
      Top             =   960
      Width           =   4290
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmEmarketing.frx":0004
      Top             =   4275
   End
   Begin Project1.chameleonButton cmdGrava 
      Height          =   405
      Left            =   5370
      TabIndex        =   13
      Top             =   4905
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
      MICON           =   "frmEmarketing.frx":0238
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPessoa 
      BackStyle       =   0  'Transparent
      Caption         =   "Pessoa"
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
      Height          =   285
      Left            =   495
      TabIndex        =   7
      Top             =   1950
      Width           =   840
   End
   Begin VB.Label lblCEP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C E P"
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
      Left            =   495
      TabIndex        =   6
      Top             =   2970
      Width           =   585
   End
   Begin VB.Label lblRamoAtivida 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atividade"
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
      Left            =   495
      TabIndex        =   5
      Top             =   2475
      Width           =   1005
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   495
      TabIndex        =   4
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
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
      Left            =   495
      TabIndex        =   0
      Top             =   975
      Width           =   675
   End
End
Attribute VB_Name = "frmEmarketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wTipoPessoa As String
Dim wVendedor As Integer

Private Sub cmbRamoAtividade_GotFocus()
 '  cmbRamoAtividade.SelStart = 0
 '  cmbRamoAtividade.SelLength = Len(cmbRamoAtividade.Text)
End Sub

Private Sub cmbRamoAtividade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub cmbRamoAtividade_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me

        
End If
End Sub

Private Sub cmdGrava_Click()
        If Trim(txtNomeContato.Text) = "" Then
            MsgBox "O campo NOME não pode estar em branco." & _
                   "Informe um nome para contato.", vbCritical, Me.Caption
            
            txtNomeContato.SetFocus
            Exit Sub
        End If
        
        If Trim(cmbRamoAtividade.Text) = "" Then
            MsgBox "Selecione um RAMO DE ATIVIDADE.", vbCritical, Me.Caption
            
            cmbRamoAtividade.SetFocus
            Exit Sub
        End If
        
        If Trim(txtCEP.Text) = "" Then
            MsgBox "O campo CEP não pode estar em branco." & _
                   "Informe um cep para contato.", vbCritical, Me.Caption
            
            txtCEP.SetFocus
            Exit Sub
        ElseIf IsNumeric(txtCEP.Text) = False Then
            MsgBox "CEP inválido." & _
                   vbLf & "Informe apenas números.", vbCritical, Me.Caption
        
            txtCEP.SetFocus
            Exit Sub
        End If
                
        If Trim(txtEmarketing.Text) = "" Then
           MsgBox "O campo E-MAIL não pode estar em branco." & _
                  "Informe um e-mail para contato.", vbCritical, Me.Caption
           
           txtEmarketing.SetFocus
           Exit Sub
        End If
        
        If ValidaEMail(txtEmarketing.Text) = True Then
           Call IncluirEMarketing
        Else
           MsgBox "E-Mail inválido!", vbCritical, Me.Caption
           txtEmarketing.SetFocus
        End If
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
    frmPedido.txtPesquisar.SetFocus
End Sub

Private Sub Form_Load()
  Call AjustaTela(frmEmarketing)
  Call LimpaCampos
  Call PreencheCombo
  wVendedor = Val(frmPedido.txtVendedor.Text)
  'Skin1.LoadSkin App.Path & "\Skin\royaleblue.skn"
 ' Skin1.LoadSkin App.Path & "\Skin\corona2.skn"
 ' Skin1.ApplySkin Me.hwnd
    
    
    
End Sub

Private Sub optFisica_Click()
    Call PreencheCombo
End Sub

Private Sub optFisica_LostFocus()
    Call PreencheCombo
End Sub

Private Sub optJuridica_Click()
    Call PreencheCombo
End Sub

Private Sub optJuridica_LostFocus()
    Call PreencheCombo
End Sub

Private Sub txtCEP_GotFocus()
   txtCEP.SelStart = 0
   txtCEP.SelLength = Len(txtCEP.Text)
End Sub

Private Sub txtCEP_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''
'''        If Trim(txtNomeContato.Text) = "" Then
'''            MsgBox "O campo NOME não pode estar em branco." & _
'''                   "Informe um nome para contato.", vbCritical, Me.Caption
'''
'''            txtNomeContato.SetFocus
'''            Exit Sub
'''        End If
'''
'''        If Trim(cmbRamoAtividade.Text) = "" Then
'''            MsgBox "Selecione um RAMO DE ATIVIDADE.", vbCritical, Me.Caption
'''
'''            cmbRamoAtividade.SetFocus
'''            Exit Sub
'''        End If
'''
'''        If Trim(txtCEP.Text) = "" Then
'''            MsgBox "O campo CEP não pode estar em branco." & _
'''                   "Informe um cep para contato.", vbCritical, Me.Caption
'''
'''            txtCEP.SetFocus
'''            Exit Sub
'''        ElseIf IsNumeric(txtCEP.Text) = False Then
'''            MsgBox "CEP inválido." & _
'''                   vbLf & "Informe apenas números.", vbCritical, Me.Caption
'''
'''            txtCEP.SetFocus
'''            Exit Sub
'''        End If
'''
'''        If Trim(txtEmarketing.Text) = "" Then
'''           MsgBox "O campo E-MAIL não pode estar em branco." & _
'''                  "Informe um e-mail para contato.", vbCritical, Me.Caption
'''
'''           txtEmarketing.SetFocus
'''           Exit Sub
'''        End If
'''
'''        If ValidaEMail(txtEmarketing.Text) = True Then
'''           Call IncluirEMarketing
'''        Else
'''           MsgBox "E-Mail inválido!", vbCritical, Me.Caption
'''           txtEmarketing.SetFocus
'''        End If
'''
'''    End If
 If KeyAscii = 27 Then
    Unload Me
        
End If
End Sub

Private Sub txtEmarketing_GotFocus()
   txtEmarketing.SelStart = 0
   txtEmarketing.SelLength = Len(txtEmarketing.Text)

End Sub

Private Sub txtEmarketing_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtEmarketing_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       Unload Me
       frmPedido.txtPesquisar.SetFocus

    
       
    End If

End Sub

Private Sub PreencheCombo()

    If optFisica.Value = True Then
        wTipoPessoa = "F"
    End If
    If optJuridica.Value = True Then
        wTipoPessoa = "J"
    End If
    If optOrgaoPublico.Value = True Then
        wTipoPessoa = "O"
    End If
    
    sql1 = ""
    sql1 = "Select RMO_DescricaoRamo From fin_RamoAtividade " & _
           "Where RMO_Pessoa = '" & wTipoPessoa & "' Order By RMO_DescricaoRamo"
   
    rdoRamo.CursorLocation = adUseClient
    rdoRamo.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
          
    cmbRamoAtividade.Clear
          
    Do While Not rdoRamo.EOF
        cmbRamoAtividade.AddItem Trim(rdoRamo.Fields("RMO_DescricaoRamo"))
        rdoRamo.MoveNext
    Loop
    rdoRamo.Close
          
End Sub

Private Sub IncluirEMarketing()

On Error GoTo ErroIncluirEMKT
    
    If rdoRamo.State = 1 Then rdoRamo.Close
    
    'SQL1 = ""
    sql1 = "Select MKT_Email From EMKTLoja Where MKT_Email = '" & txtEmarketing.Text & "'"
    
    rdoRamo.CursorLocation = adUseClient
    rdoRamo.Open sql1, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    If rdoRamo.EOF = False Then
        MsgBox "Já existe cadastrado do e-mail " & Trim(txtEmarketing.Text) & " " & _
               vbLf & "Informe um outro e-mail para cadastro.", vbInformation, Me.Caption
    
    Else
        sql1 = ""
        sql1 = "Insert Into EMKTLoja(MKT_Loja,MKT_Vendedor,MKT_Email,MKT_DataCadastro,MKT_RamoAtividade,MKT_Nome,MKT_CEP,MKT_Situacao) " & _
                "Select '" & Trim(wLoja) & "'," & wVendedor & ",'" & Trim(txtEmarketing.Text) & "','" & Format(Date, "yyyy/mm/dd") & "',RMO_Codigo,'" & _
                Trim(txtNomeContato.Text) & "','" & Trim(txtCEP.Text) & "','A' From fin_RamoAtividade " & _
                "Where RMO_Pessoa = '" & wTipoPessoa & "' And RMO_DescricaoRamo = '" & cmbRamoAtividade.Text & "'"
    
            adoCNLoja.BeginTrans
            adoCNLoja.Execute sql1
            adoCNLoja.CommitTrans
            MsgBox "Cadastro incluído com sucesso!", vbInformation, Me.Caption
            Call LimpaCampos
    End If
    
    rdoRamo.Close
    txtEmarketing.SetFocus
    Exit Sub

ErroIncluirEMKT:
        MsgBox "Falha ao incluir E-Marketing." & _
               vbLf & Err.description, vbCritical, Me.Caption
        adoCNLoja.RollbackTrans
        
        If rdoRamo.State = 1 Then rdoRamo.Close
        Exit Sub
        
End Sub

Private Sub LimpaCampos()
    txtEmarketing.Text = ""
    txtNomeContato.Text = ""
    txtCEP.Text = ""
    cmbRamoAtividade.Clear
    optFisica.Value = True
End Sub

Public Function ValidaEMail(sEMail As String) As Boolean
Dim nCharacter As Integer
Dim Count As Integer
Dim sLetra As String
    
    'Verifica se o e-mail tem no MÍNIMO 5 caracteres (a@b.c)
    If Len(sEMail) < 5 Then
        'O e-mail é inválido, pois tem menos de 5 caracteres
        ValidaEMail = False
        Exit Function
    End If

    'Verificar a existencia de arrobas no e-mail
    For nCharacter = 1 To Len(sEMail)
        If Mid(sEMail, nCharacter, 1) = "@" Then
            'OPA!!! Achou uma arroba!!! Soma 1 ao contador
            Count = Count + 1
        End If
    Next
    
    'Verifica o número de arrobas. TEM que ter """UMA""" arroba
    If Count <> 1 Then
        'O e-mail é inválido, pois tem 0 ou mais de 1 arroba
        ValidaEMail = False
        Exit Function
    Else
        'O e-mail tem 1 arroba. Verificar a posição da arroba
        If InStr(sEMail, "@") = 1 Then
            'O e-mail é inválido, pois começa com uma @
            ValidaEMail = False
            Exit Function
        ElseIf InStr(sEMail, "@") = Len(sEMail) Then
            'O e-mail é inválido, pois termina com uma @
            ValidaEMail = False
            Exit Function
        End If
    End If

    nCharacter = 0
    Count = 0
    
    'Verificar a existencia de pontos (.) no e-mail
    For nCharacter = 1 To Len(sEMail)
        If Mid(sEMail, nCharacter, 1) = "." Then
            'OPA!!! Achou um ponto!!! Soma 1 ao contador
            Count = Count + 1
        End If
    Next

    'Verifica o número de pontos. TEM que ter PELO MENOS UM ponto.
    If Count < 1 Then
        'O e-mail é inválido, pois não tem pontos.
        ValidaEMail = False
        Exit Function
    Else
        'O e-mail tem pelo menos 1 ponto. Verificar a posição do ponto:
        If InStr(sEMail, ".") = 1 Then
            'O e-mail é inválido, pois começa com um ponto
            ValidaEMail = False
            Exit Function
        ElseIf InStr(sEMail, ".") = Len(sEMail) Then
            'O e-mail é inválido, pois termina com um ponto.
            ValidaEMail = False
            Exit Function
        ElseIf InStr(InStr(sEMail, "@"), sEMail, ".") = 0 Then
            'O e-mail é inválido, pois termina com um ponto.
            ValidaEMail = False
            Exit Function
        End If
    End If

    nCharacter = 0
    Count = 0

    'Verifica se o e-mail não tem pontos consecutivos (..) após a arroba .
    If InStr(sEMail, "..") > InStr(sEMail, "@") Then
        'O e-mail é inválido, tem pontos consecutivos após o @.
        ValidaEMail = False
        Exit Function
    End If

    'Verifica se o e-mail tem caracteres inválidos
    For nCharacter = 1 To Len(sEMail)
        sLetra = Mid$(sEMail, nCharacter, 1)
        If Not (LCase(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
            'O e-mail é inválido, pois tem caracteres inválidos
            ValidaEMail = False
            Exit Function
        End If
    Next

    nCharacter = 0
    ValidaEMail = True

End Function

Private Sub txtNomeContato_GotFocus()
   txtNomeContato.SelStart = 0
   txtNomeContato.SelLength = Len(txtNomeContato.Text)

End Sub

Private Sub txtNomeContato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
   cmdGrava_Click
End If
End Sub

Private Sub txtNomeContato_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    Unload Me

End If
End Sub
