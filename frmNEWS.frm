VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form FrmNews 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   4290
   ClientTop       =   2475
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNoticia 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00832B0D&
      Height          =   4455
      Left            =   2805
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   6135
   End
   Begin VB.ListBox ListaNoticias 
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   9960
      Width           =   6135
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdNoticias 
      Height          =   4575
      Left            =   150
      TabIndex        =   2
      Top             =   765
      Width           =   6225
      _cx             =   10980
      _cy             =   8070
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   8596237
      BackColorFixed  =   8596237
      ForeColorFixed  =   16777215
      BackColorSel    =   12443628
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   16777215
      FloodColor      =   0
      SheetBorder     =   16777215
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmNEWS.frx":0000
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   16777215
      WallPaperAlignment=   9
   End
   Begin VB.Label lblLembreMe 
      Alignment       =   2  'Center
      BackColor       =   &H00832B0D&
      Caption         =   "NEWS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6225
   End
End
Attribute VB_Name = "FrmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Dim SQL As String
 Dim rdoNoticias As New ADODB.Recordset
 Dim rdoListaNoticia As New ADODB.Recordset
 Dim rdoLeitura As New ADODB.Recordset
 
Private Sub Form_Activate()

txtNoticia.left = grdNoticias.left
txtNoticia.top = grdNoticias.top
txtNoticia.Width = grdNoticias.Width
txtNoticia.Height = grdNoticias.Height

End Sub

Private Sub Form_Load()

txtNoticia.Visible = False
Call AjustaTela(FrmNews)
Call CarregaNoticias
Call LeituraGrid

txtNoticia.BackColor = vbWhite


grdNoticias.Row = 0

End Sub
Private Sub CarregaNoticias()
Dim I As Byte
ListaNoticias.Clear
grdNoticias.Rows = 1
grdNoticias.Col = 0
grdNoticias.Row = grdNoticias.Rows - 1
        grdNoticias.CellBackColor = &HFFFFFF


SQL = ""
SQL = "select NWS_Codigo,NWS_Assunto,NWS_Lido from News where NWS_Usuario = " & left(frmPedido.txtVendedor.Text, 3)

    rdoNoticias.CursorLocation = adUseClient
    rdoNoticias.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic


        Do While Not rdoNoticias.EOF
           
           grdNoticias.AddItem rdoNoticias("NWS_Codigo") & Chr(9) & Trim(rdoNoticias("NWS_Assunto")) & Chr(9) & rdoNoticias("NWS_Lido")
           rdoNoticias.MoveNext
           
        Loop
    
     rdoNoticias.Close

End Sub

Private Sub LeituraGrid()
    Dim rdoLeituraGrid As New ADODB.Recordset
    Dim I As Integer
    Dim lido As String
    
    lido = grdNoticias.TextMatrix(grdNoticias.Row, 2)


   For I = 1 To grdNoticias.Rows - 1
   
        grdNoticias.Col = 1
        grdNoticias.Row = I
        
        If grdNoticias.TextMatrix(grdNoticias.Row, 2) = "S" Then
            grdNoticias.CellBackColor = vbWhite
            
        ElseIf grdNoticias.TextMatrix(grdNoticias.Row, 2) = "N" Then
            grdNoticias.CellBackColor = RGB(244, 236, 216)
        End If
    
    Next I
    
        
End Sub

Private Sub grdNoticias_DblClick()
Dim codigo As Integer

txtNoticia.Visible = True
codigo = grdNoticias.TextMatrix(grdNoticias.Row, 0)
txtNoticia.Text = grdNoticias.TextMatrix(grdNoticias.Row, 1)

grdNoticias.CellBackColor = vbWhite

SQL = ""
SQL = "select NWS_Mensagem from news where NWS_Codigo = '" & codigo & "'"

rdoLeitura.CursorLocation = adUseClient
rdoLeitura.Open SQL, adoCNLoja, adOpenForwardOnly, adLockPessimistic

txtNoticia.Text = Replace(rdoLeitura("NWS_Mensagem"), "[10]", vbNewLine)

rdoLeitura.Close


SQL = "update news set NWS_lido = 'S' where NWS_Codigo = '" & codigo & "'"
adoCNLoja.Execute SQL

   ' Call frmPedido.verificaNovasMensagem(True)

End Sub


Private Sub grdNoticias_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    Unload Me
End If

End Sub

Private Sub ListaNoticias_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    Unload Me
End If

End Sub

Private Sub txtNoticia_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    txtNoticia.Visible = False
End If

End Sub


