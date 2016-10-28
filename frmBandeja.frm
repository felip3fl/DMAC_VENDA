VERSION 5.00
Begin VB.Form frmBandeja 
   BackColor       =   &H00000000&
   Caption         =   "DMAC Venda"
   ClientHeight    =   10125
   ClientLeft      =   2640
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "frmBandeja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   18493.15
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   Begin VB.Image Image1 
      Height          =   6210
      Left            =   14805
      Picture         =   "frmBandeja.frx":23FA
      Top             =   1770
      Width           =   6420
   End
   Begin VB.Image imgTarefas 
      Height          =   10380
      Left            =   -465
      Picture         =   "frmBandeja.frx":358F9
      Top             =   240
      Width           =   15210
   End
End
Attribute VB_Name = "frmBandeja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    frmPedido.Show 1
End Sub

Private Sub Form_Load()
    imgTarefas.top = 0
    imgTarefas.left = 0
    Me.Height = (imgTarefas.Height) - 100
    Me.Width = (imgTarefas.Width)
    Me.top = -500
    Me.left = -100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

