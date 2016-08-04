VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9435
   ClientLeft      =   4380
   ClientTop       =   2100
   ClientWidth     =   6585
   LinkTopic       =   "Form3"
   ScaleHeight     =   9435
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2250
      Left            =   1020
      TabIndex        =   0
      Top             =   1890
      Width           =   2835
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coloque num Form um CommandButton
'chamado Command1.

'No Declarations:
Private Declare Function CreateRoundRectRgn Lib _
        "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
        (ByVal hWnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Private Declare Function GetClientRect Lib "user32" _
        (ByVal hWnd As Long, lpRect As Rect) As Long
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Sub Retangulo(m_hWnd As Long, Fator As Byte)
  Dim RGN As Long
  Dim RC As Rect
  Call GetClientRect(m_hWnd, RC)
  RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, _
                           RC.Bottom, Fator, Fator)
  SetWindowRgn m_hWnd, RGN, True
End Sub
'Fator é a distância da curvatura do canto arredondado

'No evento click do CommandButton:
Private Sub Command1_Click()
  Me.BackColor = &H808080 'Apenas para destacar a cor
  'Coloca o formulário com os cantos arredondados
  'e fator 80 de área
  Retangulo Me.hWnd, 80
  Retangulo Command1.hWnd, 30
End Sub

