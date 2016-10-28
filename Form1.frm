VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   2955
   ClientTop       =   1485
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   2925
   Begin VB.TextBox txtPrecoAlternativo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6885
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox txtDescricaoAlternativa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   15
      Visible         =   0   'False
      Width           =   6600
   End
   Begin VB.TextBox txtDescricaoEspecial 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   6600
   End
   Begin VB.TextBox txtPrecoEspecial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6885
      TabIndex        =   4
      Top             =   405
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Frame fraHistEstq 
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8730
      Begin MSFlexGridLib.MSFlexGrid grdHistEstq 
         Height          =   870
         Left            =   1905
         TabIndex        =   2
         Top             =   165
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   1535
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedRows       =   2
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   192
         BackColorBkg    =   12632256
         MergeCells      =   1
      End
      Begin VB.Label lblTendencia 
         AutoSize        =   -1  'True
         Caption         =   "Tendência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   195
         TabIndex        =   3
         Top             =   360
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdEstLoja 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   3810
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      ForeColorFixed  =   16576
      BackColorSel    =   -2147483647
      BackColorBkg    =   16761024
      GridColor       =   0
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "<Loja |>Estq.|>Tran.|>Rom.|>Min.|>Max."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
