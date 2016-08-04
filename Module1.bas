Attribute VB_Name = "Module1"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Ret As String
Public Index As Integer
Public indexs As String


Public Sub WriteINI(FileName As String, Section As String, key As String, Text As String)
WritePrivateProfileString Section, key, Text, FileName
End Sub

Public Function ReadINI(FileName As String, Section As String, key As String)
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Section, key, "", Ret, Len(Ret), FileName)
    If RetLen = 0 Then
        Exit Function
    End If
    Ret = left$(Ret, RetLen)
    ReadINI = Ret
End Function
Public Function LimpaGrid(ByRef GradeUsu)
    GradeUsu.Rows = GradeUsu.FixedRows + 1
    GradeUsu.AddItem ""
    GradeUsu.RemoveItem GradeUsu.FixedRows
End Function
'Public Function ConectaBancoLoja()
'  If ConectaOdbcBalcao(adoCNLoja, Usuario, Senha) = False Then
'        MsgBox "Não foi possivel conectar-se ao banco de dados do Balcão", vbCritical, "Aviso"
'        Exit Function
'  Else
'        MsgBox "Conexão estabelecida com sucesso", vbInformation
'  End If
'End Function

