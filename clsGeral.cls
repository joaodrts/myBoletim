VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ChamaFormAguardar(FormOrigem As Object, Titulo As String, Mensagem As String, ValorMinimo As Long, ValorMaximo As Long)
    frmAguarde.Load FormOrigem, Titulo, Mensagem, ValorMinimo, ValorMaximo
End Sub

Public Sub ChamaAtualizacaoAguardar(Mensagem As String, ValorAtual As Long)
    frmAguarde.AtualizaProgresso Mensagem, ValorAtual
End Sub

Public Sub FecharForm(FormVoltar As Object, FormFechar As Object)
    FormFechar.Fechar
    On Error Resume Next
    FormVoltar.Grid.SetFocus
    On Error GoTo 0
End Sub

Public Sub HabilitaForm(Form As Object, value As Boolean)
    Form.Enabled = value
End Sub

