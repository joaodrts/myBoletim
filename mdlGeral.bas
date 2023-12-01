Attribute VB_Name = "mdlGeral"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Sub CriaTabelaNotExists()
    cn.Execute ("CREATE TABLE IF NOT EXISTS Professores ( " _
                    & "Codigo int not null, " _
                    & "Nome VARCHAR(100) NOT NULL, " _
                    & "Area VARCHAR(100) NOT NULL, " _
                    & "Primary key(Codigo)" _
                    & ");")
                    
    cn.Execute ("CREATE TABLE IF NOT EXISTS Materias ( " _
                    & "Codigo int not null, " _
                    & "NomeMateria VARCHAR(100) NOT NULL, " _
                    & "TotalHoras int, " _
                    & "CodigoProfessor int not null, " _
                    & "Primary key(Codigo)" _
                    & ");")
                    
    cn.Execute ("CREATE TABLE IF NOT EXISTS Alunos ( " _
                    & "Codigo int not null, " _
                    & "RA_Aluno int, " _
                    & "NomeAluno VARCHAR(100) NOT NULL, " _
                    & "Primary key(Codigo)" _
                    & ");")
    
    cn.Execute ("CREATE TABLE IF NOT EXISTS Notas ( " _
                    & "Sequencia serial, " _
                    & "CodigoAluno int not null, " _
                    & "Nota int not null, " _
                    & "HorasAproveitadas int not null, " _
                    & "CodigoMateria int not null " _
                    & ");")
End Sub

Public Function NaoImportado(CodigoRegistro As Long, Tabela As String) As Boolean

    Dim nReg As Long
    Dim rs As New ADODB.Recordset
    
    If Tabela <> "notas" Then
        Set rs = cn.Execute("select count(*) as nreg from " & Tabela & " where codigo = " & CodigoRegistro & "")
    Else
        Set rs = cn.Execute("select count(*) as nreg from " & Tabela & " where sequencia = " & CodigoRegistro & "")
    End If
    NaoImportado = (rs!nReg = 0)

End Function

Public Function NaoImportadoNota(CodigoAluno As Long, CodigoMateria As Long) As Boolean

    Dim nReg As Long
    Dim rs As New ADODB.Recordset
    
    Set rs = cn.Execute("select count(*) as nreg from notas where codigoaluno = " & CodigoAluno & " and codigomateria = " & CodigoMateria & "")
    NaoImportadoNota = (rs!nReg = 0)

End Function

Public Sub MensagemErro(Frase As String)
    MsgBox Frase, vbCritical, "Erro"
End Sub

Public Function ReplaceText(Texto As Variant, TextoNovo As String) As String

    If IsNull(Texto) Then
        Texto = TextoNovo
    End If
    
    Texto = Replace(Texto, "'", "")
    Texto = Replace(Texto, Chr(13), "")
    
    ReplaceText = UCase(Texto)

End Function

Public Function MensagemQuestionario(Frase As String) As Boolean
    If MsgBox(Frase, vbYesNo + vbInformation, "Questionário") = vbYes Then
        MensagemQuestionario = True
    End If
End Function

Public Sub MensagemInformativa(Frase As String)
    MsgBox Frase, vbInformation, "Aviso"
End Sub

Public Function TrataTexto(Texto As String) As String
    
    Texto = Replace(Texto, "ç", "c")
    Texto = Replace(Texto, "Ç", "C")
    Texto = Replace(Texto, "ã", "a")
    Texto = Replace(Texto, "Ã", "A")

    TrataTexto = Texto
    
End Function

