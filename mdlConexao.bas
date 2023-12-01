Attribute VB_Name = "mdlConexao"
Public cn As ADODB.Connection

Function Connecta()

        On Error GoTo erroConexao
        
        Set cn = New ADODB.Connection
        Dim StringConexao As String
        
        StringConexao = "Provider=MSDASQL;Driver=PostgreSQL ODBC Driver(ANSI); " _
                      & "SERVER=localhost;PORT=5432;UID=postgres;Pwd=vssql;DATABASE=ImportacaoBoletim; " _
                      & "ByteaAsLongVarBinary=1;BoolsAsChar=0;"
        
        With cn
            .CursorLocation = adUseClient
            .ConnectionString = StringConexao
            .Open
        End With
        
        Connecta = True
        Exit Function
        
erroConexao:

        MensagemInformativa "Ocorreu um erro na conexão com banco de dados! " & vbCrLf & Err.Description
End Function




