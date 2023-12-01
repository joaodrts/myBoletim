VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{9B3742B3-4639-4F01-8EFC-BEF3E8D567DE}#1.0#0"; "Panel2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   8460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDExcel 
      Left            =   1440
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PanelFx32.PanelFx Painel 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13150
      TitleCaption    =   "Importação de Planilha - Boletim"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackGroundStyle =   1
      gCTitleEnd      =   4194304
      gCPanelStart    =   16777215
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "E&xcluir"
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "E&ditar"
         Height          =   495
         Left            =   1560
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "&Adicionar"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Frame fraPlanilha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Selecione a Planilha: "
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   12975
         Begin VB.CommandButton cmdImportar 
            Caption         =   "&Importar"
            Height          =   495
            Left            =   11640
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblArquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Arquivo (.csv)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label txtArquivo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   11415
         End
      End
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   11640
         TabIndex        =   7
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Height          =   615
         Left            =   10080
         TabIndex        =   6
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Frame fraFiltros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Filtros"
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   120
         TabIndex        =   1
         Top             =   1250
         Width           =   12975
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtPesquisar 
            Height          =   300
            Left            =   2400
            TabIndex        =   2
            Top             =   480
            Width           =   10455
         End
         Begin VB.Label lblCriterio 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Critério <<ALT + C>>"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1950
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Pesquisar <<ALT + P>>"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2400
            TabIndex        =   4
            Top             =   240
            Width           =   2100
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   3735
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   12975
         _cx             =   22886
         _cy             =   6588
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         ShowComboButton =   1
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Obs: quando for realizar a exportação, deve informar a extenção .csv no nome do arquivo."
         Height          =   495
         Left            =   5040
         TabIndex        =   17
         Top             =   6840
         Width           =   5055
      End
      Begin VB.Label lblTotalizador 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Total de Registros: 000000"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   6720
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Myf As New clsGeral

Private Sub cmdAdicionar_Click()
    frmCadastro.Show vbModal
    CarregaGrid
End Sub

Private Sub cmdEditar_Click()
    Editar
    CarregaGrid
End Sub

Private Sub cmdExcluir_Click()
    ExcluirNota
    CarregaGrid
End Sub

Private Sub cmdExportar_Click()

    With Grid
            
        If .Rows = 1 Then Exit Sub
            
        Dim dirArquivo As String
        Dim x, y As Long
        Dim linha As String
        
        CDExcel.DialogTitle = "Selecione uma pasta"
        CDExcel.Flags = &H1
        CDExcel.ShowOpen
            
        If CDExcel.FileName <> "" Then
            
            dirArquivo = CDExcel.FileName
                    
            Open dirArquivo For Output As #1
            
            For x = 0 To .Rows - 1
            
                For y = 0 To .Cols - 1
                    linha = linha & TrataTexto(.TextMatrix(x, y)) & ";"
                Next
                
                Print #1, linha
                linha = ""
                
            Next
            
            Close #1
            
        End If
    
    End With
       
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdImportar_Click()
    
    With CD
    
        txtPesquisar.Text = ""
        .FileName = ""
        .Filter = "Selecione uma planilha (*.csv) |*.csv"
        .ShowOpen
        
        txtArquivo.Caption = .FileName
        
        If .FileName <> "" Then
            ImportarDados Replace(.FileName, .FileTitle, ""), .FileTitle
        End If
    
    End With
    
End Sub

Private Function ImportarDados(Diretorio As String, Arquivo As String) As Boolean
    
    Dim rsPlanilha As New ADODB.Recordset
    Dim cnPlanilha As New ADODB.Connection
    Dim nReg, nRegTotal As Long
    Dim sql As String
    Dim ValorPlanilha As Variant
        
    On Error GoTo Error:

        cnPlanilha.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & Diretorio & ";Extensions=asc,csv,tab,txt;"
        rsPlanilha.Open "select * from " & Arquivo & "", cnPlanilha, adOpenStatic, adLockReadOnly, adCmdText
        
        'ValorPlanilha(0) = Codigo Professor
        'ValorPlanilha(1) = Nome Professor
        'ValorPlanilha(2) = Area
        'ValorPlanilha(3) = Codigo Materia
        'ValorPlanilha(4) = Nome Materia
        'ValorPlanilha(5) = Total Horas
        'ValorPlanilha(6) = RA Aluno
        'ValorPlanilha(7) = Nome Aluno
        'ValorPlanilha(8) = Codigo Aluno
        'ValorPlanilha(9) = Codigo Materia
        'ValorPlanilha(10) = Nota
        'ValorPlanilha(11) = Horas Aproveitadas
        
        nReg = 0
        nRegTotal = rsPlanilha.RecordCount
    
        Myf.ChamaFormAguardar Me, "Importando Dados... Aguarde!", "Dados importado " & nReg & " de " & nRegTotal & "", 0, nRegTotal
        DoEvents
        
        Myf.HabilitaForm Me, False
        
        While Not rsPlanilha.EOF
        
            nReg = nReg + 1
            ValorPlanilha = Split(rsPlanilha(0), ";")
            
            If NaoImportado(CLng(ValorPlanilha(0)), "professores") Then
                sql = "INSERT INTO professores values(" & ValorPlanilha(0) & ","
                sql = sql & "'" & ReplaceText(ValorPlanilha(1), "") & "',"
                sql = sql & "'" & ReplaceText(ValorPlanilha(2), "") & "')"
                cn.Execute (sql)
            End If
                
            If NaoImportado(CLng(ValorPlanilha(3)), "materias") Then
                sql = "INSERT INTO materias values(" & ValorPlanilha(3) & ","
                sql = sql & "'" & ReplaceText(ValorPlanilha(4), "") & "',"
                sql = sql & "'" & ReplaceText(ValorPlanilha(5), "") & "',"
                sql = sql & "'" & ReplaceText(ValorPlanilha(0), "") & "')"
                cn.Execute (sql)
            End If
                
            If NaoImportado(CLng(ValorPlanilha(8)), "alunos") Then
                sql = "INSERT INTO alunos values(" & ValorPlanilha(8) & ","
                sql = sql & "'" & ReplaceText(ValorPlanilha(6), "") & "',"
                sql = sql & "'" & ReplaceText(ValorPlanilha(7), "") & "')"
                cn.Execute (sql)
            End If
            
            If NaoImportadoNota(CLng(ValorPlanilha(8)), CLng(ValorPlanilha(3))) Then
                sql = "INSERT INTO notas(codigoaluno, nota, horasaproveitadas, codigomateria) values(" & ValorPlanilha(8) & ","
                sql = sql & "" & ReplaceText(ValorPlanilha(10), "") & ","
                sql = sql & "" & ReplaceText(ValorPlanilha(11), "") & ","
                sql = sql & "" & ReplaceText(ValorPlanilha(9), "") & ")"
                cn.Execute (sql)
            End If
                                    
            Myf.ChamaAtualizacaoAguardar "Dados importado " & nReg & " de " & nRegTotal & "", CLng(nReg)
            
            CarregaGrid
            
            rsPlanilha.MoveNext
        Wend
        
        Myf.HabilitaForm Me, True
        Myf.FecharForm Me, frmAguarde
    
        rsPlanilha.Close
        cnPlanilha.Close
        Exit Function

Error:
    MensagemErro "Ocorreu algum erro na importação dos dados! Error: " & Err.Description
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        
        Case vbKeyI And Shift = 4:
            KeyCode = 0
            cmdImportar_Click
        
        Case vbKeyA And Shift = 4:
            KeyCode = 0
            cmdAdicionar_Click
            
        Case vbKeyD And Shift = 4:
            KeyCode = 0
            cmdEditar_Click
        
        Case vbKeyX And Shift = 4:
            KeyCode = 0
            cmdExcluir_Click
            
        Case vbKeyE And Shift = 4:
            KeyCode = 0
            cmdExportar_Click
            
        Case vbKeyF And Shift = 4:
            KeyCode = 0
            cmdFechar_Click
        
        Case vbKeyC And Shift = 4:
            KeyCode = 0
            cboCriterio.SetFocus
        
        Case vbKeyP And Shift = 4:
            KeyCode = 0
            txtPesquisar.SetFocus
    
    End Select

End Sub

Private Sub Form_Load()

    With Painel
        .Top = 0
        .Left = 0
    
        Me.Width = .Width
        Me.Height = .Height
    End With
    
    If Not Connecta Then End
    CriaTabelaNotExists
 
    CarregaCombo
    FormataGrid
    CarregaGrid
    
    
End Sub

Private Sub FormataGrid()
    
    With Grid
        
        .Rows = 1
        .Cols = 12
        .FixedCols = 0
        
        .FormatString = "Cod Professor|Nome Professor|Area|Cod Materia|Materia|Total Horas|RA Aluno|Cod Aluno|Nome Aluno|Nota|Horas Aproveitadas|Status Aprovação|Sequencia"
        
        .ColHidden(0) = True
        .ColHidden(3) = True
        .ColHidden(7) = True
        .ColHidden(12) = True
        
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1300
        .ColWidth(8) = 1500
        .ColWidth(9) = 1000
        .ColWidth(10) = 1700
        
        .AllowSelection = False
        .ExplorerBar = flexExSort
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
    
    End With

End Sub

Public Sub CarregaGrid()

    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    sql = "select * from alunos a " & vbCrLf
    sql = sql & "left join notas n on a.codigo = n.codigoaluno " & vbCrLf
    sql = sql & "left join materias m on n.codigomateria = m.codigo " & vbCrLf
    sql = sql & "left join professores p on m.codigoprofessor = p.codigo where sequencia is not null "
    
    If txtPesquisar.Text <> "" Then
        
        Select Case cboCriterio.ListIndex
            
            Case 0:
                If IsNumeric(txtPesquisar.Text) Then
                    sql = sql & " and codigoaluno = " & txtPesquisar.Text & ""
                End If
            Case 1:
                sql = sql & " and nomealuno like '%" & Replace(UCase(txtPesquisar.Text), "'", "") & "%'"
        
        End Select
    
    End If
    
    Set rs = cn.Execute(sql)
    
    Grid.Rows = 1
    
    While Not rs.EOF
    
        With Grid
        
            .Rows = .Rows + 1
            
            .MergeCells = flexMergeRestrictColumns

            .MergeCol(8) = True
            .MergeCol(11) = True
            
            .TextMatrix(.Rows - 1, 0) = rs!codigoprofessor
            .TextMatrix(.Rows - 1, 1) = rs!nome
            .TextMatrix(.Rows - 1, 2) = rs!area
            .TextMatrix(.Rows - 1, 3) = rs!CodigoMateria
            .TextMatrix(.Rows - 1, 4) = rs!nomemateria
            .TextMatrix(.Rows - 1, 5) = rs!totalhoras
            .TextMatrix(.Rows - 1, 6) = rs!ra_aluno
            .TextMatrix(.Rows - 1, 7) = rs!CodigoAluno
            .TextMatrix(.Rows - 1, 8) = rs!nomealuno
            .TextMatrix(.Rows - 1, 9) = rs!nota
            .TextMatrix(.Rows - 1, 10) = rs!horasaproveitadas
            .TextMatrix(.Rows - 1, 11) = GetStatusAprovacao(rs!CodigoAluno)
            .TextMatrix(.Rows - 1, 12) = rs!sequencia
        
        End With
        
        SetaNregistrosLabel
        
        rs.MoveNext
    Wend
 
    If Grid.Rows > 1 Then Grid.Select Grid.Rows - 1, 0
 
End Sub

Private Sub SetaNregistrosLabel()
    lblTotalizador.Caption = "Total de Registros: " & Format(Grid.Rows - 1, "######000000")
End Sub

Private Sub Painel_TileClick()
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Function GetStatusAprovacao(CodigoAluno As String) As String

    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim isAprovadoMedia, isAprovadoParticipacao, isCumpriuCargaHoraria As Boolean
    Dim MediaAluno As Double
    
    isAprovadoMedia = (GetMediaAluno(CodigoAluno) >= 7)
    isCumpriuCargaHoraria = GetCumpriuCargaHoraria(CodigoAluno)
    
    If isAprovadoMedia And isCumpriuCargaHoraria Then
        isAprovadoParticipacao = GetCumpriuParticipacao(CodigoAluno)
    End If
    
    If isAprovadoMedia And isCumpriuCargaHoraria And isAprovadoParticipacao Then
        GetStatusAprovacao = "AVANÇADO"
    Else
        GetStatusAprovacao = "RETIDO"
    End If

End Function

Private Function GetMediaAluno(CodigoAluno) As Double

    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim MediaAluno As Double
    
    sql = "select ((sum(nota)) / count(*)) as media from alunos a " & vbCrLf
    sql = sql & "left join notas n on a.codigo = n.codigoaluno " & vbCrLf
    sql = sql & "left join materias m on n.codigomateria = m.codigo" & vbCrLf
    sql = sql & "left join professores p on m.codigoprofessor = p.codigo" & vbCrLf
    sql = sql & "where a.codigo = " & CodigoAluno & ""
    
    Set rs = cn.Execute(sql)
    
    If Not rs.EOF Then
        GetMediaAluno = rs!media
    Else
        GetMediaAluno = 0
    End If

End Function

Private Function GetCumpriuCargaHoraria(CodigoAluno) As Boolean

    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim perc As Double
    GetCumpriuCargaHoraria = True
    
    sql = "select totalhoras, horasaproveitadas as horas from alunos a " & vbCrLf
    sql = sql & "left join notas n on a.codigo = n.codigoaluno " & vbCrLf
    sql = sql & "left join materias m on n.codigomateria = m.codigo" & vbCrLf
    sql = sql & "left join professores p on m.codigoprofessor = p.codigo" & vbCrLf
    sql = sql & "where a.codigo = " & CodigoAluno & ""
    
    Set rs = cn.Execute(sql)
    
    While Not rs.EOF
        
        perc = (rs!horas / rs!totalhoras) * 100
        
        If perc < 75 Then GetCumpriuCargaHoraria = False: Exit Function
        
        rs.MoveNext
    Wend

End Function

Private Function GetCumpriuParticipacao(CodigoAluno) As Boolean

    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim nAprovadas, nCursadas As Long
    Dim percParticipacao As Double
    Dim perc As Double
    
    sql = "select horasaproveitadas, totalhoras from alunos a " & vbCrLf
    sql = sql & "left join notas n on a.codigo = n.codigoaluno " & vbCrLf
    sql = sql & "left join materias m on n.codigomateria = m.codigo" & vbCrLf
    sql = sql & "left join professores p on m.codigoprofessor = p.codigo" & vbCrLf
    sql = sql & "where a.codigo = " & CodigoAluno & " and (nota >= 7)"
    
    Set rs = cn.Execute(sql)

    While Not rs.EOF
        perc = ((rs!horasaproveitadas / rs!totalhoras) * 100)
        If perc >= 75 Then nAprovadas = nAprovadas + 1
        rs.MoveNext
    Wend
    
    sql = "select count(m.codigo) as nCursadas from alunos a " & vbCrLf
    sql = sql & "left join notas n on a.codigo = n.codigoaluno " & vbCrLf
    sql = sql & "left join materias m on n.codigomateria = m.codigo" & vbCrLf
    sql = sql & "left join professores p on m.codigoprofessor = p.codigo" & vbCrLf
    sql = sql & "where a.codigo = " & CodigoAluno & ""
    
    Set rs = cn.Execute(sql)
    
    If Not rs.EOF Then nCursadas = rs!nCursadas
    
    percParticipacao = (nAprovadas / nCursadas) * 100
    
    GetCumpriuParticipacao = (percParticipacao >= 7.5)

End Function

Private Sub Editar()

    With Grid
        
        If .Rows > 1 Then
            
            frmCadastro.txtCodigoProfessor.Text = .TextMatrix(.Row, 0)
            frmCadastro.txtNomeProfessor.Text = .TextMatrix(.Row, 1)
            frmCadastro.txtArea.Text = .TextMatrix(.Row, 2)
            frmCadastro.txtCodigoMateria.Text = .TextMatrix(.Row, 3)
            frmCadastro.txtDescricaoMateria.Text = .TextMatrix(.Row, 4)
            frmCadastro.txtTotalHora.Text = .TextMatrix(.Row, 5)
            frmCadastro.txtRAAluno.Text = .TextMatrix(.Row, 6)
            frmCadastro.txtCodigoAluno.Text = .TextMatrix(.Row, 7)
            frmCadastro.txtNomeAluno.Text = .TextMatrix(.Row, 8)
            frmCadastro.txtNota.Text = .TextMatrix(.Row, 9)
            frmCadastro.txtHorasAproveitadas.Text = .TextMatrix(.Row, 10)
            frmCadastro.varSequencia = .TextMatrix(.Row, 12)
               
            frmCadastro.Show vbModal
        End If
        
    End With

End Sub

Private Sub ExcluirNota()
    
    With Grid
        
        If .Rows > 1 Then
            If MensagemQuestionario("Deseja realmente excluir essa nota?") Then
                cn.Execute ("delete from notas where sequencia = " & .TextMatrix(.Row, 12) & "")
            End If
        End If
        
    End With
    
End Sub

Private Sub CarregaCombo()

    With cboCriterio
        
        .Clear
        .AddItem "Código Aluno", 0
        .AddItem "Nome Aluno", 1
        .ListIndex = 0
    
    End With

End Sub

Private Sub txtPesquisar_Change()
    CarregaGrid
End Sub

Private Sub txtPesquisar_LostFocus()
    CarregaGrid
End Sub
