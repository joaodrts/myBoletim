VERSION 5.00
Object = "{9B3742B3-4639-4F01-8EFC-BEF3E8D567DE}#1.0#0"; "Panel2.ocx"
Begin VB.Form frmCadastro 
   BorderStyle     =   0  'None
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PanelFx32.PanelFx PainelCad 
      Height          =   5175
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9128
      TitleCaption    =   "Cadastro"
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
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Height          =   735
         Left            =   3000
         TabIndex        =   11
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   4560
         TabIndex        =   12
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Aluno"
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   5895
         Begin VB.TextBox txtRAAluno 
            Height          =   315
            Left            =   4080
            TabIndex        =   8
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtNomeAluno 
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtHorasAproveitadas 
            Height          =   315
            Left            =   3120
            TabIndex        =   10
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox txtNota 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox txtCodigoAluno 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Horas Aproveitadas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3120
            TabIndex        =   27
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Nota"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "RA"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4080
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Código"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Materia"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   5895
         Begin VB.TextBox txtTotalHora 
            Height          =   315
            Left            =   4080
            TabIndex        =   5
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtDescricaoMateria 
            Height          =   315
            Left            =   1200
            TabIndex        =   4
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtCodigoMateria 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Código"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Descrição"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Total Horas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4080
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraProfessor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Professor"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5895
         Begin VB.TextBox txtArea 
            Height          =   315
            Left            =   4080
            TabIndex        =   2
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtNomeProfessor 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtCodigoProfessor 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Area"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4080
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Código"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varSequencia As Long

Private Sub cmdCancelar_Click()
    If MensagemQuestionario("Deseja realmente cancelar?") Then
        Unload Me
    End If
End Sub

Private Sub cmdGravar_Click()

    If TodosCamposPreenchidos Then
        Gravar
    End If

End Sub

Private Sub Form_Load()

    With PainelCad
    
        .Top = 0
        .Left = 0
        
        Me.Width = .Width
        Me.Height = .Height
    
    End With

End Sub

Private Sub PainelCad_TileClick()
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Function TodosCamposPreenchidos() As Boolean

    If txtCodigoProfessor.Text = "" Then MensagemErro "Informe o Código do Professor!": txtCodigoProfessor.SetFocus: Exit Function
    If txtNomeProfessor.Text = "" Then MensagemErro "Informe o Nome do Professor!": txtNomeProfessor.SetFocus: Exit Function
    If txtArea.Text = "" Then MensagemErro "Informe a Area do Professor!": txtArea.SetFocus: Exit Function
    If txtCodigoMateria.Text = "" Then MensagemErro "Informe o Código da Materia!": txtCodigoMateria.SetFocus: Exit Function
    If txtDescricaoMateria.Text = "" Then MensagemErro "Informe a Descrição da Materia!": txtDescricaoMateria.SetFocus: Exit Function
    If txtTotalHora.Text = "" Then MensagemErro "Informe o Total da Hora!": txtTotalHora.SetFocus: Exit Function
    If txtCodigoAluno.Text = "" Then MensagemErro "Informe o Código do aluno!": txtCodigoAluno.SetFocus: Exit Function
    If txtNomeAluno.Text = "" Then MensagemErro "Informe o Nome do aluno!": txtNomeAluno.SetFocus: Exit Function
    If txtRAAluno.Text = "" Then MensagemErro "Informe o RA do aluno!": txtRAAluno.SetFocus: Exit Function
    If txtNota.Text = "" Then MensagemErro "Informe a Nota do aluno!": txtNota.SetFocus: Exit Function
    If txtHorasAproveitadas.Text = "" Then MensagemErro "Informe as Horas Aproveitadas do aluno!": txtHorasAproveitadas.SetFocus: Exit Function
    
    TodosCamposPreenchidos = True

End Function

Private Sub Gravar()

    Dim sql As String
    
    If Not ValoresValido Then Exit Sub
    
    If Not IsNumeric(varSequencia) Then varSequencia = 0
            
    If NaoImportado(CLng(txtCodigoProfessor.Text), "professores") Then
        sql = "INSERT INTO professores values(" & txtCodigoProfessor.Text & ","
        sql = sql & "'" & ReplaceText(txtNomeProfessor.Text, "") & "',"
        sql = sql & "'" & ReplaceText(txtArea.Text, "") & "')"
        cn.Execute (sql)
    Else
        sql = "update professores set "
        sql = sql & "nome = '" & ReplaceText(txtNomeProfessor.Text, "") & "',"
        sql = sql & "area = '" & ReplaceText(txtArea.Text, "") & "' "
        sql = sql & "where codigo = " & txtCodigoProfessor.Text & ""
        cn.Execute (sql)
    End If

    If NaoImportado(CLng(txtCodigoMateria.Text), "materias") Then
        sql = "INSERT INTO materias values(" & txtCodigoMateria.Text & ","
        sql = sql & "'" & ReplaceText(txtDescricaoMateria.Text, "") & "',"
        sql = sql & "" & ReplaceText(txtTotalHora.Text, "") & ","
        sql = sql & "" & ReplaceText(txtCodigoProfessor.Text, "") & ")"
        cn.Execute (sql)
    Else
        sql = "update materias set "
        sql = sql & "nomemateria = '" & ReplaceText(txtDescricaoMateria.Text, "") & "',"
        sql = sql & "totalhoras = " & ReplaceText(txtTotalHora.Text, "") & ","
        sql = sql & "codigoprofessor = " & ReplaceText(txtCodigoProfessor.Text, "") & " "
        sql = sql & "where codigo = " & txtCodigoMateria.Text & ""
        cn.Execute (sql)
    End If

    If NaoImportado(CLng(txtCodigoAluno.Text), "alunos") Then
        sql = "INSERT INTO alunos values(" & txtCodigoAluno.Text & ","
        sql = sql & "" & ReplaceText(txtRAAluno.Text, "") & ","
        sql = sql & "'" & ReplaceText(txtNomeAluno.Text, "") & "')"
        cn.Execute (sql)
    Else
        sql = "update alunos set "
        sql = sql & "ra_aluno = " & ReplaceText(txtRAAluno.Text, "") & ","
        sql = sql & "nomealuno = '" & ReplaceText(txtNomeAluno.Text, "") & "' "
        sql = sql & "where codigo = " & txtCodigoAluno.Text & ""
        cn.Execute (sql)
    End If
    
    If NaoImportado(CLng(varSequencia), "notas") Then
        sql = "INSERT INTO notas(codigoaluno, nota, horasaproveitadas, codigomateria) values(" & txtCodigoAluno.Text & ","
        sql = sql & "" & ReplaceText(txtNota.Text, "") & ","
        sql = sql & "" & ReplaceText(txtHorasAproveitadas.Text, "") & ","
        sql = sql & "" & ReplaceText(txtCodigoMateria.Text, "") & ")"
        cn.Execute (sql)
    Else
        sql = "update notas set codigoaluno = " & txtCodigoAluno.Text & ","
        sql = sql & "nota = " & ReplaceText(txtNota.Text, "") & ","
        sql = sql & "horasaproveitadas = " & ReplaceText(txtHorasAproveitadas.Text, "") & ","
        sql = sql & "codigomateria = " & ReplaceText(txtCodigoMateria.Text, "") & " "
        sql = sql & "where sequencia = " & varSequencia & ""
        cn.Execute (sql)
    End If

    MensagemInformativa "Registro cadastrado com sucesso!"
    Unload Me

End Sub

Private Function ValoresValido() As Boolean
    
    ValoresValido = True
    
    If Not IsNumeric(txtCodigoProfessor.Text) Then
        ValoresValido = False
        MensagemErro "Informe um código válido para o professor!"
        txtCodigoProfessor.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtCodigoMateria.Text) Then
        ValoresValido = False
        MensagemErro "Informe um código válido para a materia!"
        txtCodigoMateria.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtCodigoAluno.Text) Then
        ValoresValido = False
        MensagemErro "Informe um código válido para o aluno!"
        txtCodigoAluno.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtTotalHora.Text) Then
        ValoresValido = False
        MensagemErro "Informe uma hora válida!"
        txtTotalHora.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtHorasAproveitadas.Text) Then
        ValoresValido = False
        MensagemErro "Informe uma hora válida!"
        txtHorasAproveitadas.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtRAAluno.Text) Then
        ValoresValido = False
        MensagemErro "Informe um código válido!"
        txtRAAluno.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtNota.Text) Then
        ValoresValido = False
        MensagemErro "Informe uma nota para o aluno!"
        txtNota.SetFocus
        Exit Function
    End If
    
    If Len(Replace(txtNomeProfessor.Text, "'", "")) = 0 Then
        ValoresValido = False
        MensagemErro "Informe um nome para o professor!"
        txtNomeProfessor.SetFocus
        Exit Function
    End If
    
    If Len(Replace(txtNomeProfessor.Text, "'", "")) = 0 Then
        ValoresValido = False
        MensagemErro "Informe um nome para o professor!"
        txtNomeProfessor.SetFocus
        Exit Function
    End If
    
    If Len(Replace(txtArea.Text, "'", "")) = 0 Then
        ValoresValido = False
        MensagemErro "Informe uma area para o professor!"
        txtArea.SetFocus
        Exit Function
    End If
    
    If Len(Replace(txtNomeAluno.Text, "'", "")) = 0 Then
        ValoresValido = False
        MensagemErro "Informe um nome para o aluno!"
        txtNomeAluno.SetFocus
        Exit Function
    End If

End Function
