VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9B3742B3-4639-4F01-8EFC-BEF3E8D567DE}#1.0#0"; "Panel2.ocx"
Begin VB.Form frmAguarde 
   BorderStyle     =   0  'None
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9525
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
   ScaleHeight     =   6900
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PanelFx32.PanelFx PainelCarregando 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2355
      TitleCaption    =   "Titulo..."
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
      Begin MSComctlLib.ProgressBar Bar 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Mensagem..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2850
         TabIndex        =   2
         Top             =   480
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mvarFormOrigem As Object
Public mvarTitulo As String
Public mvarMessagem As String
Public mvarValorMinimoBar As Long
Public mvarValorMaximoBar As Long
Public mvarValorAtualBar As Long

Private Sub Form_Load()
    
    With PainelCarregando
    
        .Top = Me.Top
        .Left = Me.Left
        
        Me.Width = .Width
        Me.Height = .Height
    
    End With
    
End Sub

Public Sub Load(cFormOrigem, cTitulo As String, cMensagem As String, cValorMinimoBar As Long, cValorMaximoBar As Long)

    Set mvarFormOrigem = cFormOrigem
    mvarTitulo = cTitulo
    mvarMessagem = cMensagem
    mvarValorMinimoBar = cValorMinimoBar
    mvarValorMaximoBar = cValorMaximoBar
    mvarValorAtualBar = 0
    
    PainelCarregando.TitleCaption = mvarTitulo
    lblMensagem.Caption = mvarMessagem
    
    Bar.Min = mvarValorMinimoBar
    Bar.Max = mvarValorMaximoBar
    Bar.value = mvarValorAtualBar
    
    Show
    
End Sub

Public Sub AtualizaProgresso(cMensagem As String, cValorAtual As Long)
    
    mvarMessagem = cMensagem
    mvarValorAtualBar = cValorAtual
    
    lblMensagem.Caption = mvarMessagem
    Bar.value = cValorAtual
    DoEvents
    
End Sub

Public Sub Fechar()
    Unload Me
End Sub

