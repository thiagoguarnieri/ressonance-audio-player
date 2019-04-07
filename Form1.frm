VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ressonance"
   ClientHeight    =   10815
   ClientLeft      =   1860
   ClientTop       =   2415
   ClientWidth     =   15270
   FillColor       =   &H80000006&
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   10815
   ScaleWidth      =   15270
   Begin TabDlg.SSTab SSTab2 
      Height          =   4005
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "Gerencia Blocos, mostra eventos e gerencia cartucheiras."
      Top             =   6510
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   723
      BackColor       =   -2147483644
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Playlist"
      TabPicture(0)   =   "Form1.frx":538C4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelCalcBloco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPausa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "autocross"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "autoload"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "armar2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "deletaruma"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "player2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "player1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "salvarlista"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "apagarlista"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "abrirlista"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdHoraCerta"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "playlistView"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Eventos Executados"
      TabPicture(1)   =   "Form1.frx":538E0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tocadaslist"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cartucheiras Abertas"
      TabPicture(2)   =   "Form1.frx":538FC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstStreamAberto"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Log"
      TabPicture(3)   =   "Form1.frx":53918
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "textoLog"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Botoneira 1"
      TabPicture(4)   =   "Form1.frx":53934
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdAbrirVinSup"
      Tab(4).Control(1)=   "cmdSvVinSup"
      Tab(4).Control(2)=   "cmdLimparSup"
      Tab(4).Control(3)=   "vinheta1(31)"
      Tab(4).Control(4)=   "vinheta1(32)"
      Tab(4).Control(5)=   "vinheta1(33)"
      Tab(4).Control(6)=   "vinheta1(34)"
      Tab(4).Control(7)=   "vinheta1(35)"
      Tab(4).Control(8)=   "vinheta1(29)"
      Tab(4).Control(9)=   "vinheta1(30)"
      Tab(4).Control(10)=   "vinheta1(24)"
      Tab(4).Control(11)=   "vinheta1(25)"
      Tab(4).Control(12)=   "vinheta1(26)"
      Tab(4).Control(13)=   "vinheta1(27)"
      Tab(4).Control(14)=   "vinheta1(28)"
      Tab(4).Control(15)=   "lblVinInfo(31)"
      Tab(4).Control(16)=   "lblVinInfo(32)"
      Tab(4).Control(17)=   "lblVinInfo(33)"
      Tab(4).Control(18)=   "lblVinInfo(34)"
      Tab(4).Control(19)=   "lblVinInfo(35)"
      Tab(4).Control(20)=   "lblVinInfo(29)"
      Tab(4).Control(21)=   "lblVinInfo(30)"
      Tab(4).Control(22)=   "lblVinInfo(24)"
      Tab(4).Control(23)=   "lblVinInfo(25)"
      Tab(4).Control(24)=   "lblVinInfo(26)"
      Tab(4).Control(25)=   "lblVinInfo(27)"
      Tab(4).Control(26)=   "lblVinInfo(28)"
      Tab(4).ControlCount=   27
      Begin MSComctlLib.ListView playlistView 
         Height          =   2115
         Left            =   105
         TabIndex        =   50
         Top             =   1365
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ColHdrIcons     =   "listaImagens1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Playlist de áudios"
            Object.Width           =   14111
            ImageIndex      =   2
         EndProperty
      End
      Begin VB.CommandButton cmdAbrirVinSup 
         Caption         =   "Abrir"
         Height          =   345
         Left            =   -74790
         TabIndex        =   222
         ToolTipText     =   "Abre arquivo com as vinhetas programadas"
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton cmdSvVinSup 
         Caption         =   "Salvar"
         Height          =   345
         Left            =   -73425
         TabIndex        =   221
         ToolTipText     =   "Salva arquivo com as vinhetas programadas"
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton cmdLimparSup 
         Caption         =   "Limpar"
         Height          =   345
         Left            =   -72060
         TabIndex        =   220
         ToolTipText     =   "Retorna todos os slots azuis a condição de vazios"
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   31
         Left            =   -73425
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   214
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   32
         Left            =   -72165
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   213
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   33
         Left            =   -70905
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   212
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   34
         Left            =   -69645
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   211
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2310
         Width           =   1380
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   35
         Left            =   -68280
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   210
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2310
         Width           =   1380
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   29
         Left            =   -68280
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   207
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1155
         Width           =   1380
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   30
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   206
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2310
         Width           =   1380
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   24
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   200
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1155
         Width           =   1380
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   25
         Left            =   -73425
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   199
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1155
         Width           =   1275
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   26
         Left            =   -72165
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   198
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1155
         Width           =   1275
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   27
         Left            =   -70905
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1155
         Width           =   1275
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   28
         Left            =   -69645
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   196
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1155
         Width           =   1380
      End
      Begin VB.CommandButton cmdHoraCerta 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Hora Certa"
         Height          =   330
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   945
         Width           =   1065
      End
      Begin VB.TextBox textoLog 
         Height          =   3420
         Left            =   -74925
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   154
         Text            =   "Form1.frx":53950
         ToolTipText     =   "Log de carregamento dos arquivos de configuração"
         Top             =   450
         Width           =   8115
      End
      Begin VB.ListBox tocadaslist 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74955
         TabIndex        =   52
         ToolTipText     =   "Mostra os eventos executados durante certo período de tempo"
         Top             =   450
         Width           =   8160
      End
      Begin VB.CommandButton abrirlista 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Abrir Playlist"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "abrir programação (playlist)"
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton apagarlista 
         BackColor       =   &H000000FF&
         Caption         =   "Apagar Tudo"
         Height          =   330
         Left            =   4095
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "deletar todos os áudios da playlist"
         Top             =   945
         Width           =   1275
      End
      Begin VB.CommandButton salvarlista 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Salvar Playlist"
         Height          =   330
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "salvar programação (playlist)"
         Top             =   525
         Width           =   1275
      End
      Begin VB.CommandButton player1 
         BackColor       =   &H00FF8080&
         Caption         =   "1"
         Height          =   330
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "carregar na cartucheira 1"
         Top             =   525
         Width           =   645
      End
      Begin VB.CommandButton player2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
         Height          =   330
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "carregar no cartucheira 2"
         Top             =   945
         Width           =   645
      End
      Begin VB.CommandButton deletaruma 
         BackColor       =   &H008080FF&
         Caption         =   "Apagar Item"
         Height          =   330
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "deletar áudio selecionado"
         Top             =   945
         Width           =   1065
      End
      Begin VB.CommandButton armar2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cartucheira auxiliar"
         Height          =   750
         Left            =   6510
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "abrir em nova cartucheira"
         Top             =   525
         Width           =   960
      End
      Begin VB.CheckBox autoload 
         BackColor       =   &H000000FF&
         Caption         =   "Piloto Automático"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   42
         ToolTipText     =   "Serve para executar automaticamente um áudio"
         Top             =   3570
         Width           =   1590
      End
      Begin VB.CheckBox autocross 
         BackColor       =   &H000000FF&
         Caption         =   "Mixagem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5565
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   41
         ToolTipText     =   "Utilizado para Intercalar audios entre cart. 1 e 2"
         Top             =   3570
         Width           =   1065
      End
      Begin VB.CommandButton cmdPausa 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pausa"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Utilizado para o operador interromper o AUTOPLAY"
         Top             =   945
         Width           =   750
      End
      Begin MSComctlLib.ListView lstStreamAberto 
         Height          =   3495
         Left            =   -74955
         TabIndex        =   53
         ToolTipText     =   "Para fazer aparecer uma cartucheira minimizada na barra de tarefas, basta dar dois cliques num item da lista"
         Top             =   420
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "listaImagens1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Streams"
            Object.Width           =   11290
            ImageIndex      =   1
         EndProperty
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   31
         Left            =   -73425
         TabIndex        =   219
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   32
         Left            =   -72165
         TabIndex        =   218
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   33
         Left            =   -70905
         TabIndex        =   217
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   34
         Left            =   -69645
         TabIndex        =   216
         Top             =   3150
         Width           =   1380
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   35
         Left            =   -68280
         TabIndex        =   215
         Top             =   3150
         Width           =   1380
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   29
         Left            =   -68280
         TabIndex        =   209
         Top             =   1995
         Width           =   1380
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   30
         Left            =   -74790
         TabIndex        =   208
         Top             =   3150
         Width           =   1380
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   24
         Left            =   -74790
         TabIndex        =   205
         Top             =   1995
         Width           =   1380
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   25
         Left            =   -73425
         TabIndex        =   204
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   26
         Left            =   -72165
         TabIndex        =   203
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   27
         Left            =   -70905
         TabIndex        =   202
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   28
         Left            =   -69645
         TabIndex        =   201
         Top             =   1995
         Width           =   1380
      End
      Begin VB.Label LabelCalcBloco 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6825
         TabIndex        =   51
         ToolTipText     =   "Calculo do bloco"
         Top             =   3570
         Width           =   1380
      End
   End
   Begin VB.CommandButton avisovisto 
      Caption         =   "OK"
      Height          =   540
      Left            =   14595
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   5670
      Width           =   540
   End
   Begin VB.PictureBox vuMeter 
      BackColor       =   &H00000000&
      Height          =   960
      Left            =   9345
      ScaleHeight     =   900
      ScaleWidth      =   5625
      TabIndex        =   195
      Top             =   1785
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busca de áudios"
      Height          =   3270
      Left            =   9240
      TabIndex        =   183
      Top             =   7140
      Width           =   5895
      Begin VB.CommandButton cmdCart1busca 
         BackColor       =   &H00FF8080&
         Caption         =   "1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   193
         ToolTipText     =   "carregar na cartucheira 1"
         Top             =   1050
         Width           =   765
      End
      Begin VB.CommandButton cmdCart2busca 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   192
         ToolTipText     =   "carregar na cartucheira 2"
         Top             =   1050
         Width           =   690
      End
      Begin VB.CommandButton cmdToPlaylist 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Playlist"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   191
         ToolTipText     =   "Envia áudio selecionado para a playlist"
         Top             =   1050
         Width           =   1440
      End
      Begin VB.CommandButton cmdPastaBase 
         Height          =   330
         Left            =   2205
         Picture         =   "Form1.frx":53964
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   630
         Width           =   435
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4935
         TabIndex        =   185
         Top             =   210
         Width           =   750
      End
      Begin VB.TextBox txtBusca 
         Height          =   330
         Left            =   210
         TabIndex        =   184
         Top             =   210
         Width           =   4635
      End
      Begin MSComctlLib.ListView ListViewBusca 
         Height          =   1740
         Left            =   105
         TabIndex        =   194
         Top             =   1470
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   3069
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Eventos encontrados"
            Object.Width           =   12736
         EndProperty
      End
      Begin VB.Label lblPastaBase 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Selecione a pasta de busca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   225
         Top             =   630
         Width           =   2955
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStatusBusca 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         TabIndex        =   187
         Top             =   630
         Width           =   1905
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cartucheiras Principais"
      Height          =   4845
      Left            =   9240
      TabIndex        =   155
      Top             =   105
      Width           =   5895
      Begin VB.CommandButton play1 
         Height          =   465
         Left            =   330
         Picture         =   "Form1.frx":53A26
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton pause1 
         Height          =   465
         Left            =   960
         Picture         =   "Form1.frx":5626A
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton stop1 
         Height          =   465
         Left            =   1590
         Picture         =   "Form1.frx":58A03
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton bw1 
         Height          =   465
         Left            =   2205
         Picture         =   "Form1.frx":5B158
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   315
         Width           =   555
      End
      Begin VB.CommandButton fw1 
         Height          =   465
         Left            =   2835
         Picture         =   "Form1.frx":5D99E
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton play2 
         Height          =   540
         Left            =   315
         Picture         =   "Form1.frx":601F9
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   2715
         Width           =   540
      End
      Begin VB.CommandButton pause2 
         Height          =   540
         Left            =   930
         Picture         =   "Form1.frx":62A3D
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   2715
         Width           =   540
      End
      Begin VB.CommandButton stop2 
         Height          =   540
         Left            =   1560
         Picture         =   "Form1.frx":651D6
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   2715
         Width           =   540
      End
      Begin VB.CommandButton bw2 
         Height          =   540
         Left            =   2190
         Picture         =   "Form1.frx":6792B
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   2715
         Width           =   540
      End
      Begin VB.CommandButton fw2 
         Height          =   540
         Left            =   2820
         Picture         =   "Form1.frx":6A171
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   2715
         Width           =   540
      End
      Begin VB.CommandButton cmdFade1 
         Caption         =   "Fade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4620
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Pára a música em volume decrescente"
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton cmdFade2 
         Caption         =   "Fade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Pára a música em volume decrescente"
         Top             =   3255
         Width           =   1065
      End
      Begin MSComCtl2.FlatScrollBar ScrollVol2 
         Height          =   330
         Left            =   3675
         TabIndex        =   168
         Top             =   3675
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
         _Version        =   393216
         Arrows          =   65536
         Max             =   10
         Orientation     =   1245185
         Value           =   10
      End
      Begin MSComCtl2.FlatScrollBar ScrollVol1 
         Height          =   330
         Left            =   3675
         TabIndex        =   169
         Top             =   1260
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
         _Version        =   393216
         Arrows          =   65536
         Max             =   10
         Orientation     =   1245185
         Value           =   10
      End
      Begin VB.Label LabelPosition1 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   840
         TabIndex        =   223
         Top             =   840
         Width           =   645
      End
      Begin VB.Label infomp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CARTUCHEIRA 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   210
         TabIndex        =   181
         Top             =   3570
         Width           =   3165
         WordWrap        =   -1  'True
      End
      Begin VB.Label Labelduracao2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2205
         TabIndex        =   180
         Top             =   3255
         Width           =   750
      End
      Begin VB.Label LabelPosition2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   840
         TabIndex        =   179
         Top             =   3255
         Width           =   645
      End
      Begin VB.Label LabelMode2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   4200
         TabIndex        =   178
         Top             =   2730
         Width           =   1485
      End
      Begin VB.Label labelRest2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1575
         TabIndex        =   177
         Top             =   3255
         Width           =   645
      End
      Begin VB.Label lblhora 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   435
         Left            =   105
         TabIndex        =   176
         Top             =   4305
         Width           =   5685
      End
      Begin VB.Label labelRest1 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1575
         TabIndex        =   175
         Top             =   840
         Width           =   645
      End
      Begin VB.Label infomp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CARTUCHEIRA 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   315
         TabIndex        =   174
         Top             =   1155
         Width           =   3165
         WordWrap        =   -1  'True
      End
      Begin VB.Label LabelMode1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   4200
         TabIndex        =   173
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label lblled1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3675
         TabIndex        =   172
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblled2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3675
         TabIndex        =   171
         Top             =   3255
         Width           =   855
      End
      Begin VB.Label LabelDuracao1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2205
         TabIndex        =   170
         Top             =   840
         Width           =   645
      End
      Begin VB.Image Image1 
         DragMode        =   1  'Automatic
         Height          =   1485
         Left            =   105
         Picture         =   "Form1.frx":6C9CC
         Stretch         =   -1  'True
         Top             =   210
         Width           =   5685
      End
      Begin VB.Image Image2 
         Height          =   1485
         Left            =   105
         Picture         =   "Form1.frx":7092C
         Stretch         =   -1  'True
         Top             =   2625
         Width           =   5685
      End
   End
   Begin VB.Timer tmrAvisoPiscante1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   8115
   End
   Begin VB.Timer tmrAvisoPiscante2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3330
      Top             =   8115
   End
   Begin VB.Timer TimerFade1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6180
      Top             =   8115
   End
   Begin VB.Timer TimerFade2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5280
      Top             =   8115
   End
   Begin MSComctlLib.ImageList listaImagens1 
      Left            =   6720
      Top             =   8085
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":74CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7508A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer timerAutocross 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5730
      Top             =   8115
   End
   Begin VB.Timer TimerProgressao 
      Interval        =   1000
      Left            =   4830
      Top             =   8115
   End
   Begin VB.Timer TimerEventos 
      Interval        =   1000
      Left            =   4305
      Top             =   8115
   End
   Begin VB.Timer tmrSpectrum 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3900
      Top             =   8115
   End
   Begin MSComDlg.CommonDialog dialogo1 
      Left            =   7350
      Top             =   8085
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton uplist 
      Height          =   855
      Left            =   8505
      Picture         =   "Form1.frx":75424
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "mover prara cima"
      Top             =   7455
      Width           =   645
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   10560
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton downlist 
      Height          =   855
      Left            =   8505
      Picture         =   "Form1.frx":78083
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "mover para baixo"
      Top             =   8400
      Width           =   645
   End
   Begin TabDlg.SSTab tab1 
      Height          =   6315
      Left            =   105
      TabIndex        =   4
      ToolTipText     =   "Insere áudios, gerencia comerciais, carrega vinhetas e carrega a partir de máscaras."
      Top             =   105
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   723
      BackColor       =   -2147483644
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Eventos"
      TabPicture(0)   =   "Form1.frx":7ADC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmArquivos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Evt. Programados"
      TabPicture(1)   =   "Form1.frx":7ADDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "arvoreComercial"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Lembretes"
      TabPicture(2)   =   "Form1.frx":7ADFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "caixaderecados"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Botoneira 2"
      TabPicture(3)   =   "Form1.frx":7AE16
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "vinheta1(23)"
      Tab(3).Control(1)=   "vinheta1(22)"
      Tab(3).Control(2)=   "vinheta1(21)"
      Tab(3).Control(3)=   "vinheta1(20)"
      Tab(3).Control(4)=   "vinheta1(19)"
      Tab(3).Control(5)=   "vinheta1(18)"
      Tab(3).Control(6)=   "vinheta1(17)"
      Tab(3).Control(7)=   "vinheta1(16)"
      Tab(3).Control(8)=   "vinheta1(11)"
      Tab(3).Control(9)=   "vinheta1(10)"
      Tab(3).Control(10)=   "vinheta1(9)"
      Tab(3).Control(11)=   "vinheta1(8)"
      Tab(3).Control(12)=   "vinheta1(7)"
      Tab(3).Control(13)=   "vinheta1(6)"
      Tab(3).Control(14)=   "vinheta1(5)"
      Tab(3).Control(15)=   "vinheta1(4)"
      Tab(3).Control(16)=   "vinheta1(3)"
      Tab(3).Control(17)=   "vinheta1(2)"
      Tab(3).Control(18)=   "vinheta1(1)"
      Tab(3).Control(19)=   "vinheta1(0)"
      Tab(3).Control(20)=   "vinheta1(15)"
      Tab(3).Control(21)=   "vinheta1(14)"
      Tab(3).Control(22)=   "vinheta1(13)"
      Tab(3).Control(23)=   "vinheta1(12)"
      Tab(3).Control(24)=   "lblVinInfo(23)"
      Tab(3).Control(25)=   "lblVinInfo(22)"
      Tab(3).Control(26)=   "lblVinInfo(21)"
      Tab(3).Control(27)=   "lblVinInfo(20)"
      Tab(3).Control(28)=   "lblVinInfo(19)"
      Tab(3).Control(29)=   "lblVinInfo(18)"
      Tab(3).Control(30)=   "lblVinInfo(17)"
      Tab(3).Control(31)=   "lblVinInfo(16)"
      Tab(3).Control(32)=   "lblVinInfo(15)"
      Tab(3).Control(33)=   "lblVinInfo(14)"
      Tab(3).Control(34)=   "lblVinInfo(13)"
      Tab(3).Control(35)=   "lblVinInfo(12)"
      Tab(3).Control(36)=   "lblVinInfo(11)"
      Tab(3).Control(37)=   "lblVinInfo(10)"
      Tab(3).Control(38)=   "lblVinInfo(9)"
      Tab(3).Control(39)=   "lblVinInfo(8)"
      Tab(3).Control(40)=   "lblVinInfo(7)"
      Tab(3).Control(41)=   "lblVinInfo(6)"
      Tab(3).Control(42)=   "lblVinInfo(5)"
      Tab(3).Control(43)=   "lblVinInfo(4)"
      Tab(3).Control(44)=   "lblVinInfo(3)"
      Tab(3).Control(45)=   "lblVinInfo(2)"
      Tab(3).Control(46)=   "lblVinInfo(1)"
      Tab(3).Control(47)=   "lblVinInfo(0)"
      Tab(3).ControlCount=   48
      TabCaption(4)   =   "Botoneira 3"
      TabPicture(4)   =   "Form1.frx":7AE32
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "vinheta1(36)"
      Tab(4).Control(1)=   "vinheta1(37)"
      Tab(4).Control(2)=   "vinheta1(38)"
      Tab(4).Control(3)=   "vinheta1(39)"
      Tab(4).Control(4)=   "vinheta1(40)"
      Tab(4).Control(5)=   "vinheta1(41)"
      Tab(4).Control(6)=   "vinheta1(42)"
      Tab(4).Control(7)=   "vinheta1(43)"
      Tab(4).Control(8)=   "vinheta1(44)"
      Tab(4).Control(9)=   "vinheta1(45)"
      Tab(4).Control(10)=   "vinheta1(46)"
      Tab(4).Control(11)=   "vinheta1(47)"
      Tab(4).Control(12)=   "vinheta1(48)"
      Tab(4).Control(13)=   "vinheta1(49)"
      Tab(4).Control(14)=   "vinheta1(50)"
      Tab(4).Control(15)=   "vinheta1(51)"
      Tab(4).Control(16)=   "vinheta1(52)"
      Tab(4).Control(17)=   "vinheta1(53)"
      Tab(4).Control(18)=   "vinheta1(54)"
      Tab(4).Control(19)=   "vinheta1(55)"
      Tab(4).Control(20)=   "vinheta1(56)"
      Tab(4).Control(21)=   "vinheta1(57)"
      Tab(4).Control(22)=   "vinheta1(58)"
      Tab(4).Control(23)=   "vinheta1(59)"
      Tab(4).Control(24)=   "lblVinInfo(36)"
      Tab(4).Control(25)=   "lblVinInfo(37)"
      Tab(4).Control(26)=   "lblVinInfo(38)"
      Tab(4).Control(27)=   "lblVinInfo(39)"
      Tab(4).Control(28)=   "lblVinInfo(40)"
      Tab(4).Control(29)=   "lblVinInfo(41)"
      Tab(4).Control(30)=   "lblVinInfo(42)"
      Tab(4).Control(31)=   "lblVinInfo(43)"
      Tab(4).Control(32)=   "lblVinInfo(44)"
      Tab(4).Control(33)=   "lblVinInfo(45)"
      Tab(4).Control(34)=   "lblVinInfo(46)"
      Tab(4).Control(35)=   "lblVinInfo(47)"
      Tab(4).Control(36)=   "lblVinInfo(48)"
      Tab(4).Control(37)=   "lblVinInfo(49)"
      Tab(4).Control(38)=   "lblVinInfo(50)"
      Tab(4).Control(39)=   "lblVinInfo(51)"
      Tab(4).Control(40)=   "lblVinInfo(52)"
      Tab(4).Control(41)=   "lblVinInfo(53)"
      Tab(4).Control(42)=   "lblVinInfo(54)"
      Tab(4).Control(43)=   "lblVinInfo(55)"
      Tab(4).Control(44)=   "lblVinInfo(56)"
      Tab(4).Control(45)=   "lblVinInfo(57)"
      Tab(4).Control(46)=   "lblVinInfo(58)"
      Tab(4).Control(47)=   "lblVinInfo(59)"
      Tab(4).ControlCount=   48
      TabCaption(5)   =   "Botoneira Rotativa"
      TabPicture(5)   =   "Form1.frx":7AE4E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdVinRotativas(0)"
      Tab(5).Control(1)=   "cmdVinRotativas(5)"
      Tab(5).Control(2)=   "cmdVinRotativas(4)"
      Tab(5).Control(3)=   "cmdVinRotativas(3)"
      Tab(5).Control(4)=   "cmdVinRotativas(2)"
      Tab(5).Control(5)=   "cmdVinRotativas(1)"
      Tab(5).Control(6)=   "ListaRotativas(0)"
      Tab(5).Control(7)=   "ListaRotativas(1)"
      Tab(5).Control(8)=   "ListaRotativas(2)"
      Tab(5).Control(9)=   "ListaRotativas(3)"
      Tab(5).Control(10)=   "ListaRotativas(4)"
      Tab(5).Control(11)=   "ListaRotativas(5)"
      Tab(5).Control(12)=   "Labelrotativas(5)"
      Tab(5).Control(13)=   "Labelrotativas(4)"
      Tab(5).Control(14)=   "Labelrotativas(3)"
      Tab(5).Control(15)=   "Labelrotativas(2)"
      Tab(5).Control(16)=   "Labelrotativas(1)"
      Tab(5).Control(17)=   "Labelrotativas(0)"
      Tab(5).ControlCount=   18
      TabCaption(6)   =   "Espelho"
      TabPicture(6)   =   "Form1.frx":7AE6A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "espelhoToPlaylist"
      Tab(6).Control(1)=   "cmdTreeCart1"
      Tab(6).Control(2)=   "cmdTreeCart2"
      Tab(6).Control(3)=   "streamEspelho"
      Tab(6).Control(4)=   "Espelho"
      Tab(6).ControlCount=   5
      Begin VB.CommandButton espelhoToPlaylist 
         Caption         =   "Playlist v v"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -71115
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "transfere múltiplos áudios"
         Top             =   5775
         Width           =   1065
      End
      Begin VB.CommandButton cmdTreeCart1 
         BackColor       =   &H00FF8080&
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -67230
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton cmdTreeCart2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -67230
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   1680
         Width           =   1170
      End
      Begin VB.CommandButton streamEspelho 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cartucheira auxiliar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   -67230
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   2205
         Width           =   1170
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   36
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   37
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   38
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   39
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   40
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   41
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   42
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   43
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   44
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   45
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   46
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   47
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   48
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   49
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   50
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   51
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   52
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   53
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   54
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   55
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   56
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   57
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   58
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   59
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton cmdVinRotativas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<vazio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   0
         Left            =   -74685
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   630
         Width           =   2745
      End
      Begin VB.CommandButton cmdVinRotativas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<vazio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   5
         Left            =   -69015
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3465
         Width           =   2745
      End
      Begin VB.CommandButton cmdVinRotativas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<vazio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   4
         Left            =   -71955
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3465
         Width           =   2955
      End
      Begin VB.CommandButton cmdVinRotativas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<vazio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   3
         Left            =   -74685
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3465
         Width           =   2745
      End
      Begin VB.CommandButton cmdVinRotativas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<vazio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   2
         Left            =   -69015
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   630
         Width           =   2745
      End
      Begin VB.CommandButton cmdVinRotativas 
         BackColor       =   &H00C0FFC0&
         Caption         =   "<vazio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   1
         Left            =   -71955
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   630
         Width           =   2955
      End
      Begin MSComctlLib.ListView ListaRotativas 
         Height          =   855
         Index           =   0
         Left            =   -74685
         TabIndex        =   82
         Top             =   1680
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1508
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "vinhetas"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.TreeView arvoreComercial 
         Height          =   4740
         Left            =   -74895
         TabIndex        =   80
         Top             =   1365
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   8361
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   23
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   22
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   21
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   20
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   5145
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   19
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   18
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   17
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   16
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   4095
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   11
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   10
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   9
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   8
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   2205
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   7
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   6
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   5
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   4
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   3
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   2
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   525
         Width           =   2115
      End
      Begin VB.Frame Frame4 
         Caption         =   "arquivos"
         Height          =   735
         Left            =   -74910
         TabIndex        =   13
         Top             =   495
         Width           =   3630
         Begin VB.CommandButton abrircomerciais 
            Height          =   375
            Left            =   120
            Picture         =   "Form1.frx":7AE86
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
         Begin VB.Label MonitorCom 
            Alignment       =   2  'Center
            Caption         =   "Em espera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1125
            TabIndex        =   15
            Top             =   225
            Width           =   2415
         End
      End
      Begin VB.TextBox caixaderecados 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5580
         Left            =   -74895
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Form1.frx":7D216
         Top             =   525
         Width           =   8730
      End
      Begin VB.Frame frmArquivos 
         Caption         =   "Explorer"
         ForeColor       =   &H80000007&
         Height          =   5835
         Left            =   105
         TabIndex        =   5
         Top             =   420
         Width           =   8880
         Begin VB.CommandButton cmdCue 
            Caption         =   "CUE"
            Height          =   645
            Left            =   7665
            Picture         =   "Form1.frx":7D237
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   3045
            Width           =   1065
         End
         Begin VB.CommandButton streamPrev 
            BackColor       =   &H0080C0FF&
            Caption         =   "Cartucheira auxiliar"
            Height          =   555
            Left            =   7665
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   1575
            Width           =   1110
         End
         Begin VB.DriveListBox DrivePrincipal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   11
            Top             =   315
            Width           =   3375
         End
         Begin VB.DirListBox DirPrincipal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1665
            Left            =   105
            OLEDragMode     =   1  'Automatic
            TabIndex        =   10
            Top             =   735
            Width           =   7365
         End
         Begin VB.FileListBox FilePrincipal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3210
            Left            =   105
            MultiSelect     =   2  'Extended
            Pattern         =   "*.mp3;*.wav"
            TabIndex        =   9
            Top             =   2520
            Width           =   7365
         End
         Begin VB.CommandButton trasfermulti 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Playlist"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   7665
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "transfere múltiplos áudios"
            Top             =   3780
            Width           =   1065
         End
         Begin VB.CommandButton cmdPreview1 
            BackColor       =   &H00FF8080&
            Caption         =   "1"
            Height          =   390
            Left            =   7665
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   525
            Width           =   1110
         End
         Begin VB.CommandButton cmdPreview2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "2"
            Height          =   390
            Left            =   7665
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1050
            Width           =   1110
         End
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   15
         Left            =   -68490
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   14
         Left            =   -70590
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   13
         Left            =   -72690
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin VB.CommandButton vinheta1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "vazio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   12
         Left            =   -74790
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Clique neste botão para abrir ou tocar um evento"
         Top             =   3150
         Width           =   2115
      End
      Begin MSComctlLib.ListView ListaRotativas 
         Height          =   855
         Index           =   1
         Left            =   -71955
         TabIndex        =   89
         Top             =   1680
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   1508
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "vinhetas"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView ListaRotativas 
         Height          =   855
         Index           =   2
         Left            =   -69015
         TabIndex        =   90
         Top             =   1680
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1508
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "vinhetas"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView ListaRotativas 
         Height          =   960
         Index           =   3
         Left            =   -74685
         TabIndex        =   91
         Top             =   4515
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1693
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "vinhetas"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView ListaRotativas 
         Height          =   960
         Index           =   4
         Left            =   -71955
         TabIndex        =   92
         Top             =   4515
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   1693
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "vinhetas"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView ListaRotativas 
         Height          =   960
         Index           =   5
         Left            =   -69015
         TabIndex        =   93
         Top             =   4515
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1693
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "vinhetas"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.TreeView Espelho 
         Height          =   4740
         Left            =   -74910
         TabIndex        =   153
         Top             =   915
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8361
         _Version        =   393217
         HideSelection   =   0   'False
         Style           =   7
         ImageList       =   "listaImagens1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   36
         Left            =   -68490
         TabIndex        =   148
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   37
         Left            =   -70590
         TabIndex        =   147
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   38
         Left            =   -72690
         TabIndex        =   146
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   39
         Left            =   -74790
         TabIndex        =   145
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   40
         Left            =   -68490
         TabIndex        =   144
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   41
         Left            =   -70590
         TabIndex        =   143
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   42
         Left            =   -72690
         TabIndex        =   142
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   43
         Left            =   -74790
         TabIndex        =   141
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   44
         Left            =   -68490
         TabIndex        =   140
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   45
         Left            =   -70590
         TabIndex        =   139
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   46
         Left            =   -72690
         TabIndex        =   138
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   47
         Left            =   -74790
         TabIndex        =   137
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   48
         Left            =   -68490
         TabIndex        =   136
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   49
         Left            =   -70590
         TabIndex        =   135
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   50
         Left            =   -72795
         TabIndex        =   134
         Top             =   1995
         Width           =   2220
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   51
         Left            =   -74790
         TabIndex        =   133
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   52
         Left            =   -68490
         TabIndex        =   132
         Top             =   1155
         Width           =   2130
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   53
         Left            =   -70590
         TabIndex        =   131
         Top             =   1155
         Width           =   2130
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   54
         Left            =   -72690
         TabIndex        =   130
         Top             =   1155
         Width           =   2130
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   55
         Left            =   -74790
         TabIndex        =   129
         Top             =   1155
         Width           =   2130
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   56
         Left            =   -68490
         TabIndex        =   128
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   57
         Left            =   -70590
         TabIndex        =   127
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   58
         Left            =   -72690
         TabIndex        =   126
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   59
         Left            =   -74790
         TabIndex        =   125
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label Labelrotativas 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Vinheta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   5
         Left            =   -69015
         TabIndex        =   99
         Top             =   5460
         Width           =   2745
      End
      Begin VB.Label Labelrotativas 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Vinheta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   4
         Left            =   -71955
         TabIndex        =   98
         Top             =   5460
         Width           =   2955
      End
      Begin VB.Label Labelrotativas 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Vinheta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   3
         Left            =   -74685
         TabIndex        =   97
         Top             =   5460
         Width           =   2745
      End
      Begin VB.Label Labelrotativas 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Vinheta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   2
         Left            =   -69015
         TabIndex        =   96
         Top             =   2520
         Width           =   2745
      End
      Begin VB.Label Labelrotativas 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Vinheta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   1
         Left            =   -71955
         TabIndex        =   95
         Top             =   2520
         Width           =   2955
      End
      Begin VB.Label Labelrotativas 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Vinheta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   0
         Left            =   -74685
         TabIndex        =   94
         Top             =   2520
         Width           =   2745
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   23
         Left            =   -68490
         TabIndex        =   79
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   22
         Left            =   -70590
         TabIndex        =   78
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   21
         Left            =   -72690
         TabIndex        =   77
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   20
         Left            =   -74790
         TabIndex        =   76
         Top             =   5985
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   19
         Left            =   -68490
         TabIndex        =   75
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   18
         Left            =   -70590
         TabIndex        =   74
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   17
         Left            =   -72690
         TabIndex        =   73
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   16
         Left            =   -74790
         TabIndex        =   72
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   15
         Left            =   -68490
         TabIndex        =   71
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   14
         Left            =   -70590
         TabIndex        =   70
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   13
         Left            =   -72690
         TabIndex        =   69
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   12
         Left            =   -74790
         TabIndex        =   68
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   11
         Left            =   -68490
         TabIndex        =   67
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   10
         Left            =   -70590
         TabIndex        =   66
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   9
         Left            =   -72690
         TabIndex        =   65
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   8
         Left            =   -74790
         TabIndex        =   64
         Top             =   2940
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   7
         Left            =   -68490
         TabIndex        =   63
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   6
         Left            =   -70590
         TabIndex        =   62
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   -72690
         TabIndex        =   61
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   -74790
         TabIndex        =   60
         Top             =   1995
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   -68490
         TabIndex        =   59
         Top             =   1155
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   -70590
         TabIndex        =   58
         Top             =   1155
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   -72690
         TabIndex        =   57
         Top             =   1155
         Width           =   2115
      End
      Begin VB.Label lblVinInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   -74790
         TabIndex        =   56
         Top             =   1155
         Width           =   2115
      End
   End
   Begin VB.Label lblHora2 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   13335
      TabIndex        =   224
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Label lblLembretes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lembretes"
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   9240
      TabIndex        =   190
      Top             =   5040
      Width           =   4005
   End
   Begin VB.Label LabelAvisos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Avisos Comerciais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   9240
      TabIndex        =   189
      Top             =   5670
      Width           =   5265
   End
   Begin VB.Label lblexecutando 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Executando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9240
      TabIndex        =   55
      Top             =   6300
      Width           =   5895
   End
   Begin VB.Label lblproximo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Próxima:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9240
      TabIndex        =   54
      Top             =   6720
      Width           =   5895
   End
   Begin VB.Menu arquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu abrirmusica 
         Caption         =   "Abrir música"
      End
      Begin VB.Menu svPlst 
         Caption         =   "Salvar Playlist"
      End
   End
   Begin VB.Menu opcoes 
      Caption         =   "Ferramentas"
      Begin VB.Menu reopenEspelho 
         Caption         =   "Reabrir Espelho"
      End
      Begin VB.Menu reopenCom 
         Caption         =   "Reabrir Comerciais"
      End
      Begin VB.Menu reopenAvisos 
         Caption         =   "Reabrir Avisos"
      End
      Begin VB.Menu reopenRotativas 
         Caption         =   "Reabrir vinhetas rotativas"
      End
      Begin VB.Menu mnulistadecom 
         Caption         =   "Lista de áudios executados"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Opções"
      Begin VB.Menu cfgEditor 
         Caption         =   "Preferências"
      End
   End
   Begin VB.Menu vinhetasMenu 
      Caption         =   "Vinhetas"
      Begin VB.Menu abrVinMnu 
         Caption         =   "Abrir vinhetas"
      End
      Begin VB.Menu svVinMnu 
         Caption         =   "Salvar Vinhetas"
      End
   End
   Begin VB.Menu sobre 
      Caption         =   "Sobre"
   End
   Begin VB.Menu popupcomerciais 
      Caption         =   "popupcomerciais"
      Visible         =   0   'False
      Begin VB.Menu removecom 
         Caption         =   "Remover comercial"
      End
   End
   Begin VB.Menu popupVinhetas 
      Caption         =   "popupVinhetas"
      Visible         =   0   'False
      Begin VB.Menu apagarslot 
         Caption         =   "Esvaziar slot"
      End
   End
   Begin VB.Menu popupPlaylist 
      Caption         =   "popupPlaylist"
      Visible         =   0   'False
      Begin VB.Menu mnuCart1 
         Caption         =   "Abrir na Cartucheira 1"
      End
      Begin VB.Menu mnuCart2 
         Caption         =   "Abrir na cartucheira 2"
      End
      Begin VB.Menu mnuOpenStream 
         Caption         =   "Abrir nova cartucheira"
      End
      Begin VB.Menu mnuApagar 
         Caption         =   "Apagar áudio"
      End
   End
   Begin VB.Menu popupFileExplorer 
      Caption         =   "popupFileExplorer"
      Visible         =   0   'False
      Begin VB.Menu mnuFileCart1 
         Caption         =   "Carregar na cartucheira 1"
         Index           =   1
      End
      Begin VB.Menu mnuFileCart1 
         Caption         =   "Carregar na cartucheira 2"
         Index           =   2
      End
      Begin VB.Menu mnuFileStream 
         Caption         =   "Abrir stream"
         Index           =   3
      End
      Begin VB.Menu mnuFileCue 
         Caption         =   "Abrir escuta (CUE)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_RESTORE = 9
Dim stream1 As Long
Dim stream1info As BASS_CHANNELINFO
Dim stream2 As Long
Dim stream2info As BASS_CHANNELINFO
Dim crosspoint1 As Long
Dim crosspoint2 As Long
Dim playlist() As audioItem 'LISTA DE AUDIOS DA PLAYLIST
Dim vinhetas(60) As vinhetas 'LISTA DE VINHETAS
Dim vinhetasRotativas(5) As vinhetas 'LISTA DE VINHETAS
Dim tamanhobloco As Long
Dim listaComercial() As comerciais 'lista dos diretórios de comerciais
Dim listaAvisos() As avisos
Dim ecadcounter As Integer
Dim rotatcnt(6) As Integer
Dim itemdeapoio(2) As String
Dim vinhetaIndex As Integer
Dim initdata(4) As String 'dados de inicialização de comerciais, espelho e avisos
Dim disposPrinc As Long
Dim disposCue As Long

Private Sub abrircomerciais_Click()
ReDim Preserve listaComercial(0)
dialogo1.DialogTitle = "Salvar playlist"
dialogo1.filename = vbNullString
dialogo1.Filter = "lista de comerciais (*.lsc)|*.lsc"
dialogo1.DefaultExt = "lsc"
dialogo1.flags = &H2
dialogo1.filename = vbNullString
dialogo1.ShowOpen
If Not (dialogo1.filename Like vbNullString) Then
    Call openfileCom(dialogo1.filename, listaComercial)
End If
End Sub


Private Sub apagarlista_Click()
If playlistView.ListItems.Count > 0 Then Call limparPlaylist(playlist, tamanhobloco, playlistView)
End Sub

Private Sub apagarslot_Click()
If Not (vinheta1(vinhetaIndex).Caption Like "vazio") And Not (lblVinInfo(vinhetaIndex).Caption Like "Info") Then
Call stopFile(vinhetas(vinhetaIndex).id)
vinheta1(vinhetaIndex).Caption = "vazio"
lblVinInfo(vinhetaIndex).Caption = "Info"
Call BASS_StreamFree(vinhetas(vinhetaIndex).id)
Call BASS_MusicFree(vinhetas(vinhetaIndex).id)
End If
End Sub

Private Sub armar2_Click()
If playlistView.ListItems.Count > 0 Then
    Dim stream As Form3
    Set stream = New Form3
    Call setInfoEvent(playlist(playlistView.SelectedItem.index)) 'função para diponibilizar o nome do arquivo a ser aberto para form3
    stream.Show , Me
    handle = "S" & LTrim$(Str(stream.hWnd))
    lstStreamAberto.ListItems.Add , handle, playlistView.SelectedItem.Text
    tocadaslist.AddItem (playlistView.SelectedItem.Text & " - " & Now)
End If
End Sub
Private Sub avisovisto_Click()
tmrAvisoPiscante2.Enabled = False
tmrAvisoPiscante1.Enabled = False
LabelAvisos.BackColor = &H808080
LabelAvisos.Caption = "Comerciais"
lblLembretes.BackColor = &H808080
lblLembretes.Caption = "Lembretes"
End Sub
Private Sub bw1_Click()
Call setPosition(stream1, -1000000)
End Sub

Private Sub bw2_Click()
Call setPosition(stream2, -1000000)
End Sub

Private Sub cfgEditor_Click()
Form4.Show
End Sub

Private Sub cmdAbrirVinSup_Click()
    dialogo1.DialogTitle = "Abrir vinhetas"
    dialogo1.filename = vbNullString
    dialogo1.Filter = "lista de vinhetas (*.vin)|*.vin"
    dialogo1.DefaultExt = "vin"
    dialogo1.flags = &H2
    dialogo1.filename = vbNullString
    dialogo1.ShowOpen
    Call openVinhetas(vinhetas, dialogo1.filename)
End Sub

Private Sub cmdBusca_Click()
If txtBusca.Text = "" Then Exit Sub
    cmdBusca.Enabled = False
    cmdPastaBase.Enabled = False
    txtBusca.Enabled = False
    ListViewBusca.ListItems.Clear
If findFiles(lblPastaBase.Caption, "*.*", txtBusca.Text) Then
    cmdCart1busca.Enabled = True
    cmdCart2busca.Enabled = True
    cmdToPlaylist.Enabled = True
    
End If
    cmdBusca.Enabled = True
    cmdPastaBase.Enabled = True
    txtBusca.Enabled = True
End Sub

Private Sub cmdCart1busca_Click()
If ListViewBusca.ListItems.Count > 0 Then
Call openFile(disposPrinc, ListViewBusca.SelectedItem.Key, ListViewBusca.SelectedItem.Text, stream1, stream1info, LabelDuracao1, infomp1, LabelPosition1, labelRest1, lblled1)
tocadaslist.AddItem (ListViewBusca.SelectedItem.Text & " - " & Now)
End If
End Sub

Private Sub cmdCart2busca_Click()
If ListViewBusca.ListItems.Count > 0 Then
Call openFile(disposPrinc, ListViewBusca.SelectedItem.Key, ListViewBusca.SelectedItem.Text, stream2, stream2info, Labelduracao2, infomp2, LabelPosition2, labelRest2, lblled2)
tocadaslist.AddItem (ListViewBusca.SelectedItem.Text & " - " & Now)
End If
End Sub

Private Sub cmdCue_Click()
    If FilePrincipal.filename <> vbNullString Then
    Dim item As audioItem
    Dim streamCue As Form2
    Set streamCue = New Form2
    For X = 0 To FilePrincipal.ListCount
        If FilePrincipal.Selected(X) Then
        item.nome = FilePrincipal.List(X)
        Exit For
        End If
    Next X
    item.path = FilePrincipal.path
    Call setInfoEvent(item) 'função para diponibilizar o nome do arquivo a ser aberto para form3
    streamCue.Show , Me
    handle = "S" & LTrim$(Str(streamCue.hWnd))
    lstStreamAberto.ListItems.Add , handle, item.nome
    End If
End Sub

Private Sub cmdFade1_Click()
TimerFade1.Enabled = True
End Sub

Private Sub cmdFade2_Click()
TimerFade2.Enabled = True
End Sub

Private Sub cmdHoraCerta_Click()
Dim pos As Integer
If playlistView.ListItems.Count > 0 Then pos = playlistView.SelectedItem.index + 1 Else pos = 1
Call toPlaylist(pos, -1, "HORACERTA", "HORACERTA", playlist, tamanhobloco, playlistView, vbNullString)
End Sub

Private Sub cmdLimparSup_Click()
For cont = 0 To 59
If Not (vinheta1(cont).Caption Like "vazio") And Not (lblVinInfo(cont).Caption Like "Info") Then
Call stopFile(vinhetas(cont).id)
vinheta1(cont).Caption = "vazio"
lblVinInfo(cont).Caption = "Info"
End If
Next cont
End Sub

Private Sub cmdPastaBase_Click()
Form8.Show
End Sub

Private Sub cmdPausa_Click()
Dim pos As Integer
If playlistView.ListItems.Count > 0 Then pos = playlistView.SelectedItem.index + 1 Else pos = 1
Call toPlaylist(pos, -1, "PAUSA", "PAUSA", playlist, tamanhobloco, playlistView, vbNullString)
End Sub
Private Sub cmdSvVinSup_Click()
    dialogo1.DialogTitle = "Salvar vinhetas"
    dialogo1.Filter = "lista de vinhetas (*.vin)|*.vin"
    dialogo1.DefaultExt = "vin"
    dialogo1.flags = &H2
    dialogo1.ShowSave
    Call saveVinhetas(vinhetas, dialogo1.filename)
End Sub
Private Sub cmdToPlaylist_Click()
Dim pos As Integer
If ListViewBusca.ListItems.Count > 0 Then
    If playlistView.ListItems.Count > 0 Then
        pos = playlistView.SelectedItem.index + 1
        Call toPlaylist(pos, 0, Left$(ListViewBusca.SelectedItem.Key, Len(ListViewBusca.SelectedItem.Key) - (Len(ListViewBusca.SelectedItem.Text) + 1)), ListViewBusca.SelectedItem.Text, playlist, tamanhobloco, playlistView, vbNullString)
    Else: Call toPlaylist(1, 0, Left$(ListViewBusca.SelectedItem.Key, Len(ListViewBusca.SelectedItem.Key) - (Len(ListViewBusca.SelectedItem.Text) + 1)), ListViewBusca.SelectedItem.Text, playlist, tamanhobloco, playlistView, vbNullString)
    End If
End If
End Sub

Private Sub cmdTreeCart1_Click()
Call openFile(disposPrinc, Espelho.SelectedItem.Key, Espelho.SelectedItem.Text, stream1, stream1info, LabelDuracao1, infomp1, LabelPosition1, labelRest1, lblled1)
tocadaslist.AddItem (Espelho.SelectedItem.Text & " - " & Now)
End Sub

Private Sub cmdTreeCart2_Click()
Call openFile(disposPrinc, Espelho.SelectedItem.Key, Espelho.SelectedItem.Text, stream2, stream2info, Labelduracao2, infomp2, LabelPosition2, labelRest2, lblled2)
tocadaslist.AddItem (Espelho.SelectedItem.Text & " - " & Now)
End Sub

Private Sub cmdVinRotativas_Click(index As Integer)
If cmdVinRotativas(index).Caption = "<vazio>" Then Exit Sub
If Not (BASS_ChannelIsActive(vinhetasRotativas(index).id) = BASS_ACTIVE_PLAYING) Then
    vinhetasRotativas(index).nome = ListaRotativas(index).ListItems(rotatcnt(index) + 1).Text
    vinhetasRotativas(index).caminho = ListaRotativas(index).ListItems(rotatcnt(index) + 1).Key
    Call openFile(disposPrinc, vinhetasRotativas(index).caminho, vinhetasRotativas(index).nome, vinhetasRotativas(index).id, vinhetasRotativas(index).info, Labelrotativas(index), cmdVinRotativas(index))
    rotatcnt(index) = rotatcnt(index) + 1
    PlayFile (vinhetasRotativas(index).id)
    If rotatcnt(index) = ListaRotativas(index).ListItems.Count Then rotatcnt(index) = 0
    Labelrotativas(index).Caption = Labelrotativas(index).Caption & vbCrLf & "PROX: " & ListaRotativas(index).ListItems(rotatcnt(index) + 1).Text
Else
    stopFile (vinhetasRotativas(index).id)
End If
End Sub
Private Sub espelhoToPlaylist_Click()
Dim pos As Integer
If playlistView.ListItems.Count > 0 Then pos = playlistView.SelectedItem.index Else pos = 1
Call toPlaylist(pos, 0, Left$(Espelho.SelectedItem.Key, Len(Espelho.SelectedItem.Key) - (Len(Espelho.SelectedItem.Text) + 1)), Espelho.SelectedItem.Text, playlist, tamanhobloco, playlistView, vbNullString)
End Sub
Private Sub deletaruma_Click()
If playlistView.ListItems.Count > 0 Then Call apagar(playlistView.SelectedItem.index, playlist, tamanhobloco, playlistView)
End Sub

Private Sub downlist_Click()
On Error GoTo error
If playlistView.ListItems.Count > 1 Then
    If playlistView.SelectedItem.index < playlistView.ListItems.Count Then
        permutaVetor playlistView.SelectedItem.index, 1, playlist
        itemdeapoio(0) = playlistView.ListItems(playlistView.SelectedItem.index + 1).Text
        itemdeapoio(1) = playlistView.ListItems(playlistView.SelectedItem.index + 1).Key
        playlistView.ListItems.Remove (playlistView.SelectedItem.index + 1)
        If playlistView.SelectedItem.index <> 1 Then
            playlistView.ListItems.Add playlistView.SelectedItem.index, itemdeapoio(1), itemdeapoio(0)
        Else
            playlistView.ListItems.Add 1, itemdeapoio(1), itemdeapoio(0)
        End If
    End If
End If
Exit Sub
error:
End Sub

Private Sub Espelho_NodeClick(ByVal Node As MSComctlLib.Node)
cmdTreeCart1.Enabled = True
cmdTreeCart2.Enabled = True
streamEspelho.Enabled = True
espelhoToPlaylist.Enabled = True
End Sub

Private Sub FilePrincipal_DblClick()
Dim pos As Integer
If playlistView.ListItems.Count > 0 Then pos = playlistView.SelectedItem.index + 1 Else pos = 1
Call toPlaylist(pos, 0, FilePrincipal.path, FilePrincipal.filename, playlist, tamanhobloco, playlistView, vbNullString)
End Sub
Private Sub Form_Initialize()
Dim path As String
Dim nome As String
'----------------------------------
Dim initbuffer As String
Dim initc As Integer
initc = 1
ReDim playlist(0)
ReDim listaComercial(0)
ReDim listaAvisos(0)
'para definir a placa principal e a de escuta
Call OpenDevices("c:\ressonance\devices.txt", disposPrinc, disposCue)

'--------ARQUIVO DE CONFIGURAÇÃO-------
On Error GoTo erroPrinc
Open "c:\ressonance\configressonance.txt" For Input As #10
Do Until EOF(10)
    Line Input #10, initbuffer
    If InStr(1, initbuffer, "#") Then
    initdata(initc) = Mid$(initbuffer, 2)
    initc = initc + 1
    End If
Loop
Close #10
textoLog.Text = textoLog.Text & vbCrLf & "Carregamento dos arquivo principal com sucesso"
'---------------------------------'
Call abrirAvisos(listaAvisos, initdata(1))
Call openfileCom(initdata(3), listaComercial)
'---------ESPELHO-----------------'
On Error GoTo fileerror
Espelho.Nodes.Add , , "root", "Raiz"
Espelho.Nodes("root").Expanded = True
Open initdata(2) For Input As #15
Do Until EOF(15)
    Line Input #15, path
    Line Input #15, nome
    Espelho.Nodes.Add "root", tvwChild, path, nome, 1
    'rotina para ler os arquivos dentro do diretório
    Call addNodes(path & "\", "*.*", path)
Loop
Close #15
textoLog.Text = textoLog.Text & vbCrLf & "Carregamento do espelho de " & initdata(2) & " com sucesso"
'------------VINHETAS ROTATIVAS------------------
initc = 0
Open initdata(4) For Input As #11
Do Until EOF(11) Or initc = 6
    Line Input #11, initbuffer
    Call readFolder(initbuffer, "*.mp3", ListaRotativas(initc), cmdVinRotativas(initc))
    initc = initc + 1
Loop
Close #11
textoLog.Text = textoLog.Text & vbCrLf & "Carregamento de vinhetas rotativas de " & initdata(4) & " com sucesso"
Exit Sub
erroPrinc:
    textoLog.Text = textoLog.Text & vbCrLf & "Erro no arquivo de config. principal configressonance.txt"
    Close #10
    Exit Sub
fileerror:
    textoLog.Text = textoLog.Text & vbCrLf & "Erro na abertura de rotativas ou espelho."
    Close #15
    Close #11
End Sub

Private Sub Form_Load()
Call IntializeBass(Form1)
Call StartVu
Me.Show 'para fazer o programa aparecer antes das rotinas de carregar espelho, etc...
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 conf = MsgBox("Tem certeza que deseja fechar?", vbYesNo, "Cuidado!")
If conf = 7 Then
    Cancel = True
    Exit Sub
Else
    Set Form1 = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call releaseBass
End 'para forçar o fim do programa
End Sub
'*************************************************************************'
Private Sub DrivePrincipal_Change()
On Error GoTo deviceError
    DirPrincipal.path = DrivePrincipal.Drive
    Exit Sub
deviceError:
    MsgBox "Drive " & DrivePrincipal.Drive & " não disponível"
End Sub

Private Sub DirPrincipal_Change()
    FilePrincipal.path = DirPrincipal.path
End Sub
'*************************************************************************'
Private Sub cmdPreview1_Click()
If FilePrincipal.filename <> vbNullString Then
Call openFile(disposPrinc, FilePrincipal.path & "\" & FilePrincipal.filename, FilePrincipal.filename, stream1, stream1info, LabelDuracao1, infomp1, LabelPosition1, labelRest1, lblled1)
tocadaslist.AddItem (FilePrincipal.filename & " - " & Now)
End If
End Sub
Private Sub cmdPreview2_Click()
If FilePrincipal.filename <> vbNullString Then
Call openFile(disposPrinc, FilePrincipal.path & "\" & FilePrincipal.filename, FilePrincipal.filename, stream2, stream2info, Labelduracao2, infomp2, LabelPosition2, labelRest2, lblled2)
tocadaslist.AddItem (FilePrincipal.filename & " - " & Now)
End If
End Sub

Private Sub fw1_Click()
Call setPosition(stream1, 1000000)
End Sub

Private Sub fw2_Click()
Call setPosition(stream2, 1000000)
End Sub
Private Sub lstStreamAberto_DblClick()
ShowWindow CLng(Right$(lstStreamAberto.SelectedItem.Key, Len(lstStreamAberto.SelectedItem.Key) - 1)), SW_RESTORE 'função api para restaurar janela
End Sub

Private Sub mnuApagar_Click()
If playlistView.ListItems.Count > 0 Then Call apagar(playlistView.SelectedItem.index, playlist, tamanhobloco, playlistView)
End Sub

Private Sub mnuCart1_Click()
Call player1_Click
End Sub

Private Sub mnuCart2_Click()
Call player2_Click
End Sub

Private Sub mnuOpenStream_Click()
Call armar2_Click
End Sub

Private Sub pause1_Click()
Call pauseFile(stream1)
End Sub

Private Sub pause2_Click()
Call pauseFile(stream2)
End Sub

Private Sub play1_Click()
Call PlayFile(stream1)
End Sub

Private Sub play2_Click()
Call PlayFile(stream2)
End Sub

Private Sub player1_Click()
If playlistView.ListItems.Count > 0 Then
    tocadaslist.AddItem (playlistView.SelectedItem.Text & " - " & Now)
    Call openFile(disposPrinc, playlist(playlistView.SelectedItem.index).path & "\" & playlist(playlistView.SelectedItem.index).nome, playlist(playlistView.SelectedItem.index).nome, stream1, stream1info, LabelDuracao1, infomp1, LabelPosition1, labelRest1, lblled1)
    If autocross.value = 0 Then
    'coloca o crosspoint num ponto além do final do audio para não acionar por engano nem comer a musica pela próxima
    Call SetCrossPt(stream1, crosspoint1)
    crosspoint1 = crosspoint1 + crosspoint1
    'coloca o crosspoint no ponto correto
    Else: Call SetCrossPt(stream1, crosspoint1)
    End If
    If playlistView.SelectedItem.Text = "HORACERTA" Then
        Call toPlaylist(1, 0, "c:\ressonance\hora", "HRS" & Hour(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
        Call toPlaylist(2, 0, "c:\ressonance\hora", "MIN" & Minute(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
    End If
    If Not BASS_ChannelIsActive(stream1) = BASS_ACTIVE_PLAYING Then Call apagar(playlistView.SelectedItem.index, playlist, tamanhobloco, playlistView)
If (Not BASS_ChannelIsActive(stream1) = BASS_ACTIVE_PLAYING) And autoload.value = 1 Then
    PlayFile (stream1)
End If
End If
End Sub

Private Sub player2_Click()
If playlistView.ListItems.Count > 0 Then
    tocadaslist.AddItem (playlistView.SelectedItem.Text & " - " & Now)
    Call openFile(disposPrinc, playlist(playlistView.SelectedItem.index).path & "\" & playlist(playlistView.SelectedItem.index).nome, playlist(playlistView.SelectedItem.index).nome, stream2, stream2info, Labelduracao2, infomp2, LabelPosition2, labelRest2, lblled2)
    If autocross.value = 0 Then
    Call SetCrossPt(stream2, crosspoint2)
    crosspoint2 = crosspoint2 + crosspoint2
    Else: Call SetCrossPt(stream2, crosspoint2)
    End If
    If playlistView.SelectedItem.Text = "HORACERTA" Then
        Call toPlaylist(1, 0, "c:\ressonance\hora", "HRS" & Hour(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
        Call toPlaylist(2, 0, "c:\ressonance\hora", "MIN" & Minute(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
    End If
    If Not BASS_ChannelIsActive(stream2) = BASS_ACTIVE_PLAYING Then Call apagar(playlistView.SelectedItem.index, playlist, tamanhobloco, playlistView)
    If (Not BASS_ChannelIsActive(stream2) = BASS_ACTIVE_PLAYING) And autoload.value = 1 Then
        PlayFile (stream2)
    End If
End If
End Sub
Private Sub playlistView_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then Call apagar(playlistView.SelectedItem.index, playlist, tamanhobloco, playlistView)
If KeyCode = vbKey1 Then Call player1_Click
If KeyCode = vbKey2 Then Call player2_Click
End Sub

Private Sub playlistView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
    If playlistView.ListItems.Count <= 0 Then
        mnuCart1.Enabled = False
        mnuCart2.Enabled = False
        mnuApagar.Enabled = False
        mnuOpenStream.Enabled = False
    Else
        mnuCart1.Enabled = True
        mnuCart2.Enabled = True
        mnuApagar.Enabled = True
        mnuOpenStream.Enabled = True
    End If
 PopupMenu popupPlaylist
 End If
End Sub

Private Sub reopenAvisos_Click()
Call abrirAvisos(listaAvisos, initdata(1))
End Sub

Private Sub reopenCom_Click()
Call openfileCom(initdata(3), listaComercial)
End Sub

Private Sub reopenEspelho_Click()
On Error GoTo fileerror
Dim path As String
Dim nome As String
Espelho.Nodes.Clear
Espelho.Nodes.Add , , "root", "Raiz"
Espelho.Nodes("root").Expanded = True
Open initdata(2) For Input As #15
Do Until EOF(15)
    Line Input #15, path
    Line Input #15, nome
    Espelho.Nodes.Add "root", tvwChild, path, nome, 1
    'rotina para ler os arquivos dentro do diretório
    Call addNodes(path & "\", "*.*", path)
Loop
Close #15
Exit Sub
fileerror:
Espelho.Nodes.Clear
MsgBox "Erro na reabertura do espelho", vbInformation
End Sub

Private Sub reopenRotativas_Click()
Dim initc As Integer
Dim initbuffer As String
initc = 0
Open initdata(4) For Input As #11
Do Until EOF(11) Or initc = 6
    ListaRotativas(initc).ListItems.Clear
    rotatcnt(initc) = 0
    Line Input #11, initbuffer
    Call readFolder(initbuffer, "*.mp3", ListaRotativas(initc), cmdVinRotativas(initc))
    initc = initc + 1
Loop
Close #11
Exit Sub
End Sub

Private Sub ScrollVol1_Change()
Call setVolume(stream1, ScrollVol1.value)
End Sub

Private Sub ScrollVol1_Scroll()
Call setVolume(stream1, ScrollVol1.value)
End Sub

Private Sub ScrollVol2_Change()
Call setVolume(stream2, ScrollVol2.value)
End Sub

Private Sub ScrollVol2_Scroll()
Call setVolume(stream2, ScrollVol2.value)
End Sub

Private Sub stop1_Click()
Call stopFile(stream1)
LabelMode1.BackColor = &H0&
End Sub

Private Sub stop2_Click()
Call stopFile(stream2)
LabelMode2.BackColor = &H0&
End Sub

Private Sub streamEspelho_Click()
    Dim item2 As audioItem
    Dim stream As Form3
    Set stream = New Form3
    item2.nome = Espelho.SelectedItem.Text
    item2.path = Left$(Espelho.SelectedItem.Key, Len(Espelho.SelectedItem.Key) - (Len(Espelho.SelectedItem.Text) + 1))
    Call setInfoEvent(item2) 'função para diponibilizar o nome do arquivo a ser aberto para form3
    stream.Show , Me
    handle = "S" & LTrim$(Str(stream.hWnd))
    lstStreamAberto.ListItems.Add , handle, item2.nome
End Sub

Private Sub streamPrev_Click()
    If FilePrincipal.filename <> vbNullString Then
    Dim item As audioItem
    Dim stream As Form3
    Set stream = New Form3
    For X = 0 To FilePrincipal.ListCount
        If FilePrincipal.Selected(X) Then
        item.nome = FilePrincipal.List(X)
        Exit For
        End If
    Next X
    item.path = FilePrincipal.path
    Call setInfoEvent(item) 'função para diponibilizar o nome do arquivo a ser aberto para form3
    stream.Show , Me
    handle = "S" & LTrim$(Str(stream.hWnd))
    lstStreamAberto.ListItems.Add , handle, item.nome
    tocadaslist.AddItem (FilePrincipal.filename & " - " & Now)
    End If
End Sub

Private Sub TimerFade1_Timer()
If BASS_ChannelIsActive(stream1) = BASS_ACTIVE_PLAYING Then
    If ScrollVol1.value = 0 Then
        Call stopFile(stream1)
        TimerFade1.Enabled = False
        ScrollVol1.value = 10
        Exit Sub
    Else
        ScrollVol1.value = ScrollVol1.value - 1
        Call setVolume(stream1, ScrollVol1.value)
    End If
Else
    TimerFade1.Enabled = False
    ScrollVol1.value = 10
End If
End Sub

Private Sub TimerFade2_Timer()
If BASS_ChannelIsActive(stream2) = BASS_ACTIVE_PLAYING Then
    If ScrollVol2.value = 0 Then
        Call stopFile(stream2)
        TimerFade2.Enabled = False
        ScrollVol2.value = 10
        Exit Sub
    Else
        ScrollVol2.value = ScrollVol2.value - 1
        Call setVolume(stream2, ScrollVol2.value)
    End If
Else
    TimerFade2.Enabled = False
    ScrollVol2.value = 10
End If
End Sub

Private Sub tmrAvisoPiscante1_Timer()
If lblLembretes.BackColor = &HFF& Then lblLembretes.BackColor = &H808080 Else lblLembretes.BackColor = &HFF&
End Sub

Private Sub tmrAvisoPiscante2_Timer()
If LabelAvisos.BackColor = &HFF& Then LabelAvisos.BackColor = &H808080 Else LabelAvisos.BackColor = &HFF&
End Sub

Private Sub trasfermulti_Click()
If FilePrincipal.filename <> vbNullString Then
Dim pos As Integer
If playlistView.ListItems.Count > 0 Then pos = playlistView.SelectedItem.index + 1 Else pos = 1
For X = 0 To (FilePrincipal.ListCount - 1)
  If FilePrincipal.Selected(X) Then
    Call toPlaylist(pos, 0, FilePrincipal.path, FilePrincipal.List(X), playlist, tamanhobloco, playlistView, vbNullString)
    pos = pos + 1
  End If
  DoEvents 'para carregar dinamicamente
Next X
End If
End Sub
Private Sub txtBusca_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If txtBusca.Text = "" Then Exit Sub
        cmdBusca.Enabled = False
        cmdPastaBase.Enabled = False
        txtBusca.Enabled = False
        ListViewBusca.ListItems.Clear
    If findFiles(lblPastaBase.Caption, "*.*", txtBusca.Text) Then
        cmdCart1busca.Enabled = True
        cmdCart2busca.Enabled = True
        cmdToPlaylist.Enabled = True
    End If
        cmdBusca.Enabled = True
        cmdPastaBase.Enabled = True
        txtBusca.Enabled = True
End If
End Sub

Private Sub uplist_Click()
On Error GoTo error
If playlistView.ListItems.Count > 1 Then
    If playlistView.SelectedItem.index > 1 Then
        permutaVetor playlistView.SelectedItem.index, 0, playlist
        itemdeapoio(0) = playlistView.ListItems(playlistView.SelectedItem.index - 1).Text
        itemdeapoio(1) = playlistView.ListItems(playlistView.SelectedItem.index - 1).Key
        playlistView.ListItems.Remove (playlistView.SelectedItem.index - 1)
        playlistView.ListItems.Add playlistView.SelectedItem.index + 1, itemdeapoio(1), itemdeapoio(0)
    End If
End If
Exit Sub
error:
End Sub

'************************************************VINHETAS********************
Private Sub vinheta1_Click(index As Integer)
If vinheta1(index).Caption = "vazio" Then
    dialogo1.Filter = "arquivos mp3 e wav (*.mp3,*.wav)|*.mp3;*.wav"
    dialogo1.flags = &H1000
    dialogo1.filename = vbNullString
    dialogo1.ShowOpen
    If dialogo1.filename = vbNullString Then Exit Sub
    vinhetas(index).nome = dialogo1.FileTitle
    vinhetas(index).caminho = dialogo1.filename
    Call openFile(disposPrinc, vinhetas(index).caminho, vinhetas(index).nome, vinhetas(index).id, vinhetas(index).info, lblVinInfo(index), vinheta1(index))
Else
    If BASS_ChannelIsActive(vinhetas(index).id) = BASS_ACTIVE_PLAYING Then
        stopFile (vinhetas(index).id)
    Else
        PlayFile (vinhetas(index).id)
    End If
End If
End Sub
'***************************************************************************'
Private Sub tmrSpectrum_Timer()
    Call UpdateSpectrum(0, 0, 0, 0, 0, stream1) ' the params are if using the API MM timer
    Call UpdateSpectrum(0, 0, 0, 0, 0, stream2) ' the params are if using the API MM timer
End Sub

Private Sub vinheta1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
    vinhetaIndex = index
    PopupMenu popupVinhetas
 End If
End Sub

Private Sub vuMeter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    specmode = (specmode + 1) Mod 4  ' swap spectrum mode
    ReDim specbuf(SPECWIDTH * (SPECHEIGHT + 1)) As Byte ' clear display
End Sub

Private Sub TimerProgressao_Timer()
Call updateTimes(stream1, LabelPosition1, labelRest1, LabelMode1, lblled1)
Call updateTimes(stream2, LabelPosition2, labelRest2, LabelMode2, lblled2)
For cont = 0 To 59
    Call updateTimes(vinhetas(cont).id, , lblVinInfo(cont), , lblVinInfo(cont))
    DoEvents
Next cont
End Sub

Private Sub TimerEventos_Timer()
lblhora.Caption = "Bom dia! Hoje é " & Date & " ."
lblHora2.Caption = time
Call verifyAdv(time, listaComercial, playlist, tamanhobloco)
Call verifAvisos(time, listaAvisos)
ecadcounter = ecadcounter + 1
Call salvaLog(ecadcounter)
If BASS_ChannelIsActive(stream1) = BASS_ACTIVE_PLAYING Then
    If LabelMode1.BackColor = &H0& Then LabelMode1.BackColor = &HFF& Else LabelMode1.BackColor = &H0&
Else: LabelMode1.BackColor = &H0&
End If
If BASS_ChannelIsActive(stream2) = BASS_ACTIVE_PLAYING Then
    If LabelMode2.BackColor = &H0& Then LabelMode2.BackColor = &HFF& Else LabelMode2.BackColor = &H0&
Else: LabelMode2.BackColor = &H0&
End If
End Sub
Private Sub timerAutocross_Timer()
If playlistView.ListItems.Count > 0 Then
    If playlist(1).nome = "PAUSA" Then
    autocross.value = 0
    timerAutocross.Enabled = False
    Exit Sub
    End If
    If isAutocross(stream1, crosspoint1) Then
        Call closeFile(stream2)
    
        tocadaslist.AddItem (playlistView.ListItems(1).Text & " - " & Now)
        'HORA CERTA
        If playlistView.ListItems(1).Text = "HORACERTA" Then
            Call toPlaylist(2, 0, "c:\ressonance\hora", "HRS" & Hour(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
            Call toPlaylist(3, 0, "c:\ressonance\hora", "MIN" & Minute(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
        End If
        If Not (openFile(disposPrinc, playlist(1).path & "\" & playlist(1).nome, playlist(1).nome, stream2, stream2info, Labelduracao2, infomp2, LabelPosition2, labelRest2, lblled2)) Then
            autocross.value = 0
            timerAutocross.Enabled = False
            MsgBox "Erro na abertura do áudio", vbInformation
        End If
        Call SetCrossPt(stream2, crosspoint2)
        
        'SERVE para que o fade não interfira na proxima musica
        TimerFade2.Enabled = False
        'SERVE para o volume voltar a posição maxima caso o fade naum va ate o final
        ScrollVol2.value = 10
        
        Call PlayFile(stream2)
        'Serve para abaixar o volume em fade de musicas sem rabicho
         If (BASS_ChannelBytes2Seconds(stream1, BASS_ChannelGetLength(stream1, BASS_POS_BYTE))) > 90 Then TimerFade1.Enabled = True
        
        lblexecutando.Caption = "EXECUTANDO: " & playlist(1).nome
        Call apagar(1, playlist, tamanhobloco, playlistView)
        If playlistView.ListItems.Count > 0 Then lblproximo.Caption = "PRÓXIMO: " & playlist(1).nome
    ElseIf isAutocross(stream2, crosspoint2) Then
        Call closeFile(stream1)
        
        tocadaslist.AddItem (playlistView.ListItems(1).Text & " - " & Now)
        'HORA CERTA
        If playlistView.ListItems(1).Text = "HORACERTA" Then
            Call toPlaylist(2, 0, "c:\ressonance\hora", "HRS" & Hour(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
            Call toPlaylist(3, 0, "c:\ressonance\hora", "MIN" & Minute(Now) & ".mp3", playlist, tamanhobloco, playlistView, vbNullString)
        End If
        If Not (openFile(disposPrinc, playlist(1).path & "\" & playlist(1).nome, playlist(1).nome, stream1, stream1info, LabelDuracao1, infomp1, LabelPosition1, labelRest1, lblled1)) Then
            autocross.value = 0
            timerAutocross.Enabled = False
            MsgBox "Erro na abertura do áudio", vbInformation
        End If
        Call SetCrossPt(stream1, crosspoint1)
        'SERVE para que o fade não interfira na proxima musica
        TimerFade1.Enabled = False
        'SERVE para o volume voltar a posição maxima caso o fade naum va ate o final
        ScrollVol1.value = 10
        
        Call PlayFile(stream1)
        'Serve para abaixar o volume em fade de musicas sem rabicho
         If (BASS_ChannelBytes2Seconds(stream2, BASS_ChannelGetLength(stream2, BASS_POS_BYTE))) > 90 Then TimerFade2.Enabled = True
        
        lblexecutando.Caption = "EXECUTANDO: " & playlist(1).nome
        Call apagar(1, playlist, tamanhobloco, playlistView)
        If playlistView.ListItems.Count > 0 Then lblproximo.Caption = "PRÓXIMO: " & playlist(1).nome
    End If
Else
    timerAutocross.Enabled = False
    autocross.value = 0
End If
End Sub
'---------------------------------------------------I/O---------------------------------------------------
Private Sub salvarlista_Click()
If playlistView.ListItems.Count > 0 Then
    dialogo1.DialogTitle = "Salvar playlist"
    dialogo1.filename = vbNullString
    dialogo1.Filter = "lista de músicas (*.lst)|*.lst"
    dialogo1.DefaultExt = "lst"
    dialogo1.flags = &H2
    dialogo1.ShowSave
    If Not (dialogo1.filename Like vbNullString) Then
        Call savefileList(dialogo1.filename, playlist)
    End If
End If
End Sub

Private Sub abrirlista_Click()
dialogo1.DialogTitle = "Abrir playlist"
dialogo1.filename = vbNullString
dialogo1.Filter = "lista de músicas (*.lst)|*.lst"
dialogo1.DefaultExt = "lst"
dialogo1.flags = &H2
dialogo1.ShowOpen
If Not (dialogo1.filename Like vbNullString) Then
    Call openfileList(dialogo1.filename, playlist, tamanhobloco)
End If
End Sub
'******************************************************other**************************************************
Private Sub autocross_Click()
If autocross.value = 0 Then
    timerAutocross.Enabled = False
    autocross.BackColor = &HFF&
    'evita q o botão autocross funcione como autoplay depois q o autocross é reativado:
    crosspoint1 = crosspoint1 + crosspoint1
    crosspoint2 = crosspoint2 + crosspoint2
Else
    autocross.BackColor = &HFF00&
    'recalcula o crosspoint da musica caso o autocross seja reativado:
    If BASS_ChannelIsActive(stream1) = BASS_ACTIVE_PLAYING Then Call SetCrossPt(stream1, crosspoint1)
    If BASS_ChannelIsActive(stream2) = BASS_ACTIVE_PLAYING Then Call SetCrossPt(stream2, crosspoint2)
    timerAutocross.Enabled = True
End If
End Sub

Private Sub autoload_Click()
If autoload.value = 0 Then
    autoload.BackColor = &HFF&
Else
    autoload.BackColor = &HFF00&
End If
End Sub

Public Sub setDevices(ByRef principal As Long, ByRef cue As Long)
disposPrinc = principal
disposCue = cue
End Sub
Public Function getPrinc() As Long
getPrinc = disposPrinc
End Function

Public Function getCue() As Long
getCue = disposCue
End Function

Friend Function getPlaylist() As audioItem()
getPlaylist = playlist
End Function

Public Function getTamBloco() As Long
getTamBloco = tamanhobloco
End Function



