VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player - AUX"
   ClientHeight    =   2055
   ClientLeft      =   5355
   ClientTop       =   5595
   ClientWidth     =   2250
   FillStyle       =   0  'Solid
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2250
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "FADE "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1065
      TabIndex        =   12
      Top             =   840
      Width           =   585
   End
   Begin MSComCtl2.FlatScrollBar Slider1 
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   1140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      Arrows          =   65536
      Max             =   10
      Orientation     =   1245185
      Value           =   10
   End
   Begin VB.CheckBox Loopcheck 
      Caption         =   "LOOP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1665
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   555
   End
   Begin VB.CommandButton fw 
      Height          =   420
      Left            =   1800
      Picture         =   "Form3.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton bw 
      Height          =   420
      Left            =   1350
      Picture         =   "Form3.frx":459F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton pause 
      Height          =   420
      Left            =   900
      Picture         =   "Form3.frx":7E8C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton stop 
      Height          =   420
      Left            =   450
      Picture         =   "Form3.frx":B5EA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton play 
      Height          =   420
      Left            =   0
      Picture         =   "Form3.frx":ECC7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   420
   End
   Begin VB.Timer TimerStatus 
      Interval        =   1000
      Left            =   1470
      Top             =   0
   End
   Begin VB.Timer tmrFade 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   975
      Top             =   0
   End
   Begin VB.Label LabelNome 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   525
      Left            =   60
      TabIndex        =   10
      Top             =   1500
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label streamRest 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   1125
      TabIndex        =   8
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label streamProg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   45
      TabIndex        =   7
      Top             =   540
      Width           =   990
   End
   Begin VB.Label streamMode 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label streamTot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stream As Long
Dim streaminfo As BASS_CHANNELINFO
Dim streamData As audioItem

Private Sub bw_Click()
Call setPosition(stream, -1000000)
End Sub

Private Sub Command1_Click()
tmrFade.Enabled = True
End Sub

Private Sub Form_Load()
streamData = getInfoEvent
Call openFile(Form1.getPrinc, streamData.path & "\" & streamData.nome, streamData.nome, stream, streaminfo, streamTot, LabelNome, streamProg, streamRest, streamMode)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Dim Form As Form
    For Each Form In Forms
        If Form Is Me Then
            Set Form = Nothing
            Exit For
        End If
    Next Form
    Set Form3 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim streamHandle As String
Call closeFile(stream)
For cont = 1 To Form1.lstStreamAberto.ListItems.Count
    streamHandle = "S" & LTrim$(Str(Me.hWnd))
    If Form1.lstStreamAberto.ListItems(cont).Key = streamHandle Then
       Form1.lstStreamAberto.ListItems.Remove (cont)
       Exit Sub
    End If
Next cont
End Sub

Private Sub fw_Click()
Call setPosition(stream, 1000000)
End Sub

Private Sub Loopcheck_Click()
If Loopcheck.value = 1 Then
Call BASS_ChannelFlags(stream, BASS_SAMPLE_LOOP, BASS_SAMPLE_LOOP) 'coloca a flag de loop
Else
Call BASS_ChannelFlags(stream, 0, BASS_SAMPLE_LOOP) 'tira
End If
End Sub

Private Sub pause_Click()
Call pauseFile(stream)
End Sub

Private Sub play_Click()
Call PlayFile(stream)
End Sub

Private Sub Slider1_Change()
Call setVolume(stream, Slider1.value)
End Sub

Private Sub Slider1_Scroll()
Call setVolume(stream, Slider1.value)
End Sub

Private Sub stop_Click()
Call stopFile(stream)
LabelNome.BackColor = &H0&
End Sub
Private Sub TimerStatus_Timer()
Call updateTimes(stream, streamProg, streamRest, streamMode, streamMode)
If BASS_ChannelIsActive(stream) = BASS_ACTIVE_PLAYING Then
    If LabelNome.BackColor = &H0& Then LabelNome.BackColor = &HFF& Else LabelNome.BackColor = &H0&
Else: LabelNome.BackColor = &H0&
End If
End Sub

Private Sub tmrFade_Timer()
If BASS_ChannelIsActive(stream) = BASS_ACTIVE_PLAYING Then
    If Slider1.value = 0 Then
        Call stopFile(stream)
        tmrFade.Enabled = False
        Slider1.value = 10
        Exit Sub
    Else
        Slider1.value = Slider1.value - 1
        Call setVolume(stream, Slider1.value)
    End If
Else
    tmrFade.Enabled = False
    Slider1.value = 10
End If
End Sub
