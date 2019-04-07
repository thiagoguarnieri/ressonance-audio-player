VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "CUE"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2190
   LinkTopic       =   "Form2"
   ScaleHeight     =   2085
   ScaleWidth      =   2190
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFade 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   975
      Top             =   0
   End
   Begin VB.Timer TimerLoop 
      Interval        =   1
      Left            =   1425
      Top             =   0
   End
   Begin VB.Timer TimerStatus 
      Interval        =   100
      Left            =   675
      Top             =   0
   End
   Begin VB.CommandButton play 
      Height          =   420
      Left            =   0
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton stop 
      Height          =   420
      Left            =   450
      Picture         =   "Form2.frx":37B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton pause 
      Height          =   420
      Left            =   900
      Picture         =   "Form2.frx":6E8D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton bw 
      Height          =   420
      Left            =   1350
      Picture         =   "Form2.frx":A5EB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton fw 
      Height          =   420
      Left            =   1800
      Picture         =   "Form2.frx":DED8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   420
   End
   Begin MSComCtl2.FlatScrollBar Slider1 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   1170
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      Max             =   10
      Orientation     =   1245185
      Value           =   10
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
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   855
      Width           =   1005
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
      Height          =   285
      Left            =   1125
      TabIndex        =   9
      Top             =   855
      Width           =   1050
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
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Top             =   540
      Width           =   1005
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
      Height          =   285
      Left            =   1125
      TabIndex        =   7
      Top             =   540
      Width           =   1050
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
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   1485
      Width           =   2085
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Form_Load()
streamData = getInfoEvent
Call openFile(Form1.getCue, streamData.path & "\" & streamData.nome, streamData.nome, stream, streaminfo, streamTot, LabelNome, streamProg, streamRest, streamMode)
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
Private Sub play_Click()
Call PlayFile(stream)
Form1.tocadaslist.AddItem (Now & "-" & LabelNome.Caption)
End Sub

Private Sub Slider1_Change()
Call setVolume(stream, Slider1.value)
End Sub

Private Sub Slider1_Scroll()
Call setVolume(stream, Slider1.value)
End Sub

Private Sub stop_Click()
Call stopFile(stream)
End Sub

Private Sub TimerStatus_Timer()
Call updateTimes(stream, streamProg, streamRest, streamMode, streamMode)
End Sub

