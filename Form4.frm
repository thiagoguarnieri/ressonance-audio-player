VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form4"
   ScaleHeight     =   3045
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListDevPrinc 
      Height          =   1665
      Left            =   225
      TabIndex        =   4
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2937
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdOkPlacas 
      Caption         =   "OK"
      Height          =   420
      Left            =   5700
      TabIndex        =   3
      Top             =   2550
      Width           =   735
   End
   Begin VB.Frame frmPlacas 
      Height          =   2490
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      Begin MSComctlLib.ListView ListDevCue 
         Height          =   1665
         Left            =   3225
         TabIndex        =   5
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2937
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Selecione a placa de pré escuta"
         Height          =   240
         Left            =   3225
         TabIndex        =   2
         Top             =   375
         Width           =   2805
      End
      Begin VB.Label Label1 
         Caption         =   "Selecione a placa principal"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   2280
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkPlacas_Click()
Call Form1.setDevices(ListDevPrinc.SelectedItem.index, ListDevCue.SelectedItem.index)
'salva informações de placa principal e cue
Call saveDevices("c:\ressonance\devices.txt", ListDevPrinc.SelectedItem.index, ListDevCue.SelectedItem.index)
Unload Me
End Sub
Private Sub Form_Load()
    Dim i As BASS_DEVICEINFO
    Dim c As Integer
    c = 1
     While BASS_GetDeviceInfo(c, i)
        If (i.flags And BASS_DEVICE_ENABLED) Then  ' enabled, so add it...
            ListDevPrinc.ListItems.Add , , VBStrFromAnsiPtr(i.name)
            ListDevCue.ListItems.Add , , VBStrFromAnsiPtr(i.name)
        End If
        c = c + 1
    Wend
End Sub

