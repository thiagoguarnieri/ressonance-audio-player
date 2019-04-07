VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Selecione a Pasta onde deseja Buscar o arquivo"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form8"
   ScaleHeight     =   4725
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   390
      Left            =   5025
      TabIndex        =   2
      Top             =   4275
      Width           =   840
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   3465
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   75
      TabIndex        =   0
      Top             =   525
      Width           =   5790
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Right$(Dir1.path, "1") Like "\" Then
    Form1.lblPastaBase.Caption = Dir1.path
Else
    Form1.lblPastaBase.Caption = Dir1.path & "\"
End If
Form1.cmdBusca.Enabled = True
Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo error
Dir1.path = Drive1.Drive
Exit Sub
error:
MsgBox "Drive " & Drive1.Drive & " não disponível"
End Sub

