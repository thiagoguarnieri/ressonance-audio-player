Attribute VB_Name = "modRandomEvents"
'detecta comerciais
Public Sub verifyAdv(ByRef hora As String, vetorcomercial() As comerciais, playlist() As audioItem, tambloco As Long)
    If UBound(vetorcomercial) > 0 Then
        Form1.MonitorCom.Caption = "Monitorando"
        For cont = 1 To UBound(vetorcomercial)
            'eventos diários
            If (vetorcomercial(cont).mode Like "diário" And vetorcomercial(cont).horario Like hora) Then
                Call scanFolder(vetorcomercial(cont).fldrPath, "*.mp3", playlist, tambloco, "[EVT]")
                Call scanFolder(vetorcomercial(cont).fldrPath, "*.wav", playlist, tambloco, "[EVT]")
                Call scanFolder(vetorcomercial(cont).fldrPath, "*.lst", playlist, tambloco, "[MUS]")
                Form1.LabelAvisos.Caption = "EVENTO AUTOMÁTICO"
                Form1.tmrAvisoPiscante2.Enabled = True
            End If
            'eventos não diários
            If (vetorcomercial(cont).mode <> "diário" And vetorcomercial(cont).mode Like Format(Now(), "dddd") And vetorcomercial(cont).horario Like hora) Then
                Call scanFolder(vetorcomercial(cont).fldrPath, "*.mp3", playlist, tambloco, "[EVT]")
                Call scanFolder(vetorcomercial(cont).fldrPath, "*.wav", playlist, tambloco, "[EVT]")
                Call scanFolder(vetorcomercial(cont).fldrPath, "*.lst", playlist, tambloco, "[MUS]")
                Form1.LabelAvisos.Caption = "EVENTO AUTOMÁTICO AVULSO"
                Form1.tmrAvisoPiscante2.Enabled = True
            End If
        Next cont
    End If
End Sub
'verifica se tem algum lembrete
Public Sub verifAvisos(ByRef hora As String, listaAvisos() As avisos)
    For cont = 1 To UBound(listaAvisos)
        If listaAvisos(cont).horario = hora Then
            Form1.lblLembretes.Caption = listaAvisos(cont).aviso
            Form1.tmrAvisoPiscante1.Enabled = True
            Exit Sub
        End If
    Next cont
End Sub
