Attribute VB_Name = "modExecução"
'módulo para uso das funções do módulo bass
Dim defaultObj As Label

Function IntializeBass(Form As Object)
' change and set the current path, to prevent from VB not finding BASS.DLL
    ChDrive App.path
    ChDir App.path

    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
        End
    End If

    ' initialize BASS
    If (BASS_Init(1, 44100, 0, Form.hWnd, 0) = 0) Then
        MsgBox ("O dipositivo primário não pôde ser inciado")
        IntializeBass = False
    End If
    If (BASS_Init(2, 44100, 0, Form.hWnd, 0) = 0) Then
        Form1.cmdCue.Enabled = False
    End If
    IntializeBass = True
End Function
'COLOCANDO PARÂMETROS COMO OPTIONAL PARA SIMULAR UMA SOBRECARGA DE MÉTODO
Public Function openFile(device As Long, file As String, name As String, stream As Long, info As BASS_CHANNELINFO, Optional labelDur As Object, Optional nome As Object, Optional position As Object, Optional rest As Object, Optional led As Object) As Boolean
    On Local Error Resume Next    ' if Cancel pressed...
    If name = "PAUSA" Then
        Form1.LabelAvisos.BackColor = &HFF&
        Form1.LabelAvisos.Caption = "A pausa serve somente para parar execução automática da lista"
        Exit Function
    ElseIf name = "HORACERTA" Then
        'AKI IRÁ AS ROTINAS DE CARREGAR A HORA CERTA
        file = "c:\ressonance\hora\INTRO.mp3"
        name = "intro.mp3"
    End If
    If BASS_ChannelIsActive(stream) = BASS_ACTIVE_PLAYING Then
        Form1.LabelAvisos.BackColor = &HFF&
        Form1.LabelAvisos.Caption = "Cartucheira em execução. Pare-a para iniciar outro áudio."
        Exit Function
    End If
    ' if cancel was pressed, exit the procedure
    If Err.Number = 32755 Then
        openFile = False
        Exit Function
    End If

    Call BASS_StreamFree(stream)
    Call BASS_MusicFree(stream)
    stream = BASS_StreamCreateFile(BASSFALSE, StrPtr(file), 0, 0, BASS_SAMPLE_LOOP)
    If stream = 0 Then stream = BASS_MusicLoad(BASSFALSE, file, 0, 0, BASS_MUSIC_RAMP Or BASS_MUSIC_LOOP, 0)

    If stream = 0 Then
        Call MsgBox("Selected file couldn't be played!")
        openFile = False ' Can't load the file
        Exit Function
    End If
    Call BASS_ChannelSetDevice(stream, device)
    If Not (nome Is Nothing) Then nome.Caption = name
    If Not (labelDur Is Nothing) Then labelDur.Caption = secToTimeString(BASS_ChannelBytes2Seconds(stream, BASS_ChannelGetLength(stream, BASS_POS_BYTE)))
    If Not (position Is Nothing) Then position.Caption = "00:00"
    If Not (rest Is Nothing) Then rest.Caption = labelDur.Caption
    If Not (led Is Nothing) Then led.BackColor = &HFF00&
    'desbilita o loop
    Call BASS_ChannelFlags(stream, 0, BASS_SAMPLE_LOOP)
    openFile = True
End Function

Function PlayFile(stream As Long)
    If BASS_ChannelPlay(stream, BASSFALSE) Then PlayFile = True Else PlayFile = False
End Function

Function stopFile(stream As Long)
    If BASS_ChannelStop(stream) And BASS_ChannelSetPosition(stream, 0, BASS_POS_BYTE) Then stopFile = True Else stopFile = False
End Function

Function pauseFile(stream As Long)
    If BASS_ChannelPause(stream) Then pauseFile = True Else pauseFile = False
End Function
Function releaseBass()
    Call BASS_Free
    releaseBass = True
End Function
Function closeFile(stream As Long)
    Call BASS_StreamFree(stream)
    Call BASS_MusicFree(stream)
End Function
Function setPosition(stream As Long, factor As Long)
Call BASS_ChannelSetPosition(stream, BASS_ChannelGetPosition(stream, BASS_POS_BYTE) + factor, BASS_POS_BYTE)
End Function

'atualiza o tempo da música
'COLOCANDO PARÂMETROS COMO  OPTIONAL PARA SIMULAR UMA SOBRECARGA DE MÉTODO
Sub updateTimes(stream As Long, Optional position As Object, Optional rest As Object, Optional mode As Object, Optional led As Object)
If BASS_ChannelIsActive(stream) = BASS_ACTIVE_PLAYING Then
    If Not (position Is Nothing) Then position.Caption = secToTimeString(BASS_ChannelBytes2Seconds(stream, BASS_ChannelGetPosition(stream, BASS_POS_BYTE)))
    If Not (rest Is Nothing) Then rest.Caption = secToTimeString(BASS_ChannelBytes2Seconds(stream, BASS_ChannelGetLength(stream, BASS_POS_BYTE) - BASS_ChannelGetPosition(stream, BASS_POS_BYTE)))
    If Not (mode Is Nothing) Then mode.Caption = "Reproduzindo"
    If Not (led Is Nothing) Then led.BackColor = &HFF&
Else
    If Not (mode Is Nothing) Then mode.Caption = "idle"
    'se led igual a vermelho então colorir de preto:
    If Not (led Is Nothing) Then If led.BackColor = &HFF& Then led.BackColor = &H0&
End If
End Sub
Sub setVolume(stream As Long, val As Long)
Call BASS_ChannelSetAttribute(stream, BASS_ATTRIB_VOL, (val / 10))
End Sub
Sub StartVu()
specpos = 0
specmode = 0

    ' create bitmap to draw spectrum in - 8 bit for easy updating :)
    With bh.bmiHeader
        .biBitCount = 8
        .biPlanes = 1
        .biSize = Len(bh.bmiHeader)
        .biWidth = SPECWIDTH
        .biHeight = SPECHEIGHT  ' upside down (line 0=bottom)
        .biClrUsed = 256
        .biClrImportant = 256
    End With

    Dim a As Byte

    ' setup palette
    For a = 1 To 127
        bh.bmiColors(a).rgbGreen = 256 - 2 * a
        bh.bmiColors(a).rgbRed = 2 * a
    Next a
    For a = 0 To 31
        bh.bmiColors(128 + a).rgbBlue = 8 * a
        bh.bmiColors(128 + 32 + a).rgbBlue = 255
        bh.bmiColors(128 + 32 + a).rgbRed = 8 * a
        bh.bmiColors(128 + 64 + a).rgbRed = 255
        bh.bmiColors(128 + 64 + a).rgbBlue = 8 * (31 - a)
        bh.bmiColors(128 + 64 + a).rgbGreen = 8 * a
        bh.bmiColors(128 + 96 + a).rgbRed = 255
        bh.bmiColors(128 + 96 + a).rgbGreen = 255
        bh.bmiColors(128 + 96 + a).rgbBlue = 8 * a
    Next a

    ' setup update timer (40hz)
#If 1 Then
    Form1.tmrSpectrum.Enabled = True
#Else
    timing = timeSetEvent(25, 25, AddressOf UpdateSpectrum, 0, TIME_PERIODIC)  ' API MM timer
#End If
End Sub

' converte ms para tempo
Public Function secToTimeString(ByRef nLength As Long) As String
Dim nSeconds As Long
Dim nMinutes As Long
    On Error GoTo error
    nSeconds = nLength Mod 60
    nMinutes = nLength \ 60
    secToTimeString = Format(nMinutes, "00") & ":" & Format(nSeconds, "00")
    nSeconds = 0
    nMinutes = 0
    Exit Function
error:
    MsgBox "Erro no calculo do tempo do áudio", vbExclamation, "Erro!"
End Function


