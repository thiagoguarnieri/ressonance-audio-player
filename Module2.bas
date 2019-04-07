Attribute VB_Name = "modPlaylist"
Type audioItem
    nome As String
    path As String
    length As Long
End Type

Type vinhetas
    nome As String
    caminho As String
    id As Long
    info As BASS_CHANNELINFO
End Type

Dim streamEvent As audioItem

Public Sub toPlaylist(posicaov As Integer, pauseItem As Integer, path As String, name As String, playlist() As audioItem, length As Long, lista As ListView, pref As String)
If (pauseItem = -1) Then
    ReDim Preserve playlist(UBound(playlist) + 1) 'aumenta a lista em 1 posição
    If lista.ListItems.Count > 0 Then
        Call deslocaVetor(posicaov, playlist)  'occurs only if playlist is not empty
        playlist(posicaov).path = path
        playlist(posicaov).nome = name
        playlist(posicaov).length = 0
        lista.ListItems.Add posicaov, , playlist(posicaov).nome
    Else
        playlist(posicaov).path = path
        playlist(posicaov).nome = name
        playlist(posicaov).length = 0
        lista.ListItems.Add posicaov, , playlist(posicaov).nome
    End If
Else
    Dim tempStr As Long
    ReDim Preserve playlist(UBound(playlist) + 1) 'aumenta a lista em 1 posição
    'carregar temporariamente para calcula duração do audio
    Call BASS_StreamFree(tempStr)
    Call BASS_MusicFree(tempStr)
    tempStr = BASS_StreamCreateFile(BASSFALSE, StrPtr(path & "\" & name), 0, 0, BASS_SAMPLE_LOOP)
    If tempStr = 0 Then tempStr = BASS_MusicLoad(BASSFALSE, file, 0, 0, BASS_MUSIC_RAMP Or BASS_MUSIC_LOOP, 0)
    If lista.ListItems.Count > 0 Then 'se a lista não está vazia
        Call deslocaVetor(posicaov, playlist)  'occurs only if playlist is not empty
        playlist(posicaov).path = path
        playlist(posicaov).nome = name
        playlist(posicaov).length = (BASS_ChannelBytes2Seconds(tempStr, BASS_ChannelGetLength(tempStr, BASS_POS_BYTE)))
        length = length + playlist(posicaov).length
        Form1.LabelCalcBloco.Caption = secToTimeString(length)
        lista.ListItems.Add posicaov, , pref & "[" & secToTimeString(playlist(posicaov).length) & "] " & playlist(posicaov).nome
    Else 'se a lista está vazia
        playlist(posicaov).path = path
        playlist(posicaov).nome = name
        playlist(posicaov).length = (BASS_ChannelBytes2Seconds(tempStr, BASS_ChannelGetLength(tempStr, BASS_POS_BYTE)))
        length = length + playlist(posicaov).length
        Form1.LabelCalcBloco.Caption = secToTimeString(length)
        lista.ListItems.Add posicaov, , pref & "[" & secToTimeString(playlist(posicaov).length) & "] " & playlist(posicaov).nome
    End If
    Call BASS_StreamFree(tempStr)
    Call BASS_MusicFree(tempStr)
End If
End Sub

'desloca parte do vetor pra frente para
'liberar determinada posição para inserção
Public Sub deslocaVetor(ByRef posicao As Integer, playlist() As audioItem)
    For Y = UBound(playlist) To posicao Step -1
        playlist(Y) = playlist(Y - 1)
    Next Y
End Sub

Public Function isAutocross(ByRef stream As Long, ByRef crosspoint As Long) As Boolean
Dim position As Integer
position = BASS_ChannelBytes2Seconds(stream, BASS_ChannelGetPosition(stream, BASS_POS_BYTE)) 'vai dar -1 se naum tiver carregado
If BASS_ChannelIsActive(stream) = BASS_ACTIVE_PLAYING Or position > 0 Then
    If position >= crosspoint Then
       crosspoint = crosspoint + crosspoint 'para esta função ficar false
       isAutocross = True
    Else
       isAutocross = False
    End If
End If
End Function

Public Sub SetCrossPt(ByRef stream As Long, ByRef crosspoint As Long)
Dim ponto As Long
ponto = BASS_ChannelBytes2Seconds(stream, BASS_ChannelGetLength(stream, BASS_POS_BYTE))
If ponto >= 90 Then
crosspoint = ponto - 5
Else: crosspoint = ponto - 0.3
End If
End Sub
Public Sub limparPlaylist(playlist() As audioItem, length As Long, lista As ListView)
If lista.ListItems.Count > 0 Then
    ReDim playlist(0)
    lista.ListItems.Clear
    length = 0
    Form1.LabelCalcBloco.Caption = secToTimeString(length)
End If
End Sub

Public Sub apagar(ByRef item As Integer, playlist() As audioItem, length As Long, lista As ListView)
If lista.ListItems.Count > 0 Then
    If length > 0 Then length = length - playlist(item).length Else length = 0
    For Y = item To (lista.ListItems.Count - 1)
        playlist(Y) = playlist(Y + 1)
    Next Y
    ReDim Preserve playlist(UBound(playlist) - 1)
    lista.ListItems.Remove (item)
    Form1.LabelCalcBloco.Caption = secToTimeString(length)
End If
End Sub
Public Sub setInfoEvent(audioEvent As audioItem)
streamEvent = audioEvent
End Sub

Public Function getInfoEvent() As audioItem
getInfoEvent = streamEvent
End Function

'lida com a movimentação de itens no vetor
Public Sub permutaVetor(ByRef posicao As Integer, ByRef mode As Integer, playlist() As audioItem)
    Dim eventoapoio As audioItem
    If mode = 0 Then 'uplist
        eventoapoio = playlist(posicao - 1)
        playlist(posicao - 1) = playlist(posicao)
        playlist(posicao) = eventoapoio
    ElseIf mode = 1 Then 'downlist
        eventoapoio = playlist(posicao + 1)
        playlist(posicao + 1) = playlist(posicao)
        playlist(posicao) = eventoapoio
    End If
End Sub


