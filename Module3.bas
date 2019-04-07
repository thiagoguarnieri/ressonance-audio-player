Attribute VB_Name = "modIOoperations"
'para usar api do windows no vb, deve se criar as estruturas (types) correspondentes aos parametro de entrada
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_READONLY = &H1
Const ERROR_NO_MORE_FILES = 18&
Const FILE_ATTRIBUTE_NORMAL = &H80

Const INVALID_HANDLE_VALUE As Long = -1
Const MAX_PATH As Integer = 260

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Type comerciais
    horario As String
    mode As String
    fldrPath As String
    arquivos() As audioItem
End Type

Type avisos
    aviso As String
    horario As String
End Type
'varre diretório em busca de arquivos e carrega vertor de comerciais
'É FEITO UM DIRETORIO POR VEZ, SENDO QUE SO A PARTE RELACIONADA AOS ITENS É POPULADA
'A PARTE RELACIONADA A DATA E HORA É POPULADA EM OUTRO LUGAR
Public Sub scanFolder(folder As String, ext As String, playlist() As audioItem, tamanhobloco As Long, pref As String)
On Error Resume Next 'para evitar erros pela falta da rede
Dim iSearchHandle As Long
' File search buffer
Dim pFindFileBuff As WIN32_FIND_DATA
Dim toklist As String
toklist = folder & ext
Dim cont As Integer
cont = 0
iSearchHandle = FindFirstFile(toklist, pFindFileBuff)
If iSearchHandle <> INVALID_HANDLE_VALUE Then
    If ext <> "*.lst" Then
        Call toPlaylist(1, 0, folder, TrimNull(pFindFileBuff.cFileName), playlist, tamanhobloco, Form1.playlistView, pref)
        Do While FindNextFile(iSearchHandle, pFindFileBuff)
            Call toPlaylist(1, 0, folder, TrimNull(pFindFileBuff.cFileName), playlist, tamanhobloco, Form1.playlistView, pref)
            DoEvents
        Loop
        Call FindClose(iSearchHandle)
    Else
        Call openfileList(folder & "\" & TrimNull(pFindFileBuff.cFileName), playlist, tamanhobloco)
        Call FindClose(iSearchHandle)
    End If
End If
End Sub

'recupera nome de arquivos
Function TrimNull(sFileName As String) As String
    Dim i As Long
    ' Search for the first null character
    i = InStr(1, sFileName, vbNullChar)
    If i = 0 Then
        TrimNull = sFileName
    Else
        ' Return the file name
        TrimNull = Left$(sFileName, i - 1)
    End If
End Function

Public Sub openfileCom(file As String, listaComercial() As comerciais)
On Error GoTo fileerror
    ReDim listaComercial(0)
    Form1.arvoreComercial.Nodes.Clear
    Form1.arvoreComercial.Nodes.Add , , "root", "Raiz"
    Form1.arvoreComercial.Nodes("root").Expanded = True
    Open file For Input As #2
    Do Until EOF(2)
        ReDim Preserve listaComercial(UBound(listaComercial) + 1)
        Line Input #2, listaComercial(UBound(listaComercial)).fldrPath
        listaComercial(UBound(listaComercial)).fldrPath = listaComercial(UBound(listaComercial)).fldrPath & "\"
        Line Input #2, listaComercial(UBound(listaComercial)).mode
        Line Input #2, listaComercial(UBound(listaComercial)).horario
        Form1.arvoreComercial.Nodes.Add "root", tvwChild, listaComercial(UBound(listaComercial)).fldrPath & UBound(listaComercial), listaComercial(UBound(listaComercial)).fldrPath & " - (" & listaComercial(UBound(listaComercial)).horario & ")"
    Loop
    Close #2
    Form1.textoLog.Text = Form1.textoLog.Text & vbCrLf & "Carregamento dos comerciais de: " & file & " com sucesso"
    Exit Sub
fileerror:
   Form1.textoLog.Text = Form1.textoLog.Text & vbCrLf & "Erro na leitura do arquivo de comerciais"
   Close #2
   ReDim listaComercial(0)
   Form1.arvoreComercial.Nodes.Clear
End Sub
Public Sub savefileList(ByRef file As String, playlist() As audioItem)
On Error GoTo fileerror
Open file For Output As #3
For c = 1 To UBound(playlist) Step 1
    Print #3, playlist(c).nome
    Print #3, playlist(c).path
Next c
Close #3
Exit Sub
fileerror:
    MsgBox "Erro na gravação da lista de músicas"
    Close #3
End Sub
Public Sub openfileList(ByRef file As String, playlist() As audioItem, length As Long)
On Error GoTo fileerror
    Dim arquivo As String
    Dim tempStr As Long
    Open file For Input As #4
    Do Until EOF(4)
        ReDim Preserve playlist(UBound(playlist) + 1)
        Line Input #4, playlist(UBound(playlist)).nome
        Line Input #4, playlist(UBound(playlist)).path
        
        'Cálculo do tempo das músicas:
        If Not (playlist(UBound(playlist)).nome Like "PAUSA") And Not (playlist(UBound(playlist)).nome Like "HORACERTA") Then
            Call BASS_StreamFree(tempStr)
            Call BASS_MusicFree(tempStr)
            arquivo = playlist(UBound(playlist)).path & "\" & playlist(UBound(playlist)).nome
            tempStr = BASS_StreamCreateFile(BASSFALSE, StrPtr(arquivo), 0, 0, BASS_SAMPLE_LOOP)
            If tempStr = 0 Then tempStr = BASS_MusicLoad(BASSFALSE, arquivo, 0, 0, BASS_MUSIC_RAMP Or BASS_MUSIC_LOOP, 0)
            playlist(UBound(playlist)).length = (BASS_ChannelBytes2Seconds(tempStr, BASS_ChannelGetLength(tempStr, BASS_POS_BYTE)))
        Else
            playlist(UBound(playlist)).length = 0
        End If
        length = length + playlist(UBound(playlist)).length
        Form1.LabelCalcBloco.Caption = secToTimeString(length)
        'Adiciona no objeto Playlistview
        Form1.playlistView.ListItems.Add , , playlist(UBound(playlist)).nome
    Loop
    Close #4
    Exit Sub
fileerror:
    MsgBox "Erro na leitura do arquivo de playlist" & vbCrLf & "Verifique se o diretório dos áudios existe"
    Close #4
    ReDim playlist(0)
    Form1.playlistView.ListItems.Clear
End Sub
Public Sub abrirAvisos(listadeAvisos() As avisos, file As String)
On Error GoTo fileerror
ReDim listadeAvisos(0)
Open file For Input As #5
    Do Until EOF(5)
    ReDim Preserve listadeAvisos(UBound(listadeAvisos) + 1)
    Line Input #5, listadeAvisos(UBound(listadeAvisos)).aviso
    Line Input #5, listadeAvisos(UBound(listadeAvisos)).horario
    Loop
    Close #5
    Form1.textoLog.Text = Form1.textoLog.Text & vbCrLf & "Arquivo de avisos carregado de " & file & " com sucesso"
    Exit Sub
fileerror:
    Form1.textoLog.Text = Form1.textoLog.Text & vbCrLf & "Erro na leitura do arquivo de avisos"
    Close #5
    ReDim listadeAvisos(0)
End Sub
Public Sub openVinhetas(vinhetas() As vinhetas, filename As String)
On Error GoTo fileerror
Dim conter As Integer
conter = 0
Open filename For Input As #8
      Do Until EOF(8)
        Line Input #8, vinhetas(conter).nome
        Line Input #8, vinhetas(conter).caminho
        If Not (vinhetas(conter).nome Like "vazio") Then
            Call openFile(1, vinhetas(conter).caminho, vinhetas(conter).nome, vinhetas(conter).id, vinhetas(conter).info, Form1.lblVinInfo(conter), Form1.vinheta1(conter))
        End If
      conter = conter + 1
      Loop
      Close #8
Exit Sub
fileerror:
MsgBox "(openVinhetas) erro na abertura de vinhetas"
Close #8
End Sub
Public Sub saveVinhetas(vinhetas() As vinhetas, filename As String)
On Error GoTo error
    If Not (filename Like vbNullString) Then
        Open filename For Output As #7
        For X = 0 To 59
            If Not (vinhetas(X).nome Like "") Then
                Print #7, vinhetas(X).nome
                Print #7, vinhetas(X).caminho
            Else
                Print #7, "vazio"
                Print #7, "vazio"
            End If
        Next X
        Close #7
    End If
    Exit Sub
error:
    MsgBox "erro no salvamento das vinhetas" & vbCrLf & "arquivos remotos apenas por unidades mapeadas"
    Close #7
End Sub

Public Function findFiles(diretorio As String, extensao As String, strbusca As String) As Long
Dim iSearchHandle As Long
Dim pFindFileBuff As WIN32_FIND_DATA
Dim toklist As String
Dim newDir As String
toklist = diretorio & extensao
iSearchHandle = FindFirstFile(toklist, pFindFileBuff)
If iSearchHandle <> INVALID_HANDLE_VALUE Then
    Do
        DoEvents 'EVITA QUE A INTERFACE GRÁFICA TRAVE
        'se é um diretório
        If (pFindFileBuff.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
             'se naum é diretorio acima
            FolderName = TrimNull(pFindFileBuff.cFileName)
            If Not (FolderName Like ("..")) Then
                If Not (FolderName Like (".")) Then
                    Form1.lblStatusBusca.Caption = FolderName 'para dar um feedback
                    newDir = diretorio & FolderName & "\"
                    findFiles newDir, extensao, strbusca
                End If
            End If
        Else
            'se é um arquivo
                If (InStr(1, LCase$(TrimNull(pFindFileBuff.cFileName)), LCase$(strbusca), 0)) Then
                    ext = Right$(TrimNull(LCase$(pFindFileBuff.cFileName)), 4)
                    If ext Like ".mp3" Or ext Like ".wav" Then
                    Form1.ListViewBusca.ListItems.Add , diretorio & TrimNull(pFindFileBuff.cFileName), TrimNull(pFindFileBuff.cFileName)
                    End If
                End If
        End If
    Loop While FindNextFile(iSearchHandle, pFindFileBuff)
        Call FindClose(iSearchHandle)
        If Form1.ListViewBusca.ListItems.Count > 0 Then findFiles = 1
Else
    findFiles = 0
End If
End Function
'Salva em formato texto uma lista com os áudio executados
Public Sub salvaLog(ecadcounter As Integer)
    Dim ecadcontent As String
    Dim comcontent As String
    If ecadcounter = 1200 Then
        ecadcounter = 0
        If Form1.tocadaslist.ListCount > 0 Then
            On Error GoTo error
            Open "c:\ressonance\ressonancelog.doc" For Append As #12 'GRAVA SEM APAGAR O QUE JA TAVA
            Open "c:\ressonance\ressonancelogcom.doc" For Append As #13 'GRAVA SEM APAGAR O QUE JA TAVA
            For cont = 0 To (Form1.tocadaslist.ListCount - 1)
                If Left(Form1.tocadaslist.List(cont), 5) = "[EVT]" Then ' COM LIKE NÃO DEU CERTO
                    comcontent = comcontent & Form1.tocadaslist.List(cont) & vbCrLf
                    Else: ecadcontent = ecadcontent & Form1.tocadaslist.List(cont) & vbCrLf
                End If
            Next cont
            Print #12, ecadcontent
            Print #13, comcontent
            Form1.tocadaslist.Clear
            ecadcontent = vbNullString
            comcontent = vbNullString
            Close #12 'fecha o arquivo de ecad
            Close #13 'fecha o arquivo de comerciais
        End If
    End If
    Exit Sub
error:
    MsgBox "O relatório não pôde ser salvo", vbInformation, "Erro"
End Sub
Public Sub saveDevices(filename As String, placaPrinc As Long, placaCue As Long)
On Error GoTo error
    If Not (filename Like vbNullString) Then
        Open filename For Output As #14
            Print #14, placaPrinc
            Print #14, placaCue
        Close #14
    End If
    Exit Sub
error:
    MsgBox "erro no salvamento das configurações de placas"
    Close #14
End Sub
Public Sub OpenDevices(file As String, devPrinc As Long, devCue As Long)
On Error GoTo fileerror
    Dim princ, cue As String
    Open file For Input As #15
    Do Until EOF(15)
        Line Input #15, princ
        Line Input #15, cue
    Loop
    Close #15
    devPrinc = CLng(princ)
    devCue = CLng(cue)
    Form1.textoLog.Text = Form1.textoLog.Text & vbCrLf & "Carregamento dos dispositivos de: " & file & " com sucesso"
    Exit Sub
fileerror:
    Form1.textoLog.Text = Form1.textoLog.Text & vbCrLf & "Erro na abertura da configuração das placas de som"
    Close #15
End Sub

'mecanismo de percorrimento de espelho sem interrupção
Public Sub addNodes(diretorio As String, extensao As String, parent As String)
'On Error Resume Next 'para evitar erros pela falta da rede
Dim iSearchHandle As Long
Dim pFindFileBuff As WIN32_FIND_DATA
Dim toklist As String
Dim newDir As String
toklist = diretorio & extensao
iSearchHandle = FindFirstFile(toklist, pFindFileBuff)
If iSearchHandle <> INVALID_HANDLE_VALUE Then
    Do
        'se é um diretório
        If Not ((pFindFileBuff.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory) Then
            'se é um arquivo
                    ext = Right$(TrimNull(pFindFileBuff.cFileName), 4)
                    If ext Like ".mp3" Or ext Like ".wav" Then
                        Form1.Espelho.Nodes.Add parent, tvwChild, diretorio & "\" & TrimNull(pFindFileBuff.cFileName), TrimNull(pFindFileBuff.cFileName), 2
                    End If
        End If
        DoEvents
    Loop While FindNextFile(iSearchHandle, pFindFileBuff)
        Call FindClose(iSearchHandle)
End If
End Sub
'PARA CARREGAR VINHETAS ROTATIVAS
Public Sub readFolder(folder As String, ext As String, lista As ListView, botao As CommandButton)
On Error Resume Next 'para evitar erros pela falta da rede
Dim pathParts As Variant
Dim iSearchHandle As Long
' File search buffer
Dim pFindFileBuff As WIN32_FIND_DATA
Dim toklist As String
pathParts = Split(folder, "\")
botao.Caption = pathParts(UBound(pathParts))
toklist = folder & "\" & ext
iSearchHandle = FindFirstFile(toklist, pFindFileBuff)
If iSearchHandle <> INVALID_HANDLE_VALUE Then
    lista.ListItems.Add , folder & "\" & TrimNull(pFindFileBuff.cFileName), TrimNull(pFindFileBuff.cFileName)
    Do While FindNextFile(iSearchHandle, pFindFileBuff)
        lista.ListItems.Add , folder & "\" & TrimNull(pFindFileBuff.cFileName), TrimNull(pFindFileBuff.cFileName)
        DoEvents
    Loop
    Call FindClose(iSearchHandle)
End If
End Sub

