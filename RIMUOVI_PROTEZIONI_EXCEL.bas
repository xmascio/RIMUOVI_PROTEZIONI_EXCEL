Attribute VB_Name = "modRimuoviProtezioni"
Option Explicit

Private Const TOOL_NOME As String = "RIMUOVI_PROTEZIONI_EXCEL"
Private Const TOOL_VERSIONE As String = "1.1"
Private Const CARTELLA_LOG As String = "\\nas01\QUALITA\6-AUTOMAZIONI\RIMUOVI_PROTEZIONI_EXCEL\LOG"
Private Const FILE_LOG As String = "utilizzo_tool.csv"

' Macro principale:
' - rende visibili tutti i fogli, inclusi quelli xlSheetVeryHidden
' - rimuove la protezione da cartella di lavoro, finestre e struttura
' - rimuove la protezione da tutti i fogli
' - chiede la password quando necessaria
' - produce un report finale
'
' Nota: la protezione del progetto VBA non puo' essere rimossa in modo
' affidabile da una macro VBA. Va rimossa manualmente da:
' Editor VBA > Strumenti > Proprieta' di VBAProject > Protezione.

Public Sub RimuoviTutteLeProtezioni()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pwd As String
    Dim report As String
    Dim tentativi As Long
    Dim esito As Boolean
    Dim fogliProtetti As Collection
    Dim i As Long
    Dim avvio As Date
    Dim fogliProtettiIniziali As Long
    Dim fogliResiVisibili As Long
    Dim fogliProtettiFinali As Long
    Dim workbookProtettaIniziale As Boolean
    Dim esitoEsecuzione As String
    Dim erroreDettaglio As String

    avvio = Now
    Set wb = ActiveWorkbook
    Set fogliProtetti = New Collection

    If wb Is Nothing Then
        RegistraEsecuzione avvio, vbNullString, 0, 0, 0, 0, False, "ERRORE", "Nessuna cartella di lavoro attiva"
        MsgBox "Nessuna cartella di lavoro attiva.", vbExclamation, "Rimozione protezioni"
        Exit Sub
    End If

    workbookProtettaIniziale = wb.ProtectStructure Or wb.ProtectWindows
    For Each ws In wb.Worksheets
        If FoglioProtetto(ws) Then fogliProtettiIniziali = fogliProtettiIniziali + 1
    Next ws

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo FineConErrore

    report = "Report rimozione protezioni:" & vbCrLf & vbCrLf

    ' 1) Prova a rimuovere la protezione della cartella senza password.
    ' La protezione struttura impedisce anche di rendere visibili i fogli nascosti.
    esito = ProvaSproteggiWorkbook(wb, vbNullString)

    ' 2) Se la cartella e' ancora protetta, chiede password fino a 3 volte.
    If wb.ProtectStructure Or wb.ProtectWindows Then
        For tentativi = 1 To 3
            pwd = InputBox( _
                "La cartella di lavoro e' protetta." & vbCrLf & _
                "Inserisci la password di sblocco.", _
                "Password cartella di lavoro")

            If Len(pwd) = 0 Then Exit For

            esito = ProvaSproteggiWorkbook(wb, pwd)
            If Not (wb.ProtectStructure Or wb.ProtectWindows) Then Exit For

            MsgBox "Password non valida per la cartella di lavoro.", vbExclamation, "Rimozione protezioni"
        Next tentativi
    End If

    If wb.ProtectStructure Or wb.ProtectWindows Then
        report = report & "- Cartella di lavoro: protezione NON rimossa." & vbCrLf
    Else
        report = report & "- Cartella di lavoro: protezione rimossa o non presente." & vbCrLf
    End If

    ' 3) Rende visualizzabili tutti i fogli, se la struttura lo consente.
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            If ProvaRendiVisibileFoglio(ws) Then
                fogliResiVisibili = fogliResiVisibili + 1
                report = report & "- Foglio reso visibile: " & ws.Name & vbCrLf
            Else
                report = report & "- Foglio '" & ws.Name & "': NON reso visibile. Verificare protezione struttura cartella." & vbCrLf
            End If
        End If
    Next ws

    ' 4) Identifica i fogli ancora protetti.
    For Each ws In wb.Worksheets
        If FoglioProtetto(ws) Then
            fogliProtetti.Add ws
        End If
    Next ws

    ' 5) Prova a sproteggere i fogli senza password.
    For i = fogliProtetti.Count To 1 Step -1
        Set ws = fogliProtetti(i)
        Call ProvaSproteggiFoglio(ws, vbNullString)
    Next i

    ' 6) Per i fogli rimasti protetti chiede password fino a 3 volte.
    For Each ws In wb.Worksheets
        If FoglioProtetto(ws) Then
            For tentativi = 1 To 3
                pwd = InputBox( _
                    "Il foglio '" & ws.Name & "' e' protetto." & vbCrLf & _
                    "Inserisci la password di sblocco.", _
                    "Password foglio")

                If Len(pwd) = 0 Then Exit For

                Call ProvaSproteggiFoglio(ws, pwd)
                If Not FoglioProtetto(ws) Then Exit For

                MsgBox "Password non valida per il foglio '" & ws.Name & "'.", vbExclamation, "Rimozione protezioni"
            Next tentativi
        End If

        If FoglioProtetto(ws) Then
            fogliProtettiFinali = fogliProtettiFinali + 1
            report = report & "- Foglio '" & ws.Name & "': protezione NON rimossa." & vbCrLf
        Else
            report = report & "- Foglio '" & ws.Name & "': protezione rimossa o non presente." & vbCrLf
        End If
    Next ws

    report = report & vbCrLf & _
        "Nota: se il progetto VBA e' protetto, la relativa password va rimossa manualmente dall'Editor VBA."

Fine:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Len(erroreDettaglio) > 0 Then
        esitoEsecuzione = "ERRORE"
    ElseIf wb.ProtectStructure Or wb.ProtectWindows Or fogliProtettiFinali > 0 Then
        esitoEsecuzione = "PARZIALE"
    Else
        esitoEsecuzione = "OK"
    End If

    RegistraEsecuzione avvio, wb.Name, wb.Worksheets.Count, fogliProtettiIniziali, _
        fogliProtettiIniziali - fogliProtettiFinali, fogliResiVisibili, _
        workbookProtettaIniziale, esitoEsecuzione, erroreDettaglio

    MsgBox report, vbInformation, "Rimozione protezioni completata"
    Exit Sub

FineConErrore:
    erroreDettaglio = Err.Description
    report = report & vbCrLf & "Errore: " & erroreDettaglio
    Resume Fine
End Sub

' Registra una sola riga per esecuzione in un CSV condiviso, pronto per essere
' letto da una futura dashboard. Eventuali problemi di rete non interrompono il tool.
Private Sub RegistraEsecuzione( _
    ByVal avvio As Date, _
    ByVal nomeWorkbook As String, _
    ByVal fogliTotali As Long, _
    ByVal fogliProtetti As Long, _
    ByVal fogliSbloccati As Long, _
    ByVal fogliResiVisibili As Long, _
    ByVal workbookProtetta As Boolean, _
    ByVal esito As String, _
    ByVal dettaglioErrore As String)

    Dim percorso As String
    Dim numeroFile As Integer
    Dim nuovoFile As Boolean
    Dim durataSecondi As Long
    Dim fileAperto As Boolean

    On Error GoTo Uscita

    If Dir$(CARTELLA_LOG, vbDirectory) = vbNullString Then MkDir CARTELLA_LOG
    percorso = CARTELLA_LOG & "\" & FILE_LOG
    nuovoFile = (Dir$(percorso) = vbNullString)
    durataSecondi = DateDiff("s", avvio, Now)

    numeroFile = FreeFile
    Open percorso For Append As #numeroFile
    fileAperto = True

    If nuovoFile Then
        Print #numeroFile, "timestamp;tool;versione;utente;computer;workbook;durata_secondi;fogli_totali;fogli_protetti_iniziali;fogli_sbloccati;fogli_resi_visibili;workbook_protetta_iniziale;esito;errore"
    End If

    Print #numeroFile, CampoCsv(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ";" & _
        CampoCsv(TOOL_NOME) & ";" & CampoCsv(TOOL_VERSIONE) & ";" & _
        CampoCsv(Environ$("USERNAME")) & ";" & CampoCsv(Environ$("COMPUTERNAME")) & ";" & _
        CampoCsv(nomeWorkbook) & ";" & CStr(durataSecondi) & ";" & _
        CStr(fogliTotali) & ";" & _
        CStr(fogliProtetti) & ";" & CStr(fogliSbloccati) & ";" & CStr(fogliResiVisibili) & ";" & _
        CampoCsv(IIf(workbookProtetta, "SI", "NO")) & ";" & CampoCsv(esito) & ";" & _
        CampoCsv(dettaglioErrore)

Uscita:
    If fileAperto Then Close #numeroFile
End Sub

Private Function CampoCsv(ByVal valore As String) As String
    valore = Replace(valore, vbCr, " ")
    valore = Replace(valore, vbLf, " ")
    CampoCsv = Chr$(34) & Replace(valore, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function ProvaSproteggiWorkbook(ByVal wb As Workbook, ByVal pwd As String) As Boolean
    On Error Resume Next
    Err.Clear
    wb.Unprotect Password:=pwd
    ProvaSproteggiWorkbook = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ProvaSproteggiFoglio(ByVal ws As Worksheet, ByVal pwd As String) As Boolean
    On Error Resume Next
    Err.Clear
    ws.Unprotect Password:=pwd
    ProvaSproteggiFoglio = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FoglioProtetto(ByVal ws As Worksheet) As Boolean
    FoglioProtetto = _
        ws.ProtectContents Or _
        ws.ProtectDrawingObjects Or _
        ws.ProtectScenarios
End Function

Private Function ProvaRendiVisibileFoglio(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    Err.Clear
    ws.Visible = xlSheetVisible
    ProvaRendiVisibileFoglio = (Err.Number = 0 And ws.Visible = xlSheetVisible)
    Err.Clear
    On Error GoTo 0
End Function
