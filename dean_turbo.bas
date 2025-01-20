Attribute VB_Name = "dean_turbo"

'
Option Explicit

' ********************************************************************************
' 1) Startpunkt: Datei-Dialog für die "Anonymisierte_Daten.xlsx"
' ********************************************************************************
Public Sub StartDeAnonymisierung_Turbo()
    Dim fd As FileDialog
    Dim FilePath As String, FolderPath As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Bitte die anonymisierte XLSX-Datei auswählen"
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xlsx; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
            FolderPath = Left(FilePath, InStrRev(FilePath, "\"))
            Call HauptprozedurDeanonymisierung_Turbo(FilePath, FolderPath)
        Else
            MsgBox "Keine Datei ausgewählt.", vbExclamation
        End If
    End With
End Sub

' ********************************************************************************
' 2) Hauptprozedur: Öffnet die anonymisierte Datei, ruft DeAnonymisieren auf
' ********************************************************************************
Public Sub HauptprozedurDeanonymisierung_Turbo(ByVal FilePath As String, ByVal FolderPath As String)
    Dim wbAnon As Workbook
    
    ' Performance-Settings sichern
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean, oldScreenUpdating As Boolean
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    On Error GoTo Fehler
    
    ' Performance-Schalter
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 2a) Anonymisierte Mappe öffnen
    Set wbAnon = Workbooks.Open(FilePath)
    
    ' 2b) De-Anonymisieren
    DeAnonymisieren_Turbo wbAnon, FolderPath
    
    ' 2c) Datei schließen (nicht überschreiben)
    wbAnon.Close SaveChanges:=False
    
    MsgBox "De-Anonymisierung (Turbo) abgeschlossen.", vbInformation
    
Aufraeumen:
    ' Ursprungszustand wiederherstellen
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
    
Fehler:
    MsgBox "Fehler: " & Err.Description, vbCritical
    Resume Aufraeumen
End Sub

' ********************************************************************************
' 3) Turbo-DeAnonymisierung: nur 2 Schleifen statt pro Spalte
'    - Erst GZ/GZ Neu/Kommune
'    - Dann Personenfelder via nr2
' ********************************************************************************
Private Sub DeAnonymisieren_Turbo(ByVal wbAnon As Workbook, ByVal FolderPath As String)
    Dim wsQuelle As Worksheet
    Dim lastRow As Long, LastCol As Long
    
    Dim dataAnon As Variant, dataOut As Variant
    
    ' --- Dictionary für GZ / Kommune ---
    Dim dictGZ As Object
    Dim dictKommune As Object
    
    ' --- Personen-Zuordnung ---
    Dim dataZuordNr2 As Variant
    Dim dictNr2Row As Object
    
    ' 1) Basis-Infos ermitteln
    Set wsQuelle = wbAnon.Sheets(1)
    lastRow = wsQuelle.Cells(wsQuelle.Rows.Count, 1).End(xlUp).Row
    LastCol = wsQuelle.Cells(1, wsQuelle.Columns.Count).End(xlToLeft).Column
    
    ' 2) Kompletten Bereich in dataAnon lesen
    dataAnon = wsQuelle.Range(wsQuelle.Cells(1, 1), wsQuelle.Cells(lastRow, LastCol)).Value
    
    ' 3) dataOut = dataAnon (Blockkopie)
    '    -> Alle "unveränderten" Spalten sind direkt mitkopiert.
    dataOut = dataAnon
    
    ' 4) Dictionaries laden
    Set dictGZ = LadeZuordnungInDictionary(FolderPath & "Zuordnung_GZ.xlsx")
    Set dictKommune = LadeZuordnungInDictionary(FolderPath & "Zuordnung_Kommune.xlsx")
    
    ' 5) Zuordnung_Nr2 & Dictionary
    dataZuordNr2 = LadeZuordnungNr2(FolderPath & "Zuordnung_Nr2.xlsx")
    Set dictNr2Row = BaueDictNr2Index(dataZuordNr2)
    
    ' 6) Spaltenindizes in der anonymisierten Datei (z.B. "GZ", "nr2", "Name", usw.)
    Dim col_nr2 As Long, col_gz As Long, col_gzNeu As Long, col_kommune As Long
    Dim col_name As Long, col_vorname As Long
    Dim col_gebdat As Long, col_strasse As Long, col_hnr As Long, col_plz As Long
    
    col_nr2 = FindColumnIndex(dataAnon, "nr2")
    col_gz = FindColumnIndex(dataAnon, "GZ")
    col_gzNeu = FindColumnIndex(dataAnon, "GZ Neu")
    col_kommune = FindColumnIndex(dataAnon, "Kommune")
    
    col_name = FindColumnIndex(dataAnon, "Name")
    col_vorname = FindColumnIndex(dataAnon, "Vorname")
    col_gebdat = FindColumnIndex(dataAnon, "Geb.Dat.")
    col_strasse = FindColumnIndex(dataAnon, "Straße")
    col_hnr = FindColumnIndex(dataAnon, "Hausnummer")
    col_plz = FindColumnIndex(dataAnon, "PLZ")
    
    ' 6b) Spaltenindizes im Zuordnungs-Array (Originalwerte)
    Dim idxNr2 As Long, idxName As Long, idxVorname As Long
    Dim idxGebDat As Long, idxStrasse As Long, idxHnr As Long, idxPlz As Long
    
    idxNr2 = FindColumnIndex(dataZuordNr2, "nr2")
    idxName = FindColumnIndex(dataZuordNr2, "Name")
    idxVorname = FindColumnIndex(dataZuordNr2, "Vorname")
    idxGebDat = FindColumnIndex(dataZuordNr2, "Geb.Dat.")
    idxStrasse = FindColumnIndex(dataZuordNr2, "Straße")
    idxHnr = FindColumnIndex(dataZuordNr2, "Hausnummer")
    idxPlz = FindColumnIndex(dataZuordNr2, "PLZ")
    
    ' ------------------------------------------------------------------
    ' 7) Erste Schleife für GZ, GZ Neu, Kommune
    ' ------------------------------------------------------------------
    Dim i As Long
    For i = 2 To lastRow
        If i Mod 2000 = 0 Then DoEvents
        
        ' GZ
        If col_gz > 0 Then
            dataOut(i, col_gz) = DeanonymisiereAusDictionary( _
                CStr(dataAnon(i, col_gz)), dictGZ)
        End If
        
        ' GZ Neu
        If col_gzNeu > 0 Then
            dataOut(i, col_gzNeu) = DeanonymisiereAusDictionary( _
                CStr(dataAnon(i, col_gzNeu)), dictGZ)
        End If
        
        ' Kommune
        If col_kommune > 0 Then
            dataOut(i, col_kommune) = DeanonymisiereAusDictionary( _
                CStr(dataAnon(i, col_kommune)), dictKommune)
        End If
    Next i
    
    ' ------------------------------------------------------------------
    ' 8) Zweite Schleife für personenbezogene Daten (Name, Vorname, ...)
    '    Hier nutzen wir "nr2" als Schlüssel, NICHT den "xxx_ANON_xx"-String.
    '    => 1x Dictionary-Lookup pro Zeile statt 7x!
    ' ------------------------------------------------------------------
    If col_nr2 > 0 And idxNr2 > 0 And Not (dictNr2Row Is Nothing) Then
        For i = 2 To lastRow
            If i Mod 2000 = 0 Then DoEvents
            
            Dim sNr2 As String
            sNr2 = CStr(dataAnon(i, col_nr2))  ' ID in nr2-Spalte
            
            If dictNr2Row.Exists(sNr2) Then
                Dim rowZuord As Long
                rowZuord = dictNr2Row(sNr2)
                
                ' Name
                If col_name > 0 And idxName > 0 Then
                    dataOut(i, col_name) = dataZuordNr2(rowZuord, idxName)
                End If
                
                ' Vorname
                If col_vorname > 0 And idxVorname > 0 Then
                    dataOut(i, col_vorname) = dataZuordNr2(rowZuord, idxVorname)
                End If
                
                ' Geb.Dat.
                If col_gebdat > 0 And idxGebDat > 0 Then
                    dataOut(i, col_gebdat) = dataZuordNr2(rowZuord, idxGebDat)
                End If
                
                ' Straße
                If col_strasse > 0 And idxStrasse > 0 Then
                    dataOut(i, col_strasse) = dataZuordNr2(rowZuord, idxStrasse)
                End If
                
                ' Hausnummer
                If col_hnr > 0 And idxHnr > 0 Then
                    dataOut(i, col_hnr) = dataZuordNr2(rowZuord, idxHnr)
                End If
                
                ' PLZ
                If col_plz > 0 And idxPlz > 0 Then
                    dataOut(i, col_plz) = dataZuordNr2(rowZuord, idxPlz)
                End If
            End If
        Next i
    End If
    
    ' 9) Ergebnis in neues Workbook "Rekonstruierte_Daten.xlsx"
    Dim wbResult As Workbook, wsResult As Worksheet
    Set wbResult = Workbooks.Add
    Set wsResult = wbResult.Sheets(1)
    
    wsResult.Range(wsResult.Cells(1, 1), wsResult.Cells(lastRow, LastCol)).Value = dataOut
    
    wbResult.SaveAs FileName:=FolderPath & "Rekonstruierte_Daten.xlsx"
    wbResult.Close False
End Sub

' ********************************************************************************
' Hilfsfunktionen
' ********************************************************************************

' ------------------------------------------------------------------------------
' 2-Spalten-Zuordnung laden (Original | Anonymisiert),
' für die DE-Anonymisierung braucht man: dict(Anonymisiert) = Original
' ------------------------------------------------------------------------------
Private Function LadeZuordnungInDictionary(ByVal xlsPath As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If Dir(xlsPath) = "" Then
        Set LadeZuordnungInDictionary = dict
        Exit Function
    End If
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(xlsPath, ReadOnly:=True)
    Set ws = wb.Sheets(1)
    
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lr
        Dim originalVal As String, anonVal As String
        originalVal = CStr(ws.Cells(r, 1).Value)
        anonVal = CStr(ws.Cells(r, 2).Value)
        If anonVal <> "" Then
            dict(anonVal) = originalVal
        End If
    Next r
    
    wb.Close False
    Set LadeZuordnungInDictionary = dict
End Function

' ------------------------------------------------------------------------------
' Lädt Zuordnung_Nr2.xlsx (nr2, Name, Vorname, Geb.Dat., Straße, Hausnummer, PLZ)
' ------------------------------------------------------------------------------
Private Function LadeZuordnungNr2(ByVal xlsPath As String) As Variant
    If Dir(xlsPath) = "" Then
        Dim dummy(1 To 1, 1 To 1)
        dummy(1, 1) = "Keine Zuordnung_Nr2.xlsx"
        LadeZuordnungNr2 = dummy
        Exit Function
    End If
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(xlsPath, ReadOnly:=True)
    Set ws = wb.Sheets(1)
    
    Dim lr As Long, lc As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim ret As Variant
    ret = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
    
    wb.Close False
    LadeZuordnungNr2 = ret
End Function

' ------------------------------------------------------------------------------
' Dictionary: dictNr2Row("42") = 10  -> Zeile 10 in dataZuordNr2
' ------------------------------------------------------------------------------
Private Function BaueDictNr2Index(ByVal dataZuordNr2 As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim col_nr2 As Long
    col_nr2 = FindColumnIndex(dataZuordNr2, "nr2")
    If col_nr2 < 1 Then
        Set BaueDictNr2Index = dict
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = UBound(dataZuordNr2, 1)
    
    Dim r As Long
    For r = 2 To lastRow
        Dim keyVal As Variant
        keyVal = dataZuordNr2(r, col_nr2)
        If Not IsEmpty(keyVal) And keyVal <> "" Then
            dict(CStr(keyVal)) = r
        End If
    Next r
    
    Set BaueDictNr2Index = dict
End Function

' ------------------------------------------------------------------------------
' Sucht Spaltenindex in data(1, j) (Zeile 1), gibt -1 zurück, wenn nicht gefunden
' ------------------------------------------------------------------------------
Private Function FindColumnIndex(ByVal dataIn As Variant, ByVal colName As String) As Long
    If Not IsArray(dataIn) Then
        FindColumnIndex = -1
        Exit Function
    End If
    
    Dim j As Long
    For j = LBound(dataIn, 2) To UBound(dataIn, 2)
        If CStr(dataIn(1, j)) = colName Then
            FindColumnIndex = j
            Exit Function
        End If
    Next j
    FindColumnIndex = -1
End Function

' ------------------------------------------------------------------------------
' Dictionary-Lookup: z.B. dict("GZ_ANON_1") = "ABC123"
' ------------------------------------------------------------------------------
Private Function DeanonymisiereAusDictionary(ByVal anonKey As String, ByVal dict As Object) As String
    If dict Is Nothing Then
        DeanonymisiereAusDictionary = anonKey
    ElseIf dict.Exists(anonKey) Then
        DeanonymisiereAusDictionary = dict(anonKey)
    Else
        DeanonymisiereAusDictionary = anonKey
    End If
End Function


