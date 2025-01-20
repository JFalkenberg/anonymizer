Attribute VB_Name = "anon_turbo"

Option Explicit

' ------------------------------------------------------------------------------
' Startet den Dateidialog und ruft anschlie�end die Hauptprozedur auf
' ------------------------------------------------------------------------------
Public Sub StartAnonymisierung()
    Dim fd As FileDialog
    Dim FilePath As String
    Dim FolderPath As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Bitte eine XLSX-Datei ausw�hlen"
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xlsx; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
            FolderPath = Left(FilePath, InStrRev(FilePath, "\"))
            
            Call Hauptprozedur(FilePath, FolderPath)
        Else
            MsgBox "Keine Datei ausgew�hlt.", vbExclamation
        End If
    End With
End Sub

' ------------------------------------------------------------------------------
' �ffnet die gew�hlte Datei, ruft die Anonymisierung auf und speichert Ergebnisse
' ------------------------------------------------------------------------------
Public Sub Hauptprozedur(ByVal FilePath As String, ByVal FolderPath As String)
    Dim wb As Workbook
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean, oldScreenUpdating As Boolean
    
    ' Einstellungen sichern
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    On Error GoTo Fehler
    
    ' Performance-Schalter
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Original-Arbeitsmappe �ffnen
    Set wb = Workbooks.Open(FilePath)
    
    ' Anonymisierung durchf�hren
    Call AnonymisierenMitZeilenID(wb, FolderPath)
    
    ' Original-Datei schlie�en (mitspeichern, damit nr2 �berschrieben ist)
    wb.Close SaveChanges:=True
    
    MsgBox "Anonymisierung abgeschlossen.", vbInformation
    
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

' ------------------------------------------------------------------------------
' Liest Daten blockweise ein, anonymisiert sie und erzeugt die Ergebnis-Dateien
' ------------------------------------------------------------------------------
Private Sub AnonymisierenMitZeilenID(ByVal wb As Workbook, ByVal FolderPath As String)
    Dim wsQuelle As Worksheet
    Dim lastRow As Long, LastCol As Long
    
    Dim dataIn As Variant        ' Array f�r Originaldaten
    Dim dataOut As Variant       ' Array f�r anonymisierte Daten
    Dim dataZuordnung As Variant ' Array f�r pers�nliche Zuordnungswerte ("nr2", Name, etc.)
    
    Dim dictGZ As Object, dictKommune As Object
    Dim i As Long, j As Long
    
    ' F�r die Spaltenindizes
    Dim col_nr2 As Long
    Dim col_name As Long, col_vorname As Long, col_gebdat As Long
    Dim col_strasse As Long, col_Hausnr As Long, col_plz As Long
    
    ' Tabellenverweis (erste Tabelle)
    Set wsQuelle = wb.Sheets(1)
    
    ' Letzte belegte Zeile/Spalte ermitteln
    lastRow = wsQuelle.Cells(wsQuelle.Rows.Count, 1).End(xlUp).Row
    LastCol = wsQuelle.Cells(1, wsQuelle.Columns.Count).End(xlToLeft).Column
    
    ' Originaldaten in Array einlesen
    dataIn = wsQuelle.Range(wsQuelle.Cells(1, 1), wsQuelle.Cells(lastRow, LastCol)).Value
    
    ' Ziel-Arrays erstellen
    ReDim dataOut(1 To lastRow, 1 To LastCol)
    
    ' F�r die Personen-Zuordnung (gleiche Zeilenanzahl, 7 Spalten: nr2, Name, Vorname, Geb.Dat., Stra�e, Hausnummer, PLZ)
    ReDim dataZuordnung(1 To lastRow, 1 To 7)
    
    ' Dictionaries f�r GZ/GZ Neu und Kommune
    Set dictGZ = CreateObject("Scripting.Dictionary")
    Set dictKommune = CreateObject("Scripting.Dictionary")
    
    ' --------------------------------------------------
    ' 1) Spaltenindex suchen (Zeile 1 enth�lt �berschriften)
    ' --------------------------------------------------
    col_nr2 = FindColumnIndex(dataIn, "nr2")
    col_name = FindColumnIndex(dataIn, "Name")
    col_vorname = FindColumnIndex(dataIn, "Vorname")
    col_gebdat = FindColumnIndex(dataIn, "Geb.Dat.")
    col_strasse = FindColumnIndex(dataIn, "Stra�e")
    col_Hausnr = FindColumnIndex(dataIn, "Hausnummer")
    col_plz = FindColumnIndex(dataIn, "PLZ")
    
    ' --------------------------------------------------
    ' 2) �berschriften in dataOut �bernehmen
    '    + �berschriften f�r dataZuordnung
    ' --------------------------------------------------
    For j = 1 To LastCol
        dataOut(1, j) = dataIn(1, j)
    Next j
    
    ' �berschriften in der Zuordnungs-Tabelle
    dataZuordnung(1, 1) = "nr2"
    dataZuordnung(1, 2) = "Name"
    dataZuordnung(1, 3) = "Vorname"
    dataZuordnung(1, 4) = "Geb.Dat."
    dataZuordnung(1, 5) = "Stra�e"
    dataZuordnung(1, 6) = "Hausnummer"
    dataZuordnung(1, 7) = "PLZ"
    
    ' --------------------------------------------------
    ' 3) Hauptschleife: Daten verarbeiten/anonymisieren
    ' --------------------------------------------------
    Dim zeilenID As Long
    zeilenID = 1  ' Startwert f�r nr2
    
    Dim zeileZuord As Long
    zeileZuord = 2 ' Ab Zeile 2 in dataZuordnung (Zeile 1 = �berschrift)
    
    For i = 2 To lastRow
        ' Schleife �ber Spalten
        For j = 1 To LastCol
            Dim spaltenName As String
            Dim origWert As Variant
            
            spaltenName = CStr(dataIn(1, j)) ' �berschrift
            origWert = dataIn(i, j)
            
            Select Case spaltenName
                Case "nr2"
                    ' Zeilen-ID eintragen
                    dataOut(i, j) = zeilenID
                
                Case "GZ", "GZ Neu"
                    ' Dictionary f�r GZ verwenden
                    dataOut(i, j) = GeneriereAnonymisiertenWert(dictGZ, origWert, "GZ_ANON_")
                
                Case "Kommune"
                    ' Dictionary f�r Kommune verwenden
                    dataOut(i, j) = GeneriereAnonymisiertenWert(dictKommune, origWert, "KOM_")
                
                Case "Name", "Vorname", "Geb.Dat.", "Stra�e", "Hausnummer", "PLZ"
                    ' In der Anonymisierten Datei => "NAME_ANON_#", etc.
                    dataOut(i, j) = spaltenName & "_ANON_" & zeilenID
                    
                Case Else
                    ' Unver�ndert
                    dataOut(i, j) = origWert
            End Select
        Next j
        
        ' ------------------------------------------------
        ' Extra: Die Personenwerte in dataZuordnung speichern
        ' ------------------------------------------------
        ' Wichtig: Erst ab Zeile 2 bef�llen
        If col_nr2 > 0 Then dataZuordnung(zeileZuord, 1) = zeilenID
        If col_name > 0 Then dataZuordnung(zeileZuord, 2) = dataIn(i, col_name)
        If col_vorname > 0 Then dataZuordnung(zeileZuord, 3) = dataIn(i, col_vorname)
        If col_gebdat > 0 Then dataZuordnung(zeileZuord, 4) = dataIn(i, col_gebdat)
        If col_strasse > 0 Then dataZuordnung(zeileZuord, 5) = dataIn(i, col_strasse)
        If col_Hausnr > 0 Then dataZuordnung(zeileZuord, 6) = dataIn(i, col_Hausnr)
        If col_plz > 0 Then dataZuordnung(zeileZuord, 7) = dataIn(i, col_plz)
        
        ' Z�hler
        zeilenID = zeilenID + 1
        zeileZuord = zeileZuord + 1
    Next i
    
    ' --------------------------------------------------
    ' 4) dataOut zur�ckschreiben ins Original-Blatt
    '    (so ist in der ge�ffneten Datei "nr2" �berschrieben)
    ' --------------------------------------------------
    wsQuelle.Range(wsQuelle.Cells(1, 1), wsQuelle.Cells(lastRow, LastCol)).Value = dataOut
    
    ' --------------------------------------------------
    ' 5) Neue Workbooks erzeugen + Arrays hineinschreiben
    ' --------------------------------------------------
    
    ' 5a) Anonymisierte Datei
    Dim wbAnon As Workbook, wsAnon As Worksheet
    Set wbAnon = Workbooks.Add
    Set wsAnon = wbAnon.Sheets(1)
    
    wsAnon.Range(wsAnon.Cells(1, 1), wsAnon.Cells(lastRow, LastCol)).Value = dataOut
    
    Dim fnameAnon As String
    fnameAnon = FolderPath & "Anonymisierte_Daten.xlsx"
    wbAnon.SaveAs fnameAnon
    wbAnon.Close False
    
    ' 5b) Zuordnungsdatei (nr2, Name, Vorname, ...)
    Dim wbZuordNr2 As Workbook, wsZuordNr2 As Worksheet
    Dim rowsZuord As Long
    
    ' Max Zeile in dataZuordnung ist zeileZuord - 1
    ' (weil wir zeileZuord am Ende immer um 1 erh�ht haben)
    rowsZuord = zeileZuord - 1
    
    Set wbZuordNr2 = Workbooks.Add
    Set wsZuordNr2 = wbZuordNr2.Sheets(1)
    
    wsZuordNr2.Range("A1:G" & rowsZuord).Value = dataZuordnung
    
    Dim fnameZuordNr2 As String
    fnameZuordNr2 = FolderPath & "Zuordnung_Nr2.xlsx"
    wbZuordNr2.SaveAs fnameZuordNr2
    wbZuordNr2.Close False
    
    ' 5c) Zuordnung GZ
    ExportiereZuordnung dictGZ, FolderPath & "Zuordnung_GZ.xlsx", "Original GZ", "Anonymisiert GZ"
    
    ' 5d) Zuordnung Kommune
    ExportiereZuordnung dictKommune, FolderPath & "Zuordnung_Kommune.xlsx", "Original Kommune", "Anonymisiert Kommune"
    
    ' Fertig
End Sub

' ------------------------------------------------------------------------------
' Liefert den Spaltenindex f�r den �bergebenen Spaltennamen (1. Zeile in dataIn)
' oder -1, falls nicht gefunden.
' dataIn(Zeile, Spalte) -> dataIn(1, j) sind die �berschriften
' ------------------------------------------------------------------------------
Private Function FindColumnIndex(ByVal dataIn As Variant, ByVal colName As String) As Long
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
' Erzeugt einen anonymisierten Wert auf Basis eines Dictionaries.
' Existiert "OriginalWert" schon, wird der bereits vergebene anonymisierte
' zur�ckgegeben. Andernfalls wird ein neuer generiert und gespeichert.
' ------------------------------------------------------------------------------
Private Function GeneriereAnonymisiertenWert(ByVal dict As Object, _
                                             ByVal OriginalWert As Variant, _
                                             ByVal Prefix As String) As String
    Dim key As String
    key = Trim(UCase(CStr(OriginalWert)))
    
    If key = "" Then
        ' Leer
        GeneriereAnonymisiertenWert = ""
    Else
        If Not dict.Exists(key) Then
            dict.Add key, Prefix & (dict.Count + 1)
        End If
        GeneriereAnonymisiertenWert = dict(key)
    End If
End Function

' ------------------------------------------------------------------------------
' Exportiert den Inhalt eines Dictionary in eine neue Arbeitsmappe.
' Spalte1 = Original, Spalte2 = Anonymisiert
' ------------------------------------------------------------------------------
Private Sub ExportiereZuordnung(ByVal dict As Object, _
                                ByVal FileName As String, _
                                ByVal Spalte1 As String, _
                                ByVal Spalte2 As String)
    ' Wenn Dictionary leer ist, nichts tun
    If dict.Count = 0 Then Exit Sub
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    ws.Range("A1").Value = Spalte1
    ws.Range("B1").Value = Spalte2
    
    Dim r As Long
    r = 2
    
    Dim k As Variant
    For Each k In dict.Keys
        ws.Cells(r, 1).Value = k
        ws.Cells(r, 2).Value = dict(k)
        r = r + 1
    Next k
    
    wb.SaveAs FileName
    wb.Close False
End Sub


