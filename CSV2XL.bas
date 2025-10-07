Attribute VB_Name = "CSV2XL"
' CSV2XL - CSV Import Makro für Excel (Mac & Windows)
' Importiert CSV-Dateien mit Power Query und formatiert sie automatisch
' Kompatibel mit Excel für Mac OSX und Windows

Option Explicit

Sub ImportCSVWithPowerQuery()
    Dim csvFilePath As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim pq As WorkbookQuery
    Dim queryName As String
    Dim tableName As String
    Dim lastCol As Long
    Dim col As Long
    Dim emptyColumns As String
    Dim i As Long

    On Error GoTo ErrorHandler

    ' CSV-Datei auswählen
    csvFilePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="CSV-Datei auswählen")

    ' Abbruch wenn keine Datei gewählt wurde
    If csvFilePath = "False" Then Exit Sub

    ' Eindeutigen Namen für Query und Tabelle generieren
    queryName = "CSV_Import_" & Format(Now, "yyyymmddhhmmss")
    tableName = "Table_" & Format(Now, "yyyymmddhhmmss")

    ' Neues Arbeitsblatt erstellen
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "CSV Import " & Format(Now, "hh-mm-ss")

    ' Power Query erstellen (M-Code)
    Dim mCode As String
    mCode = "let" & vbCrLf & _
            "    Source = Csv.Document(File.Contents(""" & csvFilePath & """)," & _
            "[Delimiter="","", Columns=null, Encoding=65001, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
            "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & vbCrLf & _
            "in" & vbCrLf & _
            "    PromotedHeaders"

    ' Power Query hinzufügen und ausführen
    #If Mac Then
        ' Mac-spezifische Implementation
        Set pq = ThisWorkbook.Queries.Add(queryName, mCode)
    #Else
        ' Windows-spezifische Implementation
        Set pq = ThisWorkbook.Queries.Add(queryName, mCode)
    #End If

    ' Query mit Arbeitsblatt verbinden
    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName, _
        Destination:=ws.Range("A1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With

    ' Tabelle formatieren
    Set tbl = ws.ListObjects(1)
    tbl.Name = tableName

    ' Tabellenstil anwenden: "Dunkelblaugrün, Mittel 2"
    ' TableStyleMedium2 entspricht "Dunkelblaugrün, Mittel 2"
    tbl.TableStyle = "TableStyleMedium16"

    ' Leere Spalten identifizieren und löschen
    lastCol = tbl.Range.Columns.Count

    ' Von rechts nach links durchgehen um Indexprobleme zu vermeiden
    For col = lastCol To 1 Step -1
        If WorksheetFunction.CountA(tbl.ListColumns(col).DataBodyRange) = 0 Then
            tbl.ListColumns(col).Delete
        End If
    Next col

    ' Kopfzeile fixieren
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Spaltenbreite optimieren
    tbl.Range.Columns.AutoFit

    ' Zur ersten Zelle gehen
    ws.Range("A1").Select

    MsgBox "CSV-Datei erfolgreich importiert!" & vbCrLf & _
           "Arbeitsblatt: " & ws.Name & vbCrLf & _
           "Tabelle: " & tbl.Name, vbInformation, "CSV2XL"

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Import der CSV-Datei:" & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "CSV2XL Fehler"
End Sub

' Alternative Methode ohne Power Query für bessere Kompatibilität
Sub ImportCSVDirect()
    Dim csvFilePath As String
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim tbl As ListObject
    Dim lastCol As Long
    Dim col As Long

    On Error GoTo ErrorHandler

    ' CSV-Datei auswählen
    csvFilePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="CSV-Datei auswählen")

    If csvFilePath = "False" Then Exit Sub

    ' Neues Arbeitsblatt erstellen
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "CSV Import " & Format(Now, "hh-mm-ss")

    ' CSV importieren
    Set qt = ws.QueryTables.Add( _
        Connection:="TEXT;" & csvFilePath, _
        Destination:=ws.Range("A1"))

    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(1)
        #If Mac Then
            .TextFilePlatform = 65001 ' UTF-8 für Mac
        #Else
            .TextFilePlatform = 65001 ' UTF-8 für Windows
        #End If
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    ' Als Tabelle formatieren
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes)
    tbl.TableStyle = "TableStyleMedium16"

    ' Leere Spalten löschen
    lastCol = tbl.Range.Columns.Count
    For col = lastCol To 1 Step -1
        If WorksheetFunction.CountA(tbl.ListColumns(col).DataBodyRange) = 0 Then
            tbl.ListColumns(col).Delete
        End If
    Next col

    ' Kopfzeile fixieren
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Spaltenbreite optimieren
    tbl.Range.Columns.AutoFit

    ws.Range("A1").Select

    MsgBox "CSV-Datei erfolgreich importiert!", vbInformation, "CSV2XL"

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Import: " & Err.Description, vbCritical, "CSV2XL Fehler"
End Sub
