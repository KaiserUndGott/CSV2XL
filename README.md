# CSV2XL

VBA Makro für Microsoft Excel (Mac & Windows) zum automatischen Import von CSV-Dateien mit Power Query und professioneller Formatierung.

## 🎯 Features

- ✅ **Plattformübergreifend**: Funktioniert unter Excel für Mac OSX und Windows
- ✅ **Power Query Integration**: Nutzt Power Query für zuverlässigen CSV-Import
- ✅ **UTF-8 Unterstützung**: Korrekte Darstellung von Umlauten und Sonderzeichen
- ✅ **Automatische Formatierung**: Tabellenstil "Dunkelblaugrün, Mittel 2"
- ✅ **Fixierte Kopfzeile**: Erste Zeile bleibt beim Scrollen sichtbar
- ✅ **Leere Spalten entfernen**: Automatisches Bereinigen leerer Spalten
- ✅ **Optimierte Spaltenbreite**: Automatische Anpassung der Spaltenbreiten

## 📋 Voraussetzungen

- Microsoft Excel 2016 oder neuer (Mac oder Windows)
- Power Query aktiviert (in den meisten Excel-Versionen standardmäßig enthalten)

## 🚀 Installation

### Schritt 1: Makro importieren

1. Öffnen Sie Microsoft Excel
2. Drücken Sie `Alt + F11` (Windows) oder `Option + F11` (Mac) um den VBA-Editor zu öffnen
3. Klicken Sie auf `Datei` > `Datei importieren...`
4. Wählen Sie die Datei `CSV2XL.bas` aus
5. Das Modul "CSV2XL" wird nun in Ihrer Arbeitsmappe angezeigt

### Schritt 2: Makro verwenden

#### Methode 1: Power Query Import (empfohlen)

1. Drücken Sie `Alt + F8` (Windows) oder `Option + F8` (Mac)
2. Wählen Sie `ImportCSVWithPowerQuery` aus der Liste
3. Klicken Sie auf `Ausführen`
4. Wählen Sie Ihre CSV-Datei aus

#### Methode 2: Direkter Import (Fallback)

Falls Power Query nicht verfügbar ist:

1. Drücken Sie `Alt + F8` (Windows) oder `Option + F8` (Mac)
2. Wählen Sie `ImportCSVDirect` aus der Liste
3. Klicken Sie auf `Ausführen`
4. Wählen Sie Ihre CSV-Datei aus

### Schritt 3: Schnellzugriff einrichten (optional)

#### Für häufige Nutzung - Schaltfläche im Menüband:

1. Klicken Sie mit der rechten Maustaste auf das Menüband
2. Wählen Sie `Menüband anpassen`
3. Erstellen Sie eine neue Registerkarte oder Gruppe
4. Fügen Sie das Makro `ImportCSVWithPowerQuery` hinzu

#### Tastenkombination zuweisen:

1. Öffnen Sie den VBA-Editor (`Alt/Option + F11`)
2. Doppelklicken Sie auf das Modul "CSV2XL"
3. Ändern Sie die erste Zeile der Sub zu:
   ```vba
   Sub ImportCSVWithPowerQuery()
   ' Drücken Sie jetzt: Extras > Makros > Makros
   ' Wählen Sie das Makro und klicken auf "Optionen"
   ' Weisen Sie eine Tastenkombination zu (z.B. Strg+Shift+I)
   ```

## 📝 CSV-Datei Anforderungen

Die CSV-Datei sollte folgendes Format haben:

- **Encoding**: UTF-8
- **Trennzeichen**: Komma (`,`)
- **Erste Zeile**: Spaltenüberschriften

Beispiel:
```csv
Name,Alter,Stadt,E-Mail
Max Mustermann,35,Berlin,max@example.com
Anna Schmidt,28,München,anna@example.com
```

## 🎨 Anpassungen

### Tabellenstil ändern

Öffnen Sie die `.bas` Datei und ändern Sie die Zeile:

```vba
tbl.TableStyle = "TableStyleMedium10"
```

Verfügbare Stile:
- `TableStyleMedium1` bis `TableStyleMedium28`
- `TableStyleLight1` bis `TableStyleLight21`
- `TableStyleDark1` bis `TableStyleDark11`

### Trennzeichen ändern

Für andere Trennzeichen (z.B. Semikolon) ändern Sie im M-Code:

```vba
mCode = "let" & vbCrLf & _
        "    Source = Csv.Document(File.Contents(""" & csvFilePath & """)," & _
        "[Delimiter="";"", Columns=null, Encoding=65001, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
```

## 🐛 Fehlerbehebung

### "Power Query nicht gefunden"
- Verwenden Sie `ImportCSVDirect` statt `ImportCSVWithPowerQuery`
- Stellen Sie sicher, dass Power Query in Excel aktiviert ist

### "Encoding-Probleme / Umlaute falsch dargestellt"
- Stellen Sie sicher, dass die CSV-Datei in UTF-8 kodiert ist
- In Windows: Speichern Sie die CSV mit "UTF-8" Encoding

### "Makro kann nicht ausgeführt werden"
- Überprüfen Sie die Makro-Sicherheitseinstellungen in Excel
- Mac: `Excel` > `Einstellungen` > `Sicherheit & Datenschutz`
- Windows: `Datei` > `Optionen` > `Trust Center` > `Einstellungen für das Trust Center`

### "Leere Spalten werden nicht entfernt"
- Stellen Sie sicher, dass die Spalten wirklich komplett leer sind
- Das Makro entfernt nur Spalten ohne jeglichen Inhalt

## 🔧 Technische Details

### Funktionsweise

1. **Dateiauswahl**: Öffnet einen Dialog zur Auswahl der CSV-Datei
2. **Power Query**: Erstellt eine M-Code Query für den Import
3. **Encoding**: UTF-8 (Code 65001) für korrekte Sonderzeichen
4. **Tabellenkonvertierung**: Wandelt Daten in formatierte Excel-Tabelle
5. **Bereinigung**: Entfernt leere Spalten
6. **Formatierung**: Wendet Tabellenstil und fixiert Kopfzeile

### Kompatibilität

- ✅ Excel 2016+ für Mac
- ✅ Excel 2016+ für Windows
- ✅ Excel 2019, 2021
- ✅ Microsoft 365 Excel

## 📄 Lizenz

MIT License - siehe [LICENSE](LICENSE) Datei

## 🤝 Beitragen

Contributions sind willkommen! Bitte öffnen Sie ein Issue oder Pull Request.

## 📧 Support

Bei Fragen oder Problemen öffnen Sie bitte ein [Issue](https://github.com/KaiserUndGott/CSV2XL/issues) auf GitHub.

---

**Hinweis**: Dieses Makro wurde für den produktiven Einsatz entwickelt und getestet. Bei Problemen oder Verbesserungsvorschlägen freuen wir uns über Ihr Feedback!
