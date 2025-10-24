# Easy Folder Permissions Manager (easyFPFM) V0.0.1

## Übersicht
Ein PowerShell-Script mit integrierter WPF GUI zur einfachen Verwaltung von Windows-Ordnerberechtigungen im modernen Windows 11 Design.

## 🎨 **Neues Windows 11 Design**
- **Moderne Oberfläche**: Zeitloses Windows 11 Design mit Fluent Design Elementen
- **Intuitive Icons**: Emoji-basierte Icons für bessere Benutzerfreundlichkeit
- **Responsive Layout**: Optimiert für verschiedene Bildschirmgrößen
- **Hover-Effekte**: Moderne Animationen und Übergänge

## 🚀 **Erweiterte Features**
- ✅ **Berechtigungen anzeigen**: Aktuelle Ordnerberechtigungen übersichtlich anzeigen
- ✅ **Benutzer hinzufügen**: Neue Benutzer mit spezifischen Rechten hinzufügen
- ✅ **Berechtigungen entfernen**: Sichere Entfernung von Benutzerberechtigungen
- ✅ **Flexible Anwendung**: Hauptordner, nur Unterordner oder beide
- ✅ **Erweiterte Filter**: Benutzer-, Rechte- und Vererbungsfilter
- ✅ **Backup-System**: Automatische Sicherung vor Änderungen
- ✅ **CSV-Export**: Exportieren der Berechtigungen als CSV-Datei
- ✅ **Automatische Reports**: Detaillierte Berichte bei jeder Änderung
- ✅ **Live-Filter**: Echtzeitfilterung während der Eingabe
- ✅ **Fehlerbehandlung**: Umfassende Validierung und Fehlermeldungen

## Systemanforderungen
- Windows 10/11 oder Windows Server 2016+
- PowerShell 5.1 oder höher
- Administratorrechte erforderlich
- .NET Framework 4.5+ (für WPF)

## Installation & Start
1. **Script herunterladen**: `easyFPFM_V0.0.1.ps1`
2. **Als Administrator ausführen**:
   ```powershell
   # PowerShell als Administrator öffnen
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   .\easyFPFM_V0.0.1.ps1
   ```

## 📖 **Bedienung**

### 1. 📁 Ordner auswählen
- **📂 Durchsuchen**: Ordner über Dialog auswählen
- **Pfad eingeben**: Direkteingabe des Ordnerpfads möglich
- **🔄 Laden**: Aktuelle Berechtigungen anzeigen
- **💾 Backup**: Sicherung der aktuellen Berechtigungen erstellen

### 2. 🔍 Filter und Optionen
- **Unterordner einbeziehen**: Checkbox für rekursive Anzeige
- **Nur vererbte Berechtigungen**: Filter für vererbte Rechte
- **Benutzer Filter**: Live-Suche nach Benutzernamen
- **Rechte Filter**: Live-Suche nach Berechtigungstypen
- **🔍 Filter Button**: Manuelle Filteranwendung

### 3. 📋 Berechtigungen anzeigen
- **DataGrid**: Übersichtliche Tabelle mit allen Berechtigungen
  - Ordner, Benutzer, Rechte, Zugriffstyp, Vererbt-Status
- **Live-Zähler**: Anzeige der gefilterten/gesamten Einträge
- **🗑️ Ausgewählte entfernen**: Mehrfachauswahl zum Löschen

### 4. 👤 Benutzer verwalten
- **Benutzername**: Windows-Benutzername oder Gruppe eingeben
- **👥 Benutzer-Browser**: Auswahl aus verfügbaren Benutzern (geplant)
- **Rechte auswählen**:
  - `🔓 Vollzugriff (FullControl)`: Alle Rechte
  - `✏️ Ändern (Modify)`: Lesen, Schreiben, Ausführen, Löschen
  - `📖 Lesen und Ausführen (ReadAndExecute)`: Lesen und Ausführen
  - `👁️ Nur Lesen (Read)`: Nur Lesen
  - `✍️ Nur Schreiben (Write)`: Nur Schreiben
- **Anwendungsbereich**:
  - `📁 Nur Hauptordner`: Berechtigung nur für den ausgewählten Ordner
  - `📂 Nur Unterordner`: Berechtigung nur für alle Unterordner
  - `🗂️ Hauptordner + Unterordner`: Berechtigung für alles
- **➕ Berechtigung hinzufügen**: Neue Berechtigung erstellen
- **➖ Berechtigung entfernen**: Bestehende Berechtigung löschen

### 5. 📊 Reports und Export
- **📁 Reports öffnen**: Öffnet den Reports-Ordner
- **📄 CSV Export**: Exportiert alle Berechtigungen als CSV-Datei
- **Automatische Reports**: Bei jeder Berechtigungsänderung
- **Speicherort**: `%USERPROFILE%\Documents\FolderPermissions_Reports\`
- **Backup-Format**: XML-Dateien mit Zeitstempel
- **Report-Format**: Textdatei mit detaillierter Dokumentation

### 6. ⚙️ Weitere Funktionen
- **🔄 Aktualisieren**: Berechtigungen neu laden
- **⚙️ Einstellungen**: Konfigurationsoptionen (geplant)
- **❌ Beenden**: Anwendung schließen

## Benutzerbeispiele

### Beispiel 1: Einzelnen Benutzer hinzufügen
```
Ordner: C:\Projekte\WebApp
Benutzer: DOMAIN\john.doe
Rechte: Modify
Anwenden auf: Hauptordner + Unterordner
```

### Beispiel 2: Gruppe nur für Unterordner
```
Ordner: C:\Daten
Benutzer: Entwickler
Rechte: ReadAndExecute
Anwenden auf: Nur Unterordner
```

## Sicherheitshinweise
⚠️ **Wichtige Hinweise**:
- Script erfordert Administratorrechte
- Berechtigungsänderungen sind sofort wirksam
- Backup der aktuellen Berechtigungen wird empfohlen
- Reports werden automatisch erstellt zur Nachverfolgung

## Fehlerbehebung

### Häufige Probleme
1. **"Execution Policy" Fehler**:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

2. **"Zugriff verweigert"**:
   - PowerShell als Administrator starten
   - Benutzerrechte auf Ordner prüfen

3. **WPF lädt nicht**:
   - .NET Framework 4.5+ installieren
   - Windows Updates prüfen

### Debug-Modus
Für erweiterte Fehlerdiagnose:
```powershell
$DebugPreference = "Continue"
.\easyFPFM_V0.0.1.ps1
```

## Report-Format
```
=== FOLDER PERMISSIONS REPORT ===
Datum/Zeit: 24.10.2025 15:30:45
Aktion: ADD
Ordner: C:\TestFolder
Benutzer: TestUser
Rechte: Modify
Angewendet auf: Both

=== AKTUELLE BERECHTIGUNGEN NACH ÄNDERUNG ===
C:\TestFolder | BUILTIN\Administrators | FullControl | Allow
C:\TestFolder | TestUser | Modify | Allow
...
```

## Technische Details
- **PowerShell Version**: 5.1+
- **GUI Framework**: WPF (Windows Presentation Foundation)
- **Berechtigungs-API**: .NET System.Security.AccessControl
- **Fehlerbehandlung**: Try-Catch mit MessageBox-Ausgabe
- **Threading**: UI-Thread für alle Operationen

## Changelog
### V0.0.1 (24.10.2025)
- Erste Version mit vollständiger GUI
- Grundfunktionen für Berechtigungsverwaltung
- Automatisches Report-System
- Unterstützung für Haupt-/Unterordner-Optionen

## Support
Bei Problemen oder Fragen:
- GitHub Issues erstellen
- PowerShell-Logs prüfen
- Reports-Ordner für Debugging nutzen

## Lizenz
Entwickelt von PhinIT Development
Für interne Nutzung und Weiterentwicklung freigegeben.
