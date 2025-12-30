# Easy Folder Permissions - Manager & Reader V0.0.1

## 🇩🇪 DEUTSCH

### Übersicht
Zwei PowerShell-Tools mit WPF GUI zur Verwaltung und Analyse von Windows-Ordnerberechtigungen im modernen Windows 11 Design.

### 📦 Tools

#### **easyFPManager** - Berechtigungsverwaltung
- ✅ Berechtigungen anzeigen, hinzufügen und entfernen
- ✅ Flexible Anwendung (Hauptordner, Unterordner, beide)
- ✅ Backup-System und automatische Reports
- ✅ CSV/HTML-Export, Live-Filter
- ✅ Erweiterte Fehlerbehandlung

#### **easyFPReader** - Berechtigungsanalyse
- ✅ Rekursive Berechtigungsanalyse mit Baumstruktur
- ✅ HTML-Export mit interaktiver Baumansicht
- ✅ OnPrem zu EntraID/SharePoint UPN-Mapping
- ✅ Manuelle Benutzer-Zuordnungen
- ✅ CSV-Export für SharePoint-Integration

### 🚀 Systemanforderungen
- Windows 10/11 oder Windows Server 2016+
- PowerShell 5.1 oder höher
- Administratorrechte erforderlich
- .NET Framework 4.5+

### 📖 Schnellstart

**easyFPManager starten:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\easyFPManager_V0.0.1.ps1
```

**easyFPReader starten:**
```powershell
.\easyFPReader_V0.0.1.ps1
```

### ⚙️ Features easyFPManager
| Feature | Beschreibung |
|---------|-------------|
| 📁 Ordner-Browser | Auswahl über Dialog oder Pfadeingabe |
| 🔍 Filter | Benutzer, Rechte, Vererbung in Echtzeit |
| 👤 Benutzer-Verwaltung | Hinzufügen/Entfernen mit Benutzer-Browser |
| 📊 Export | CSV, HTML und XML-Backups |
| 💾 Reports | Automatische Reports bei jeder Änderung |

### ⚙️ Features easyFPReader
| Feature | Beschreibung |
|---------|-------------|
| 📁 Baumstruktur | Hierarchische Darstellung aller Ordner |
| 📄 HTML-Export | Interaktive Reports mit Expand/Collapse |
| 🔄 UPN-Mapping | Automatische OnPrem → EntraID Konvertierung |
| 👥 Benutzer-Mapping | Manuelle Anpassungen für SharePoint |
| 📊 CSV-Export | Für SharePoint Online Integration |

### ⚠️ Sicherheitshinweise
- Script erfordert Administratorrechte
- Änderungen sind sofort wirksam
- Reports werden automatisch erstellt
- Backup vor größeren Änderungen empfohlen

### 📁 Report-Speicherort
```
%USERPROFILE%\Documents\FolderPermissions_Reports\
```

---

## 🇬🇧 ENGLISH

### Overview
Two PowerShell tools with WPF GUI for managing and analyzing Windows folder permissions in modern Windows 11 design.

### 📦 Tools

#### **easyFPManager** - Permission Management
- ✅ View, add, and remove permissions
- ✅ Flexible application (main folder, subfolders, both)
- ✅ Backup system and automatic reports
- ✅ CSV/HTML export, real-time filters
- ✅ Advanced error handling

#### **easyFPReader** - Permission Analysis
- ✅ Recursive permission analysis with tree structure
- ✅ HTML export with interactive tree view
- ✅ OnPrem to EntraID/SharePoint UPN mapping
- ✅ Manual user mappings
- ✅ CSV export for SharePoint integration

### 🚀 System Requirements
- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Administrator rights required
- .NET Framework 4.5+

### 📖 Quick Start

**Launch easyFPManager:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\easyFPManager_V0.0.1.ps1
```

**Launch easyFPReader:**
```powershell
.\easyFPReader_V0.0.1.ps1
```

### ⚙️ easyFPManager Features
| Feature | Description |
|---------|-------------|
| 📁 Folder Browser | Selection via dialog or path input |
| 🔍 Filter | User, rights, inheritance in real-time |
| 👤 User Management | Add/remove with user browser |
| 📊 Export | CSV, HTML and XML backups |
| 💾 Reports | Automatic reports on every change |

### ⚙️ easyFPReader Features
| Feature | Description |
|---------|-------------|
| 📁 Tree Structure | Hierarchical view of all folders |
| 📄 HTML Export | Interactive reports with expand/collapse |
| 🔄 UPN Mapping | Automatic OnPrem → EntraID conversion |
| 👥 User Mapping | Manual adjustments for SharePoint |
| 📊 CSV Export | For SharePoint Online integration |

### ⚠️ Security Notes
- Scripts require administrator rights
- Changes take effect immediately
- Reports are automatically generated
- Backup recommended before major changes

### 📁 Report Location
```
%USERPROFILE%\Documents\FolderPermissions_Reports\
```

---

**Developed by PhinIT.DE © 2025**
