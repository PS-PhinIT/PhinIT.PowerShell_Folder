# Easy Folder Permissions - Manager & Reader V0.0.2

PowerShell-Tools mit WPF-GUI zur Verwaltung und Analyse von Windows-Ordnerberechtigungen (NTFS)
sowie – als Beta – zur Übertragung der Berechtigungen nach SharePoint Online.

📄 [CHANGELOG](CHANGELOG.md) · 🔒 [SECURITY](SECURITY.md) · 🤝 [CONTRIBUTING](CONTRIBUTING.md) · 🔍 [Audit-Bericht](docs/AUDIT.md)

## 📁 Ordnerstruktur / Folder structure

```
PhinIT.PowerShell_Folder/
├── src/
│   ├── easyFPManager/
│   │   └── easyFPManager_V0.0.2.ps1          # Berechtigungen verwalten (Admin)
│   ├── easyFPReader/
│   │   └── easyFPReader_V0.0.2.ps1           # Berechtigungen analysieren + SharePoint-Mapping
│   └── beta/
│       └── easyFolder-SPOPermissions/
│           └── easyFolder-SPOPermissions_V0.0.2.ps1   # BETA: Mapping-CSV → SharePoint Online
├── docs/
│   └── AUDIT.md                              # Audit-Bericht V0.0.2
├── CHANGELOG.md
├── CONTRIBUTING.md
├── SECURITY.md
└── README.md
```

> Vorkompilierte EXE-Dateien liegen nicht mehr im Repository, sondern werden als
> Assets unter **GitHub Releases** veröffentlicht. / Prebuilt EXE files are published as
> **GitHub Release** assets instead of being stored in the repository.

---

## 🇩🇪 DEUTSCH

### 📦 Tools

#### **easyFPManager** – Berechtigungsverwaltung
- ✅ Berechtigungen anzeigen, hinzufügen und gezielt entfernen (nur die ausgewählte Regel)
- ✅ Saubere NTFS-Vererbung: *Nur Hauptordner*, *Nur Unterordner*, *Hauptordner + Unterordner*
- ✅ Wiederherstellbares Backup (`icacls /save`) plus lesbares XML-Backup
- ✅ Automatische Reports, CSV-/HTML-Export, Live-Filter
- ✅ Benutzer-/Gruppenauswahl per eigenem Dialog (lokale Konten, Gruppen, Systemkonten)

#### **easyFPReader** – Berechtigungsanalyse
- ✅ Rekursive Analyse mehrerer Stammordner mit Baumansicht
- ✅ HTML-Export (sicher kodiert) mit Einrückung nach Ordnerebene
- ✅ OnPrem → Entra ID/SharePoint UPN-Mapping (Tab 2) und manuelle Zuordnungen (Tab 3)
- ✅ CSV-Export als Grundlage für das Beta-Tool

#### **easyFolder-SPOPermissions** – SharePoint Online (BETA)
- 🧪 Import der Reader-CSV, Benutzervalidierung, Anlage der Ordner, Setzen der Berechtigungen
- 🧪 Erkennt Bibliothek und lokalisierte Berechtigungsstufen (Vollzugriff/Bearbeiten/Mitwirken/Lesen) automatisch
- ⚠️ Beta – zuerst in einer Test-Site ausprobieren!

### 🚀 Systemanforderungen
| | Manager | Reader | Beta (SPO) |
|---|---|---|---|
| Betriebssystem | Windows 10/11, Server 2016+ | Windows 10/11, Server 2016+ | Windows 10/11 |
| PowerShell | 5.1 (empfohlen) oder 7.x | 5.1 oder 7.x | 5.1 + PnP.PowerShell 1.12 **oder** 7.4+ + PnP.PowerShell 2.x/3.x |
| Rechte | **Administrator** | Leserechte auf die Ordner | SharePoint-Admin + eigene Entra-ID-App-Registrierung |

### 📖 Schnellstart
```powershell
# einmalig (heruntergeladene Dateien entsperren)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Get-ChildItem .\src -Recurse -Filter *.ps1 | Unblock-File

# Manager (als Administrator)
.\src\easyFPManager\easyFPManager_V0.0.2.ps1

# Reader
.\src\easyFPReader\easyFPReader_V0.0.2.ps1

# Beta: SharePoint Online
Install-Module PnP.PowerShell -Scope CurrentUser          # PS 7.4+
# Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser   # Windows PowerShell 5.1
.\src\beta\easyFolder-SPOPermissions\easyFolder-SPOPermissions_V0.0.2.ps1
```

### ⚙️ Optionen „Anwenden auf“ (Manager)
| Option | Wirkung |
|---|---|
| Nur Hauptordner (inkl. Dateien) | Regel gilt für den Ordner und die direkt enthaltenen Dateien, nicht für Unterordner |
| Nur Unterordner (inkl. Inhalt) | Vererbbare Regel auf jedem direkten Unterordner, nicht auf dem Hauptordner |
| Hauptordner + Unterordner + Dateien | Eine vererbbare Regel auf dem Hauptordner |

Unterordner mit **deaktivierter Vererbung** erhalten bei den beiden letzten Optionen zusätzlich eine explizite Regel.

### 💾 Reports & Backups
```
%USERPROFILE%\Documents\FolderPermissions_Reports\
```
Wiederherstellen eines Backups (als Administrator, im **übergeordneten** Ordner des gesicherten Ordners):
```powershell
icacls "D:\Daten" /restore "$env:USERPROFILE\Documents\FolderPermissions_Reports\Backup_<Ordner>_<Zeitstempel>.acl"
```

### ⚠️ Sicherheitshinweise
- Änderungen des Managers sind sofort wirksam – vorher Backup erstellen.
- Skripte sind ab V0.0.2 **nicht** im Repository signiert; signierte Fassungen gibt es als Release-Assets (siehe [CONTRIBUTING](CONTRIBUTING.md)).
- Sicherheitslücken bitte gemäß [SECURITY.md](SECURITY.md) melden.

---

## 🇬🇧 ENGLISH

### 📦 Tools
- **easyFPManager** – view, add and precisely remove NTFS permissions. It uses proper inheritance, restorable `icacls` backups, reports and CSV/HTML export.
- **easyFPReader** – recursive permission analysis with tree view and a safely encoded HTML export. It also creates the OnPrem → Entra ID/SharePoint UPN mapping and exports it to CSV.
- **easyFolder-SPOPermissions (BETA)** – applies the reader's mapping CSV to SharePoint Online (PnP.PowerShell). Try it on a test site first.

### 🚀 Requirements
- Windows 10/11 or Windows Server 2016+.
- PowerShell 5.1 or 7.x.
- The Manager requires **administrator** rights.
- Beta tool:
  - PnP.PowerShell 1.12 on Windows PowerShell 5.1, or PnP.PowerShell 2.x/3.x on PowerShell 7.4+.
  - Your own Entra ID app registration.

### 📖 Quick start
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Get-ChildItem .\src -Recurse -Filter *.ps1 | Unblock-File
.\src\easyFPManager\easyFPManager_V0.0.2.ps1      # run as administrator
.\src\easyFPReader\easyFPReader_V0.0.2.ps1
.\src\beta\easyFolder-SPOPermissions\easyFolder-SPOPermissions_V0.0.2.ps1
```

### 💾 Reports & backups
Reports and backups are stored in `%USERPROFILE%\Documents\FolderPermissions_Reports\`.

To restore a backup, run `icacls "<parent folder>" /restore "<backup>.acl"` as administrator.

### ⚠️ Security notes
- Manager changes take effect immediately, so create a backup first.
- Starting with V0.0.2, scripts in the repository are unsigned. Signed builds are published as release assets.
- Report vulnerabilities as described in [SECURITY.md](SECURITY.md).

---

**Developed by PhinIT.DE © 2025**
