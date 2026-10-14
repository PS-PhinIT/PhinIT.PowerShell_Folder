# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung an [Semantic Versioning](https://semver.org/lang/de/).

> **Regel:** Jeder Eintrag gehört zu einer konkreten Version mit Datum. Es gibt keine
> Sammelabschnitte ohne Version – Änderungen werden mit der Version eingetragen, in der sie erscheinen.
>
> **Herkunft der Einträge bis 0.0.1:** Für das Repository existieren keine Pull Requests.
> Die Einträge bis einschließlich 0.0.1 wurden daher aus der Git-Commit-Historie rekonstruiert
> (Commit-Hashes in Klammern).

---

## [0.0.2] - 2026-10-14

Umfassendes Audit (Sicherheit, Performance, Funktionen, Verweise, GUI, Befehle) mit
Behebung der Befunde. Der vollständige Audit-Bericht steht in [`docs/AUDIT.md`](docs/AUDIT.md).

### Hinzugefügt
- Neue Ordnerstruktur: `src/<Tool>/`, `src/beta/<Tool>/`, `docs/`.
- `CHANGELOG.md`, `SECURITY.md`, `CONTRIBUTING.md`, `docs/AUDIT.md`, `.gitignore`, `.gitattributes`.
- **easyFPManager**
  - Wiederherstellbares ACL-Backup per `icacls /save` (`.acl`-Datei) zusätzlich zum lesbaren XML-Backup.
  - Eigener WPF-Dialog zur Benutzer-/Gruppenauswahl (ersetzt `Out-GridView`, das in EXE-Builds und unter PowerShell 7 nicht zuverlässig verfügbar ist).
  - Bestätigungsdialog vor dem Hinzufügen von Berechtigungen.
  - Spalten „Vererbung an“ (InheritanceFlags/PropagationFlags) in Grid, CSV- und HTML-Export.
  - Tooltips in der Oberfläche; Reports protokollieren den ausführenden Benutzer.
- **easyFPReader**
  - Validierung von Tenant-Domain, SharePoint-URL (HTTPS) und UPN-Format bei manuellen Zuordnungen.
  - Unterstützung mehrerer Stammordner bei der URL-Umrechnung (längster passender Stammordner gewinnt).
- **easyFolder-SPOPermissions (Beta)**
  - Umrechnung der vom Reader exportierten vollständigen URLs in site-relative Pfade (mit Prüfung auf Tenant und Site).
  - Automatische Ermittlung der Dokumentbibliothek aus dem Pfad und der lokalisierten Berechtigungsstufen (z. B. „Vollzugriff“, „Bearbeiten“).
  - URL-Prüfung vor dem Verbindungsaufbau.

### Geändert
- Alle Skripte auf Version **V0.0.2** angehoben und als **UTF-8 mit BOM** gespeichert (korrekte Umlaute/Emojis unter Windows PowerShell 5.1).
- **easyFPManager**
  - Hinzufügen nutzt jetzt NTFS-Vererbung statt einer expliziten Regel auf jedem Unterordner:
    - *Nur Hauptordner*: Ordner + enthaltene Dateien.
    - *Nur Unterordner*: vererbbare Regel auf den direkten Unterordnern.
    - *Beide*: eine vererbbare Regel auf dem Hauptordner.
    - Unterordner mit deaktivierter Vererbung erhalten zusätzlich eine explizite Regel.
  - Benutzer/Gruppen werden über die SID aufgelöst (eindeutig, sprachunabhängig).
  - `Create-PermissionReport` in `New-PermissionReport` umbenannt (zulässiges PowerShell-Verb).
  - Reports liegen unter `Dokumente\FolderPermissions_Reports` (auch bei umgeleitetem Dokumente-Ordner).
  - CSV-Export mit `;` als Trennzeichen (Excel, deutsch).
- **easyFPReader**
  - Tabs getauscht: Tab 2 = SharePoint Permissions-Mapping, Tab 3 = Benutzer-Mapping (passend zu den Hinweistexten).
  - HTML-Export behält die Analysereihenfolge bei und rückt relativ ein.
- **easyFolder-SPOPermissions (Beta)**
  - Ordner werden mit `Resolve-PnPFolder` angelegt.
  - Die Option „Vererbung unterbrechen“ wurde klarer benannt: Bestehende Berechtigungen am Zielordner werden ersetzt.
  - Validierung legt keine Benutzer mehr an. Nicht in der Site vorhandene Benutzer werden als „⚠️ Nicht in Site“ markiert und beim Anwenden per EnsureUser aufgelöst.

### Behoben
- **Sicherheit**
  - HTML-Exporte (Manager und Reader) kodieren alle Werte (HTML-/Script-Injection über Ordner-, Benutzer- oder Gruppennamen).
  - Manager „Entfernen“ löschte per `PurgeAccessRules` **alle** expliziten Regeln eines Benutzers (inkl. Deny). Jetzt wird nur die ausgewählte Regel entfernt; geerbte Einträge werden übersprungen und gemeldet.
  - Filter-Eingaben werden als Wildcard maskiert.
- **Funktionen**
  - Manager: Regeln ohne Vererbungs-Flags galten nicht für Dateien im Ordner.
  - Manager: `Get-Acl`/`Set-Acl` mit `-LiteralPath` – Pfade mit `[` `]` funktionieren.
  - Manager: Backup-Dateinamen sind gegen ungültige Zeichen abgesichert.
  - Reader: Spalte „Anpassung“ im CSV-Export war wegen überschriebener `$_`-Variable immer falsch.
  - Reader: Regex für Computerkonten korrigiert.
  - Reader: Entfernen eines Ordners entfernt auch die Daten seiner Unterordner.
  - Reader: Einzelne Auswahl wird korrekt als Array behandelt.
  - Beta: `??`-Operator (nur PowerShell 7) verhinderte den Start unter Windows PowerShell 5.1.
  - Beta: `-ClearExisting` wurde pro Eintrag angewendet – pro Ordner blieb nur der letzte Benutzer übrig. Jetzt wird nur einmal pro Ordner und Lauf bereinigt.
  - Beta: Rechte-Zuordnung fehlerhaft – „Write, ReadAndExecute“ wurde als Lesen gesetzt. Numerische Rechte wurden nicht erkannt.
  - Beta: Feste Bibliothek `Documents` und falscher `Add-PnPFolder`-Aufruf ersetzt.
  - Beta: `#` in Ordnernamen schnitt den Pfad ab (URL-Fragment). Berechtigungen landeten dadurch am Eltern-Ordner oder an der Bibliothek.
  - Beta: Übergeordnete Ordner werden vor Unterordnern verarbeitet. Das erstmalige Unterbrechen der Vererbung löscht sonst bereits gesetzte Unterordner-Berechtigungen.
  - Beta: Redundantes `Set-PnPWebPermission "Full Control"` beim Site-Admin entfernt (schlug auf deutschen Sites fehl).
- **Performance**
  - Iterative Ordnersuche ohne Folgen von Junctions/Symlinks; Zugriffsfehler brechen den Scan nicht ab.
  - Listen/`StringBuilder` statt `+=` auf Arrays und Strings.
  - ACL wird pro Ordner nur einmal geschrieben.
  - Beta: Site-Benutzer, Rollen und Bibliotheken werden pro Verbindung zwischengespeichert. Es wird nicht mehr `Get-PnPUser` je Eintrag aufgerufen.
- **GUI**
  - Oberfläche friert bei langen Vorgängen nicht mehr ein (Dispatcher-Pump statt `Start-Sleep`, Wartecursor, unbestimmte Fortschrittsanzeige).
  - Timer wird beim Schließen gestoppt.
  - Dialogsymbole passen zur Aktion.
  - Fehler in Exporten werden abgefangen und angezeigt.
- `Add-Type` wird vor `EnableVisualStyles()` ausgeführt (Start ohne vorgeladene WinForms-Assembly).

### Entfernt
- Alte Versionen: `easyFPManager_V0.0.1.ps1/.exe`, `easyFPReader_V0.0.1.ps1/.exe`, `# Beta/easyFolder-SPOPermissions_V0.0.1.ps1`.
- Alle Screenshots (`Screenshot1_easyFPManager.jpg`, `Screenshot1_easyFPReader.jpg`).
- Vorkompilierte EXE-Dateien aus dem Repository. Binärdateien werden künftig über GitHub Releases verteilt (siehe README).
- Authenticode-Signaturblöcke: Sie waren durch die Änderungen ungültig. Neu signieren siehe [`CONTRIBUTING.md`](CONTRIBUTING.md).
- Beta: ungenutzter Code (`SPOConnectionConfig`, `Show-AppRegistrationHelp`, `$Global:AdminCredentials`, `$Global:ValidatedUsers`).

### Sicherheit
- Siehe Abschnitt „Behoben → Sicherheit“ sowie [`SECURITY.md`](SECURITY.md).

---

## [0.0.1] - 2025-10-24

Erstveröffentlichung ([a7ab86d](https://github.com/PS-PhinIT/PhinIT.PowerShell_Folder/commit/a7ab86d)).

### Hinzugefügt
- **easyFPManager V0.0.1**
  - WPF-Tool zum Anzeigen, Hinzufügen und Entfernen von NTFS-Berechtigungen.
  - Backup, automatische Reports, CSV-/HTML-Export und Live-Filter.
- **easyFPReader V0.0.1**
  - Rekursive Berechtigungsanalyse mit Baumansicht und HTML-Export.
  - OnPrem → Entra ID/SharePoint UPN-Mapping mit CSV-Export.
- **easyFolder-SPOPermissions V0.0.1 (Beta)**
  - Import der Reader-CSV und Anwendung der Berechtigungen in SharePoint Online (PnP.PowerShell).
- Signierte Skripte, mit ps2exe erstellte EXE-Dateien und Screenshots.
- `README.md` in Deutsch und Englisch mit neuen Toolnamen, Installation und Systemanforderungen ([c484924](https://github.com/PS-PhinIT/PhinIT.PowerShell_Folder/commit/c484924)).

### Geändert
- Beta-Skript in den Ordner `# Beta` verschoben ([207f091](https://github.com/PS-PhinIT/PhinIT.PowerShell_Folder/commit/207f091)).

### Nachträge zu 0.0.1 (ohne Versionssprung)

#### 2025-10-30 ([701806b](https://github.com/PS-PhinIT/PhinIT.PowerShell_Folder/commit/701806b))
- **Geändert**: easyFPManager und easyFPReader für ps2exe-GUI-Builds optimiert:
  - `$ProgressPreference = 'SilentlyContinue'`.
  - Visual Styles werden vor der GUI-Erstellung aktiviert.
  - `ShowDialog()`-Rückgaben werden abgefangen.
  - Kompilierhinweise im Skriptkopf ergänzt.
- **Geändert**: EXE-Dateien neu erstellt und Skripte neu signiert.

#### 2025-12-30 ([97c8e4f](https://github.com/PS-PhinIT/PhinIT.PowerShell_Folder/commit/97c8e4f))
- **Geändert**: Fußzeile der README auf „PhinIT.DE“ aktualisiert.

---

[0.0.1]: https://github.com/PS-PhinIT/PhinIT.PowerShell_Folder/commit/a7ab86d
