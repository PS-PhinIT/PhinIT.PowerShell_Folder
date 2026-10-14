# Audit-Bericht – PhinIT.PowerShell_Folder

| | |
|---|---|
| **Datum** | 2026-10-14 |
| **Geprüfter Stand** | Commit `97c8e4f` (V0.0.1) |
| **Ergebnisstand** | V0.0.2 (siehe [CHANGELOG](../CHANGELOG.md)) |
| **Umfang** | `easyFPManager`, `easyFPReader`, `easyFolder-SPOPermissions` (Beta), Repository-Struktur, Dokumentation |
| **Methodik** | Manuelle Code-Review aller Skripte; Parser-Prüfung (PowerShell-AST) auf Syntaxfehler und PS-7-only-Syntax; XAML-Wohlgeformtheit; Abgleich aller `FindName`-Verweise mit der XAML; Auflösbarkeit aller Befehle; Funktionstests der Kernfunktionen (Pfadumrechnung, Rechte-Mapping, HTML-Kodierung, Filter, Ordner-Scan) |

**Schweregrade:** 🔴 Hoch · 🟠 Mittel · 🟡 Niedrig · ⚪ Info

**Status:**
- ✅ Behoben.
- 📝 Dokumentiert: bewusst nicht geändert, mit Begründung.

---

## 1. Sicherheit

| ID | Tool | Befund | Schwere | Status |
|----|------|--------|---------|--------|
| S-01 | Reader, Manager | HTML-Exporte schrieben Ordner-, Benutzer- und Gruppennamen unkodiert ins HTML. Ein präparierter Ordnername (z. B. `<script>`) wurde beim Öffnen des Reports ausgeführt (Stored XSS). | 🔴 | ✅ Alle Werte per `HtmlEncode`; `meta charset` gesetzt. |
| S-02 | Manager | „Entfernen“ nutzte `PurgeAccessRules` und löschte damit **alle** expliziten Regeln des Benutzers inkl. Deny-Regeln, nicht nur die ausgewählte. Bei geerbten Einträgen wurde ein Erfolg gemeldet, ohne dass sich etwas änderte. | 🔴 | ✅ Nur die exakt ausgewählte Regel wird entfernt (Identität, Rechte, Typ, Flags). Geerbte Einträge werden übersprungen und gemeldet. |
| S-03 | Manager | Benutzer wurden als `NTAccount`-Name verarbeitet. Mehrdeutige oder lokalisierte Namen (z. B. „Jeder“/„Everyone“) konnten fehlschlagen oder falsch zuordnen. | 🟠 | ✅ Auflösung über SID (`Resolve-Principal`) mit klarer Fehlermeldung. |
| S-04 | Manager | Das „Backup“ war nur ein XML-Export der Anzeige und nicht wiederherstellbar. | 🟠 | ✅ Zusätzlich `icacls /save /t /c` (`.acl`). Wiederherstellung siehe README. |
| S-05 | Repo | EXE-Dateien lagen im Repository. Sie waren nicht reproduzierbar, nicht prüfbar und nach Skriptänderungen veraltet. Die Authenticode-Signaturen werden durch jede Änderung ungültig. | 🟠 | ✅ EXE und Signaturblöcke entfernt. Build/Signatur in CONTRIBUTING dokumentiert; Binärdateien über GitHub Releases. |
| S-06 | Manager | Filtereingaben gingen unmaskiert in `-like` (Wildcard-Injection, unerwartete Treffer). | 🟡 | ✅ `WildcardPattern::Escape`. |
| S-07 | Beta | Pfade aus der CSV wurden ungeprüft verwendet. Eine URL zu einer anderen Site oder einem anderen Tenant führte zu unerwarteten Zielen. | 🟠 | ✅ Tenant- und Site-Grenze wird geprüft, sonst Fehler pro Eintrag. |
| S-08 | Beta | Die Registry (`HKCU:\Software\PhinIT\easyFolder-SPOPermissions`) speichert ClientId, TenantId, SiteURL und AdminUPN. | ⚪ | 📝 Keine Geheimnisse gespeichert. Anmeldung interaktiv bzw. per Device Code, keine Passwörter oder Secrets. |
| S-09 | Reader, Beta | CSV-Exporte könnten theoretisch Formeln enthalten (CSV-/Formel-Injection in Excel). | 🟡 | 📝 Werte stammen aus AD-/NTFS-Namen, die nicht mit `=`, `+`, `-`, `@` beginnen können. Kein Handlungsbedarf. |
| S-10 | Manager | Das Skript erfordert Administratorrechte. | ⚪ | 📝 Durch `#Requires -RunAsAdministrator` erzwungen; für ps2exe `-requireAdmin` dokumentiert. |

## 2. Performance

| ID | Tool | Befund | Schwere | Status |
|----|------|--------|---------|--------|
| P-01 | Manager | Hinzufügen schrieb eine explizite Regel auf **jeden** Unterordner (`Get-ChildItem -Recurse` + `Set-Acl` je Ordner). Das war sehr langsam bei großen Strukturen und blähte die ACLs auf. | 🔴 | ✅ Eine vererbbare Regel. Explizite Regeln nur dort, wo die Vererbung deaktiviert ist. |
| P-02 | alle | Arrays und Strings wurden mit `+=` aufgebaut (O(n²)). | 🟠 | ✅ `List[object]` und `StringBuilder`. |
| P-03 | Manager, Reader | Bei rekursiven Scans konnten Junctions/Symlinks zu Doppel- oder Endlosschleifen führen. Ein Zugriffsfehler brach den ganzen Scan ab. | 🟠 | ✅ Iterativer Scan ohne Reparse-Points; Zugriffsfehler werden übersprungen. |
| P-04 | Manager | Beim Entfernen wurde die ACL je Regel geschrieben. | 🟡 | ✅ Eine ACL-Schreiboperation pro Ordner, nur bei tatsächlicher Änderung. |
| P-05 | Beta | `Get-PnPUser` (alle Site-Benutzer) wurde pro CSV-Eintrag geladen; dazu kam `Start-Sleep` pro Eintrag. | 🟠 | ✅ Cache für Benutzer, Rollen und Bibliotheken pro Verbindung; kein `Start-Sleep`. |
| P-06 | Manager | DataGrid ohne Virtualisierung. | 🟡 | ✅ Zeilen-Virtualisierung aktiviert. |

## 3. Funktionen

| ID | Tool | Befund | Schwere | Status |
|----|------|--------|---------|--------|
| F-01 | Manager | Regeln wurden ohne Vererbungs-Flags gesetzt. Dateien im Ordner erhielten die Berechtigung nicht, und „Nur Hauptordner“, „Nur Unterordner“ und „Beide“ unterschieden sich kaum. | 🔴 | ✅ Saubere Semantik über InheritanceFlags und PropagationFlags (siehe README). |
| F-02 | Beta | `-ClearExisting` lief bei **jedem** Eintrag. Pro Ordner blieb nur der zuletzt verarbeitete Benutzer übrig (Datenverlust). | 🔴 | ✅ Bereinigung nur einmal pro Ordner und Lauf. |
| F-03 | Beta | Falsche Reihenfolge im Rechte-Mapping: „Write, ReadAndExecute“ wurde als Lesen gesetzt. Numerische Rechte (z. B. `268435456`) wurden nicht erkannt. | 🔴 | ✅ Reihenfolge Vollzugriff → Ändern → Schreiben → Lesen; numerische Werte werden erkannt. |
| F-04 | Beta | Die Reader-CSV enthält vollständige URLs, das Skript erwartete site-relative Pfade. Ordner und Berechtigungen schlugen daher immer fehl. | 🔴 | ✅ `ConvertTo-SPOSiteRelativePath`. |
| F-05 | Beta | Die Bibliothek war fest als `"Documents"` hinterlegt, und `Add-PnPFolder -Name <Pfad> -Folder "/"` war falsch. | 🔴 | ✅ Bibliothek wird aus dem Pfad ermittelt; `Resolve-PnPFolder`. |
| F-06 | Beta | Englische Rollennamen („Full Control“, „Edit“ …) waren hart codiert. Auf deutschsprachigen Sites schlug das fehl. | 🟠 | ✅ Auflösung über `RoleTypeKind` (`Get-PnPRoleDefinition`). |
| F-07 | Reader | Die Spalte „Anpassung“ im CSV war wegen eines überschriebenen `$_` im verschachtelten `Where-Object` immer falsch. | 🟠 | ✅ |
| F-08 | Reader | Der Regex für Computerkonten (`$` am Ende) griff nicht. Systemkonten wurden unvollständig erkannt. | 🟡 | ✅ |
| F-09 | Reader | Beim Entfernen eines Stammordners blieben die Daten seiner Unterordner in der Analyse. | 🟡 | ✅ |
| F-10 | Reader | Die URL-Umrechnung kannte nur einen Stammordner. | 🟡 | ✅ Längster passender Stammordner. |
| F-11 | Manager, Reader | `-Path` statt `-LiteralPath`: Ordner mit `[` oder `]` im Namen funktionierten nicht. | 🟠 | ✅ |
| F-12 | Reader | Gruppennamen mit Leerzeichen erzeugen ungültige UPNs. | 🟡 | 📝 Bekannte Einschränkung. Lösung über manuelle Zuordnung (Tab „Benutzer-Mapping“). |
| F-13 | Beta | „Vererbung unterbrechen“ war irreführend: Ordner-Berechtigungen unterbrechen in SharePoint **immer** die Vererbung. | 🟡 | ✅ Option umbenannt und mit Tooltip versehen: „Bestehende Berechtigungen ersetzen“. |
| F-14 | Beta | Ein `#` im Ordnernamen wurde beim Parsen der URL als Fragment abgeschnitten. Die Berechtigung (ggf. mit `ClearExisting`) landete dadurch am Eltern-Ordner oder an der ganzen Bibliothek. | 🔴 | ✅ Pfad wird ohne `System.Uri` zerlegt; nur prozentkodierte Browser-URLs werden dekodiert. |
| F-15 | Beta | Das erstmalige Unterbrechen der Vererbung (`clearSubscopes`) löschte bereits gesetzte Berechtigungen von Unterordnern, wenn diese in der CSV vor dem Eltern-Ordner standen. | 🟠 | ✅ Verarbeitung nach Pfadtiefe (Eltern zuerst, sonst CSV-Reihenfolge). |

## 4. Verweise (Referenzen, Pfade, Abhängigkeiten)

| ID | Tool | Befund | Schwere | Status |
|----|------|--------|---------|--------|
| V-01 | Repo | Die README verwies auf Skripte im Wurzelverzeichnis und auf V0.0.1. | 🟡 | ✅ README auf `src/…` und V0.0.2 aktualisiert. |
| V-02 | Reader | Die Hinweistexte nannten „Tab 2“ und „Tab 3“, aber die Reihenfolge der Tabs war vertauscht. | 🟡 | ✅ Tabs getauscht. |
| V-03 | alle | Alle `FindName`-Verweise wurden gegen die XAML geprüft. | ⚪ | ✅ Keine fehlenden Controls. |
| V-04 | Beta | Ungenutzter Code: `SPOConnectionConfig`, `Show-AppRegistrationHelp`, `$Global:AdminCredentials`, `$Global:ValidatedUsers`. `Test-SPOConnection` war ebenfalls ungenutzt. | 🟡 | ✅ Toter Code entfernt; `Test-SPOConnection` wird jetzt zur URL-Prüfung verwendet. |
| V-05 | Beta | Die PnP-„Management Shell“-App wurde von Microsoft entfernt (2024). `-UseWebLogin` gibt es nur in PnP.PowerShell 1.x. | ⚪ | 📝 Eigene App-Registrierung ist Pflicht (Setup-Dialog vorhanden). Der Fallback bleibt für 1.x erhalten. |

## 5. GUI

| ID | Tool | Befund | Schwere | Status |
|----|------|--------|---------|--------|
| G-01 | alle | Die GUI fror bei langen Vorgängen ein; Fortschritt und Status wurden nicht aktualisiert (`Start-Sleep` im UI-Thread). | 🟠 | ✅ Dispatcher-Pump (`Update-UI`), Wartecursor, unbestimmte Fortschrittsanzeige. |
| G-02 | Manager | `Out-GridView` für die Benutzerauswahl war in EXE-Builds und PowerShell 7 nicht verfügbar. | 🟠 | ✅ Eigener WPF-Auswahldialog. |
| G-03 | Manager | Hinzufügen erfolgte ohne Rückfrage; Lösch-Dialoge nutzten ein neutrales Symbol. | 🟡 | ✅ Bestätigung mit Zusammenfassung; Warnsymbol beim Löschen. |
| G-04 | Manager | Der Timer lief nach dem Schließen weiter. | 🟡 | ✅ Wird bei `Closed` gestoppt. |
| G-05 | Beta | Der Hinweis „Admin wird automatisch hinzugefügt“ galt nur bei aktivierter Option. | 🟡 | ✅ Text korrigiert. |
| G-06 | alle | Dateien ohne BOM zeigten unter Windows PowerShell 5.1 Umlaute und Emojis verstümmelt an. | 🟠 | ✅ UTF-8 mit BOM; `.gitattributes`. |

## 6. Befehle (Cmdlets, Syntax, Kompatibilität)

| ID | Tool | Befund | Schwere | Status |
|----|------|--------|---------|--------|
| B-01 | Beta | Der `??`-Operator (nur PowerShell 7) führte unter 5.1 trotz `#Requires -Version 5.1` zu einem Parserfehler. | 🔴 | ✅ 5.1-kompatibel umgeschrieben. Alle Skripte sind auf PS-7-only-Syntax geprüft. |
| B-02 | Manager | `Create-PermissionReport` verwendete ein nicht zulässiges Verb. | ⚪ | ✅ `New-PermissionReport`. |
| B-03 | Manager, Reader | `EnableVisualStyles()` wurde vor `Add-Type System.Windows.Forms` aufgerufen und schlug in einer frischen Sitzung fehl. | 🟠 | ✅ Reihenfolge korrigiert. |
| B-04 | Beta | `Set-PnPWebPermission -AddRole "Full Control"` beim Site-Admin war redundant und auf lokalisierten Sites fehlerhaft. | 🟡 | ✅ Nur noch `Add-PnPSiteCollectionAdmin`. |
| B-05 | Beta | `Import-Module PnP.PowerShell -Force` bei jeder Prüfung konnte eine bestehende Verbindung zurücksetzen. | 🟡 | ✅ Nur noch laden, wenn das Modul noch nicht geladen ist. |
| B-06 | alle | PSScriptAnalyzer war in der Prüfumgebung nicht verfügbar. | ⚪ | 📝 Ersatzweise Parser- und AST-Prüfungen. Empfehlung: lokal `Invoke-ScriptAnalyzer -Path .\src -Recurse` (siehe CONTRIBUTING). |

## 7. Repository & Dokumentation

| ID | Befund | Status |
|----|--------|--------|
| R-01 | Es fehlten `CHANGELOG.md`, `SECURITY.md`, `CONTRIBUTING.md`, `.gitignore` und `.gitattributes`. | ✅ Angelegt. Der CHANGELOG wurde aus der Git-Historie rekonstruiert (keine PRs vorhanden). |
| R-02 | Alte Versionen (V0.0.1 als `.ps1` und `.exe`) und Screenshots lagen im Wurzelverzeichnis. | ✅ Entfernt. Neue Struktur `src/`, `src/beta/`, `docs/`. |
| R-03 | Der Ordnername `# Beta` (Leerzeichen und `#`) war problematisch in Shells und URLs. | ✅ `src/beta/`. |
| R-04 | Keine Lizenzdatei vorhanden. | 📝 Entscheidung des Rechteinhabers; nicht automatisch ergänzt. |

---

## Empfehlungen für Folgeversionen
- **Tag und Release**: Nach dem Merge den Tag `v0.0.2` setzen. Die signierten Skripte und EXE-Dateien als Release-Assets veröffentlichen.
- **CI**: GitHub Action mit PSScriptAnalyzer und Pester auf `windows-latest`.
- **Beta-Tool**: Gruppen-Mapping (SharePoint-/M365-Gruppen statt UPN) und Testlauf-Modus (Dry-Run) ergänzen.
