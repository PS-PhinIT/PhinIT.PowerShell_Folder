# Sicherheitsrichtlinie / Security Policy

## Unterstützte Versionen / Supported versions

| Version | Unterstützt / Supported |
|---------|-------------------------|
| 0.0.2   | ✅ |
| 0.0.1   | ❌ (bitte auf 0.0.2 aktualisieren / please upgrade) |

Das Beta-Tool `easyFolder-SPOPermissions` erhält Sicherheitskorrekturen nach Möglichkeit, aber ohne Zusage.
The beta tool `easyFolder-SPOPermissions` receives security fixes on a best-effort basis.

## Schwachstelle melden / Reporting a vulnerability

Bitte **keine öffentlichen Issues** für Sicherheitslücken anlegen. Nutzen Sie stattdessen:

1. **GitHub Private Vulnerability Reporting**: *Security* → *Report a vulnerability* in diesem Repository, oder
2. Kontakt über [www.phinit.de](https://www.phinit.de).

Bitte angeben: betroffenes Tool und Version, Schritte zur Reproduktion, erwartetes und tatsächliches Verhalten, mögliche Auswirkungen.
Wir bestätigen den Eingang in der Regel innerhalb von 7 Tagen und informieren über den weiteren Verlauf.

Please **do not open public issues** for security vulnerabilities. Use GitHub *Private vulnerability reporting* or contact us via [www.phinit.de](https://www.phinit.de).

## Sicherheitsrelevante Hinweise zur Nutzung / Usage notes

- **easyFPManager** ändert NTFS-Berechtigungen mit Administratorrechten.
  - Vor Änderungen ein Backup erstellen. Wiederherstellen mit `icacls <übergeordneter Ordner> /restore <datei>.acl`.
- **HTML-Reports** kodieren alle Werte (seit 0.0.2). Reports aus 0.0.1 nicht mehr öffnen, wenn Ordnernamen aus unbekannter Quelle stammen.
- **easyFolder-SPOPermissions** speichert in `HKCU:\Software\PhinIT\easyFolder-SPOPermissions` nur:
  - ClientId und TenantId;
  - Site-URL und Admin-UPN.
  Es werden keine Passwörter, Secrets oder Tokens gespeichert.
- **Integrität**: Nur Skripte bzw. EXE-Dateien aus den offiziellen GitHub Releases verwenden.
  - Signatur prüfen mit `Get-AuthenticodeSignature <Datei>`.
