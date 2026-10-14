# Mitwirken / Contributing

Danke für Ihr Interesse an PhinIT.PowerShell_Folder! / Thanks for contributing!

## Projektstruktur
- `src/<Tool>/<Tool>_V<Version>.ps1` – Produktive Tools (eine Datei pro Tool, Version im Dateinamen).
- `src/beta/<Tool>/` – Beta-Tools.
- `docs/` – Weiterführende Dokumentation (z. B. Audit-Berichte).
- **Keine** Binärdateien (EXE), Screenshots, Reports oder Logs einchecken – siehe `.gitignore`.

## Konventionen
- **Kodierung**: Alle `.ps1`-Dateien als **UTF-8 mit BOM** speichern (sonst Umlaute/Emojis unter Windows PowerShell 5.1 fehlerhaft).
- **Zeilenenden**: CRLF für `.ps1` (über `.gitattributes` geregelt).
- **Kompatibilität**: Windows PowerShell 5.1 muss unterstützt werden.
  - Keine PS-7-only-Syntax: kein `??`, `?.`, Ternary `a ? b : c`, Pipeline-Ketten `&&` und `||`.
- **Pfade**: `-LiteralPath` statt `-Path` für Benutzer-/Dateisystempfade.
- **Ausgabe**: Werte in HTML immer mit `[System.Net.WebUtility]::HtmlEncode()` kodieren.
- **Sammlungen**: `System.Collections.Generic.List[object]` statt `+=` auf Arrays in Schleifen.
- **GUI**: Lange Schleifen mit `Update-UI` (Dispatcher-Pump) statt `Start-Sleep` aktualisieren.
- **Funktionsnamen**: Zulässige PowerShell-Verben (`Get-Verb`).

## Version erhöhen
1. Datei umbenennen (`git mv`), z. B. `easyFPManager_V0.0.2.ps1` → `easyFPManager_V0.0.3.ps1`.
2. Version im Skriptkopf (`.SYNOPSIS`, `.VERSION`/`$Global:AppVersion`), im Fenstertitel und im ps2exe-Beispiel anpassen.
3. **CHANGELOG.md** um einen Abschnitt `## [x.y.z] - JJJJ-MM-TT` ergänzen.
   - Einträge immer direkt einer Version mit Datum zuordnen – **keine** Sammelabschnitte ohne Version.
4. README (Pfade/Version) aktualisieren.
5. Nach dem Merge Tag `vx.y.z` setzen und ein GitHub Release mit signierten Skripten und EXE-Dateien erstellen.

## Prüfen vor dem Pull Request
```powershell
# Syntax prüfen
Get-ChildItem .\src -Recurse -Filter *.ps1 | ForEach-Object {
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$null, [ref]$errors) | Out-Null
    if ($errors) { Write-Warning "$($_.Name): $($errors.Message -join '; ')" }
}

# Statische Analyse
Install-Module PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path .\src -Recurse -Severity Warning,Error
```
Zusätzlich jedes geänderte Tool manuell unter **Windows PowerShell 5.1** starten.

## EXE erstellen und signieren (Release)
```powershell
Install-Module ps2exe -Scope CurrentUser

# Signieren (Code-Signing-Zertifikat im Benutzerspeicher)
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
Get-ChildItem .\src -Recurse -Filter *.ps1 | ForEach-Object {
    Set-AuthenticodeSignature -FilePath $_.FullName -Certificate $cert -TimestampServer "http://timestamp.digicert.com" -HashAlgorithm SHA256
}

# EXE bauen (Manager mit -requireAdmin)
Invoke-ps2exe .\src\easyFPManager\easyFPManager_V0.0.2.ps1 .\easyFPManager_V0.0.2.exe -noConsole -STA -x64 -requireAdmin -version "0.0.2.0"
Invoke-ps2exe .\src\easyFPReader\easyFPReader_V0.0.2.ps1  .\easyFPReader_V0.0.2.exe  -noConsole -STA -x64 -version "0.0.2.0"
```
Signierte Skripte und EXE-Dateien als Release-Assets hochladen – **nicht** committen. Die Signatur im Repository würde bei jeder Änderung ungültig.

## Commits & Pull Requests
- Aussagekräftige Commit-Nachrichten (Deutsch oder Englisch).
- Ein Pull Request pro Thema; Beschreibung mit Bezug auf den CHANGELOG-Eintrag.
