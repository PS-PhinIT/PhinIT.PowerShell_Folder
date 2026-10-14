#Requires -Version 5.1
#Requires -RunAsAdministrator

# PS2EXE OPTIMIERUNGEN - KRITISCH FÜR GUI-EXE
$ProgressPreference = 'SilentlyContinue'  # Verhindert blinkende Progress-Fenster bei -noConsole

<#
.SYNOPSIS
    Easy Folder Permissions Manager (easyFPManager) V0.0.2
    
.DESCRIPTION
    PowerShell Script mit WPF GUI zur Verwaltung von Ordnerberechtigungen
    - Anzeige aktueller Berechtigungen (inkl. Vererbungs-/Propagierungs-Flags)
    - Hinzufügen von Benutzern mit spezifischen Rechten
    - Unterstützung für Hauptordner, Unterordner oder beide
    - Gezieltes Entfernen einzelner (expliziter) Berechtigungseinträge
    - Wiederherstellbares Backup (icacls /save) + lesbares XML-Backup
    - Automatische Report-Erstellung bei Änderungen
    
.AUTHOR
    PhinIT Development
    
.VERSION
    0.0.2
    
.NOTES
    WICHTIG FÜR EXE-KONVERTIERUNG MIT PS2EXE:
    - Alle Debug-Meldungen werden AUSSCHLIESSLICH in die Log-Datei geschrieben
    - Keine Write-Host, Write-Error oder Write-Warning Ausgaben in der Shell!
    - Debug-Logfile: %USERPROFILE%\Documents\FolderPermissions_Reports\Debug.log
    - $ProgressPreference = 'SilentlyContinue' am Anfang (verhindert Progress-Fenster)
    - Visual Styles direkt nach dem Laden der Assemblies und VOR GUI-Erstellung aktiviert
    - ShowDialog() mit [VOID] versehen (verhindert "False" MessageBox)
    - Datei ist UTF-8 mit BOM gespeichert (deutsche Umlaute unter Windows PowerShell 5.1)
    
    KOMPILIERUNG:
    Invoke-ps2exe .\easyFPManager_V0.0.2.ps1 .\easyFPManager_V0.0.2.exe -noConsole -STA -x64 -requireAdmin `
           -title "Easy Folder Permissions Manager" `
           -version "0.0.2.0" -company "PhinIT" -copyright "© 2025 Andreas Hepp"
#>

# Assemblies laden (Out-Null verhindert MessageBoxen bei -noConsole)
Add-Type -AssemblyName PresentationFramework | Out-Null
Add-Type -AssemblyName PresentationCore | Out-Null
Add-Type -AssemblyName WindowsBase | Out-Null
Add-Type -AssemblyName System.Windows.Forms | Out-Null

# Visual Styles aktivieren (erst NACH dem Laden von System.Windows.Forms möglich, aber VOR allen GUI-Objekten)
[System.Windows.Forms.Application]::EnableVisualStyles()

# Globale Variablen
$Global:AppVersion = "0.0.2"
$Global:ReportsPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "FolderPermissions_Reports"
$Global:CurrentFolder = ""
$Global:PermissionsData = @()

# Logfile-Pfad für Debug-Informationen
$Global:LogFilePath = Join-Path $Global:ReportsPath "Debug.log"

# Logging-Funktion (nur in Datei, nicht in Shell)
function Write-LogEntry {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        Add-Content -LiteralPath $Global:LogFilePath -Value $logMessage -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {
        # Fehler beim Schreiben ins Log werden ignoriert
    }
}

# Report-Ordner erstellen falls nicht vorhanden
try {
    if (-not (Test-Path -LiteralPath $Global:ReportsPath)) {
        New-Item -Path $Global:ReportsPath -ItemType Directory -Force | Out-Null
    }
}
catch {
    [System.Windows.MessageBox]::Show("Report-Ordner konnte nicht erstellt werden: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
}

# ========== HILFSFUNKTIONEN ==========

# Wandelt einen Benutzer-/Gruppennamen (oder eine SID) in eine SecurityIdentifier um.
# Wirft eine aussagekräftige Exception, falls das Konto nicht aufgelöst werden kann.
function Resolve-Principal {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    
    $trimmed = $Name.Trim()
    if ($trimmed -match '^S-1-[0-9-]+$') {
        return New-Object System.Security.Principal.SecurityIdentifier($trimmed)
    }
    
    try {
        $account = New-Object System.Security.Principal.NTAccount($trimmed)
        return $account.Translate([System.Security.Principal.SecurityIdentifier])
    }
    catch {
        throw "Benutzer oder Gruppe '$trimmed' konnte nicht aufgelöst werden. Bitte Schreibweise (z.B. DOMAIN\Benutzer) prüfen."
    }
}

# Liefert alle Unterordner (rekursiv) eines Ordners.
# Iterativ (kein Stack-Overflow bei tiefen Strukturen), überspringt Junctions/Symlinks
# (verhindert Endlosschleifen) und setzt bei "Zugriff verweigert" mit dem nächsten Ordner fort.
function Get-SubDirectoryList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )
    
    $result = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($FolderPath)
    
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            $children = ([System.IO.DirectoryInfo]::new($current)).GetDirectories()
        }
        catch {
            Write-LogEntry "Ordner konnte nicht gelesen werden: $current - $($_.Exception.Message)" "WARN"
            continue
        }
        
        foreach ($child in $children) {
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                Write-LogEntry "Reparse-Point übersprungen: $($child.FullName)" "INFO"
                continue
            }
            $result.Add($child.FullName)
            $stack.Push($child.FullName)
        }
    }
    
    return ,$result
}

# Entfernt alle expliziten Allow/Deny-Einträge eines Kontos aus einem Ordner.
# Schreibt die ACL nur zurück, wenn tatsächlich etwas geändert wurde (Performance).
function Remove-ExplicitRulesForSid {
    param(
        [string]$Path,
        [System.Security.Principal.SecurityIdentifier]$Sid
    )
    
    $acl = Get-Acl -LiteralPath $Path
    $hasRules = $false
    foreach ($rule in $acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])) {
        if ($rule.IdentityReference -eq $Sid) { $hasRules = $true; break }
    }
    
    if ($hasRules) {
        $acl.PurgeAccessRules($Sid)
        Set-Acl -LiteralPath $Path -AclObject $acl
        return $true
    }
    return $false
}

# Erweiterte Funktionen
function Remove-FolderPermission {
    param(
        [string]$FolderPath,
        [string]$Username,
        [string]$ApplyTo
    )
    
    try {
        $sid = Resolve-Principal -Name $Username
        $targets = New-Object System.Collections.Generic.List[string]
        
        if ($ApplyTo -eq "MainOnly" -or $ApplyTo -eq "Both") {
            $targets.Add($FolderPath)
        }
        if ($ApplyTo -eq "SubOnly" -or $ApplyTo -eq "Both") {
            foreach ($dir in (Get-SubDirectoryList -FolderPath $FolderPath)) { $targets.Add($dir) }
        }
        
        $changed = 0
        $errors = New-Object System.Collections.Generic.List[string]
        foreach ($target in $targets) {
            try {
                if (Remove-ExplicitRulesForSid -Path $target -Sid $sid) { $changed++ }
            }
            catch {
                $errors.Add("$target : $($_.Exception.Message)")
                Write-LogEntry "Entfernen fehlgeschlagen auf $target : $($_.Exception.Message)" "ERROR"
            }
        }
        
        New-PermissionReport -Action "REMOVE" -Path $FolderPath -User $Username -Rights "ALL (explizit)" -ApplyTo $ApplyTo
        
        if ($errors.Count -gt 0) {
            $details = ($errors | Select-Object -First 20) -join "`n"
            [System.Windows.MessageBox]::Show("Berechtigungen wurden auf $changed Ordner(n) entfernt, $($errors.Count) Fehler:`n`n$details", "Teilweise erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
        
        return $true
    }
    catch {
        Write-LogEntry "Fehler beim Entfernen der Berechtigung: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Entfernen der Berechtigung: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return $false
    }
}

function Backup-FolderPermissions {
    param(
        [string]$FolderPath
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        # Abschließenden Backslash entfernen (außer bei Laufwerkswurzel), sonst Quoting-Problem bei icacls
        if ($FolderPath.Length -gt 3) { $FolderPath = $FolderPath.TrimEnd('\') }
        $safeName = ($FolderPath -replace '[\\/:*?"<>|]', '_').Trim('_')
        if ($safeName.Length -gt 80) { $safeName = $safeName.Substring($safeName.Length - 80) }
        $baseName = Join-Path $Global:ReportsPath "Backup_$($safeName)_$timestamp"
        
        # 1) Lesbares XML-Backup (Dokumentation)
        $xmlFile = "$baseName.xml"
        $permissions = Get-FolderPermissions -FolderPath $FolderPath -IncludeSubfolders $true
        $permissions | Export-Clixml -LiteralPath $xmlFile
        
        # 2) Wiederherstellbares ACL-Backup (icacls /save)
        #    Wiederherstellung: icacls "<Übergeordneter Ordner>" /restore "<Datei>.acl"
        $aclFile = "$baseName.acl"
        $icaclsOutput = & icacls.exe $FolderPath /save $aclFile /t /c /q 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-LogEntry "icacls /save meldete Exitcode $LASTEXITCODE : $($icaclsOutput -join ' ')" "WARN"
        }
        
        return $xmlFile
    }
    catch {
        Write-LogEntry "Fehler beim Erstellen des Backups: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Erstellen des Backups: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return $null
    }
}

function Get-FilteredPermissions {
    param(
        [array]$Permissions,
        [string]$UserFilter = "",
        [string]$RightsFilter = "",
        [bool]$ShowInheritedOnly = $false
    )
    
    $filtered = $Permissions
    
    # Platzhalterzeichen ([ ] * ?) aus Benutzereingaben maskieren, damit -like nicht fehlschlägt
    if (-not [string]::IsNullOrEmpty($UserFilter)) {
        $userPattern = "*$([System.Management.Automation.WildcardPattern]::Escape($UserFilter))*"
        $filtered = $filtered | Where-Object { $_.User -like $userPattern }
    }
    
    if (-not [string]::IsNullOrEmpty($RightsFilter)) {
        $rightsPattern = "*$([System.Management.Automation.WildcardPattern]::Escape($RightsFilter))*"
        $filtered = $filtered | Where-Object { "$($_.Rights)" -like $rightsPattern }
    }
    
    if ($ShowInheritedOnly) {
        $filtered = $filtered | Where-Object { $_.Inherited -eq $true }
    }
    
    # Immer ein Array zurückgeben (DataGrid.ItemsSource benötigt IEnumerable)
    return ,@($filtered)
}

function Get-SystemUsers {
    try {
        $users = New-Object System.Collections.Generic.List[object]
        
        # Lokale Benutzer
        $localUsers = Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -eq $true }
        foreach ($user in $localUsers) {
            $users.Add([PSCustomObject]@{
                Name = $user.Name
                FullName = if ($user.FullName) { "$($user.Name) ($($user.FullName))" } else { $user.Name }
                Type = "Lokaler Benutzer"
                SID = "$($user.SID)"
            })
        }
        
        # Lokale Gruppen
        $localGroups = Get-LocalGroup -ErrorAction SilentlyContinue
        foreach ($group in $localGroups) {
            $users.Add([PSCustomObject]@{
                Name = $group.Name
                FullName = if ($group.Description) { "$($group.Name) ($($group.Description))" } else { $group.Name }
                Type = "Lokale Gruppe"
                SID = "$($group.SID)"
            })
        }
        
        # Bekannte System-Accounts über Well-Known-SIDs hinzufügen.
        # Die Namen werden in die Systemsprache übersetzt (z.B. "Jeder" statt "Everyone" auf deutschem Windows),
        # da englische Namen auf lokalisierten Systemen nicht aufgelöst werden können.
        $wellKnownSids = @(
            @{ Sid = "S-1-1-0";      Type = "System-Gruppe" },   # Everyone / Jeder
            @{ Sid = "S-1-5-11";     Type = "System-Gruppe" },   # Authenticated Users
            @{ Sid = "S-1-5-18";     Type = "System-Account" },  # SYSTEM
            @{ Sid = "S-1-3-0";      Type = "System-Account" },  # CREATOR OWNER
            @{ Sid = "S-1-5-32-544"; Type = "Lokale Gruppe" },   # Administrators
            @{ Sid = "S-1-5-32-545"; Type = "Lokale Gruppe" }    # Users
        )
        
        foreach ($account in $wellKnownSids) {
            try {
                $sidObject = New-Object System.Security.Principal.SecurityIdentifier($account.Sid)
                $localizedName = $sidObject.Translate([System.Security.Principal.NTAccount]).Value
                $shortName = ($localizedName -split '\\')[-1]
                if (-not ($users | Where-Object { $_.SID -eq $account.Sid -or $_.Name -eq $shortName })) {
                    $users.Add([PSCustomObject]@{
                        Name = $localizedName
                        FullName = $localizedName
                        Type = $account.Type
                        SID = $account.Sid
                    })
                }
            }
            catch {
                Write-LogEntry "Well-Known-SID $($account.Sid) konnte nicht übersetzt werden: $($_.Exception.Message)" "WARN"
            }
        }
        
        return @($users | Sort-Object Type, Name)
    }
    catch {
        Write-LogEntry "Fehler beim Laden der Benutzer: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Einfacher WPF-Auswahldialog für Benutzer/Gruppen
# (ersetzt Out-GridView, das in PS2EXE-EXEs und PowerShell 7 ohne Zusatzmodul nicht verfügbar ist)
function Show-PrincipalPicker {
    param(
        [array]$Principals,
        $Owner
    )
    
    $pickerXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Benutzer oder Gruppe auswählen" Height="520" Width="720"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize"
        Background="#F3F3F3" FontFamily="Segoe UI" FontSize="13">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Label Content="Suche:" Grid.Column="0" VerticalAlignment="Center"/>
            <TextBox Name="txtPickerFilter" Grid.Column="1" Height="28" VerticalContentAlignment="Center"/>
        </Grid>
        <DataGrid Name="dgPicker" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True"
                  SelectionMode="Single" GridLinesVisibility="Horizontal" HeadersVisibility="Column">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="*"/>
                <DataGridTextColumn Header="Beschreibung" Binding="{Binding FullName}" Width="3*"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnPickerOk" Content="Übernehmen" Width="110" Height="30" Margin="0,0,8,0" IsDefault="True"
                    Background="#0078D4" Foreground="White" BorderThickness="0"/>
            <Button Name="btnPickerCancel" Content="Abbrechen" Width="100" Height="30" IsCancel="True"
                    Background="#D13438" Foreground="White" BorderThickness="0"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    $pickerReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$pickerXaml)
    $picker = [Windows.Markup.XamlReader]::Load($pickerReader)
    $pickerReader.Close()
    if ($Owner) { $picker.Owner = $Owner }
    
    $txtPickerFilter = $picker.FindName("txtPickerFilter")
    $dgPicker = $picker.FindName("dgPicker")
    $btnPickerOk = $picker.FindName("btnPickerOk")
    
    $dgPicker.ItemsSource = @($Principals)
    $script:PickerSelection = $null
    
    $txtPickerFilter.Add_TextChanged({
        $text = $txtPickerFilter.Text
        if ([string]::IsNullOrWhiteSpace($text)) {
            $dgPicker.ItemsSource = @($Principals)
        } else {
            $dgPicker.ItemsSource = @($Principals | Where-Object {
                $_.Name.IndexOf($text, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                "$($_.FullName)".IndexOf($text, [StringComparison]::OrdinalIgnoreCase) -ge 0
            })
        }
    })
    
    $confirmSelection = {
        if ($dgPicker.SelectedItem) {
            $script:PickerSelection = $dgPicker.SelectedItem
            $picker.DialogResult = $true
        }
    }
    $btnPickerOk.Add_Click($confirmSelection)
    $dgPicker.Add_MouseDoubleClick($confirmSelection)
    
    [void]$picker.ShowDialog()
    return $script:PickerSelection
}

function Remove-SelectedPermissions {
    param(
        [array]$SelectedPermissions,
        [string]$FolderPath
    )
    
    $successCount = 0
    $errorCount = 0
    $errors = New-Object System.Collections.Generic.List[string]
    
    # Vererbte Einträge können nicht direkt entfernt werden (nur am Ursprungsordner)
    $explicitPermissions = New-Object System.Collections.Generic.List[object]
    foreach ($permission in $SelectedPermissions) {
        if ($permission.Inherited) {
            $errorCount++
            $errors.Add("$($permission.User) auf $($permission.Folder): Vererbter Eintrag - muss am übergeordneten Ordner entfernt werden")
        } else {
            $explicitPermissions.Add($permission)
        }
    }
    
    # Pro Ordner nur einmal ACL lesen/schreiben (Performance)
    foreach ($group in ($explicitPermissions | Group-Object -Property Folder)) {
        try {
            $acl = Get-Acl -LiteralPath $group.Name
            # Gleiche Namensauflösung wie bei der Anzeige (NTAccount bzw. SID-String bei nicht auflösbaren Konten)
            $rules = @($acl.GetAccessRules($true, $false, [System.Security.Principal.NTAccount]))
            $modified = $false
            
            foreach ($permission in $group.Group) {
                try {
                    # Nur den exakt ausgewählten Eintrag entfernen (nicht alle Einträge des Benutzers)
                    $match = $rules | Where-Object {
                        $_.IdentityReference.Value -eq $permission.User -and
                        $_.FileSystemRights -eq $permission.Rights -and
                        $_.AccessControlType -eq $permission.AccessType -and
                        $_.InheritanceFlags -eq $permission.InheritanceFlags -and
                        $_.PropagationFlags -eq $permission.PropagationFlags
                    } | Select-Object -First 1
                    
                    if ($match) {
                        $acl.RemoveAccessRuleSpecific($match)
                        $modified = $true
                        $successCount++
                    } else {
                        $errorCount++
                        $errors.Add("$($permission.User) auf $($permission.Folder): Eintrag nicht mehr vorhanden (bitte aktualisieren)")
                    }
                }
                catch {
                    $errorCount++
                    $errors.Add("$($permission.User) auf $($permission.Folder): $($_.Exception.Message)")
                }
            }
            
            if ($modified) {
                Set-Acl -LiteralPath $group.Name -AclObject $acl
            }
        }
        catch {
            foreach ($permission in $group.Group) {
                $errorCount++
                $errors.Add("$($permission.User) auf $($permission.Folder): $($_.Exception.Message)")
            }
        }
    }
    
    # Report erstellen
    if ($successCount -gt 0) {
        New-PermissionReport -Action "BULK_REMOVE" -Path $FolderPath -User "Multiple ($successCount)" -Rights "Various" -ApplyTo "Selected"
    }
    
    foreach ($err in $errors) { Write-LogEntry "Entfernen: $err" "WARN" }
    
    return @{
        Success = $successCount
        Errors = $errorCount
        ErrorDetails = $errors.ToArray()
    }
}

# HTML-sicheres Encoding (verhindert HTML-/Script-Injection über Ordner- oder Kontonamen)
function ConvertTo-HtmlSafe {
    param([object]$Value)
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

function Export-PermissionsToHtml {
    param(
        [array]$Permissions,
        [string]$FilePath
    )
    
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(@"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Ordnerberechtigungen Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .inherited { color: #666; font-style: italic; }
        .allow { color: #107C10; }
        .deny { color: #D13438; }
    </style>
</head>
<body>
    <h1>Ordnerberechtigungen Report</h1>
    <p>Erstellt am: $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")</p>
    <p>Anzahl Einträge: $($Permissions.Count)</p>
    
    <table>
        <tr>
            <th>Ordner</th>
            <th>Benutzer/Gruppe</th>
            <th>Rechte</th>
            <th>Typ</th>
            <th>Vererbt</th>
            <th>Vererbung an</th>
        </tr>

"@)
    
    foreach ($perm in $Permissions) {
        $inheritedClass = if ($perm.Inherited) { "inherited" } else { "" }
        $accessClass = if ("$($perm.AccessType)" -eq "Allow") { "allow" } else { "deny" }
        $inheritedText = if ($perm.Inherited) { "Ja" } else { "Nein" }
        
        [void]$sb.Append(@"
        <tr class="$inheritedClass">
            <td>$(ConvertTo-HtmlSafe $perm.Folder)</td>
            <td>$(ConvertTo-HtmlSafe $perm.User)</td>
            <td>$(ConvertTo-HtmlSafe $perm.Rights)</td>
            <td class="$accessClass">$(ConvertTo-HtmlSafe $perm.AccessType)</td>
            <td>$inheritedText</td>
            <td>$(ConvertTo-HtmlSafe $perm.InheritanceFlags) / $(ConvertTo-HtmlSafe $perm.PropagationFlags)</td>
        </tr>

"@)
    }
    
    [void]$sb.Append(@"
    </table>
</body>
</html>
"@)
    
    $sb.ToString() | Out-File -LiteralPath $FilePath -Encoding UTF8
}

# Hauptfunktionen
function Get-FolderPermissions {
    param(
        [string]$FolderPath,
        [bool]$IncludeSubfolders = $false
    )
    
    # List statt "+=" (vermeidet quadratische Laufzeit bei großen Ordnerstrukturen)
    $permissions = New-Object System.Collections.Generic.List[object]
    
    try {
        $folders = New-Object System.Collections.Generic.List[string]
        $folders.Add($FolderPath)
        if ($IncludeSubfolders) {
            foreach ($dir in (Get-SubDirectoryList -FolderPath $FolderPath)) { $folders.Add($dir) }
        }
        
        foreach ($folder in $folders) {
            try {
                $acl = Get-Acl -LiteralPath $folder -ErrorAction Stop
            }
            catch {
                Write-LogEntry "ACL konnte nicht gelesen werden: $folder - $($_.Exception.Message)" "WARN"
                continue
            }
            
            foreach ($access in $acl.Access) {
                $permissions.Add([PSCustomObject]@{
                    Folder = $folder
                    User = $access.IdentityReference.Value
                    Rights = $access.FileSystemRights
                    AccessType = $access.AccessControlType
                    Inherited = $access.IsInherited
                    InheritanceFlags = $access.InheritanceFlags
                    PropagationFlags = $access.PropagationFlags
                })
            }
        }
    }
    catch {
        Write-LogEntry "Fehler beim Lesen der Berechtigungen: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Lesen der Berechtigungen: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
    
    return ,$permissions.ToArray()
}

function Add-FolderPermission {
    param(
        [string]$FolderPath,
        [string]$Username,
        [string]$Rights,
        [string]$ApplyTo
    )
    
    try {
        # Konto vorab auflösen -> klare Fehlermeldung statt "IdentityNotMappedException"
        $sid = Resolve-Principal -Name $Username
        $rightsEnum = [System.Security.AccessControl.FileSystemRights]$Rights
        $allow = [System.Security.AccessControl.AccessControlType]::Allow
        
        $ciOi = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
        $noPropagation = [System.Security.AccessControl.PropagationFlags]::None
        
        # "Dieser Ordner, Unterordner und Dateien" - Standard-Regel für Ordnerfreigaben
        $inheritableRule = New-Object System.Security.AccessControl.FileSystemAccessRule($sid, $rightsEnum, $ciOi, $noPropagation, $allow)
        
        # "Dieser Ordner und Dateien" (keine Unterordner) - für "Nur Hauptordner"
        $mainOnlyRule = New-Object System.Security.AccessControl.FileSystemAccessRule($sid, $rightsEnum,
            [System.Security.AccessControl.InheritanceFlags]::ObjectInherit,
            [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit, $allow)
        
        $errors = New-Object System.Collections.Generic.List[string]
        
        switch ($ApplyTo) {
            "MainOnly" {
                $acl = Get-Acl -LiteralPath $FolderPath
                $acl.SetAccessRule($mainOnlyRule)
                Set-Acl -LiteralPath $FolderPath -AclObject $acl
            }
            "SubOnly" {
                # Regel auf jeden direkten Unterordner setzen - Vererbung erledigt die tieferen Ebenen
                $directChildren = @(([System.IO.DirectoryInfo]::new($FolderPath)).GetDirectories() |
                    Where-Object { ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0 })
                foreach ($child in $directChildren) {
                    try {
                        $acl = Get-Acl -LiteralPath $child.FullName
                        $acl.SetAccessRule($inheritableRule)
                        Set-Acl -LiteralPath $child.FullName -AclObject $acl
                    }
                    catch { $errors.Add("$($child.FullName) : $($_.Exception.Message)") }
                }
            }
            "Both" {
                # Eine vererbbare Regel am Hauptordner wirkt auf alle Unterordner und Dateien
                $acl = Get-Acl -LiteralPath $FolderPath
                $acl.SetAccessRule($inheritableRule)
                Set-Acl -LiteralPath $FolderPath -AclObject $acl
            }
        }
        
        # Unterordner mit deaktivierter Vererbung erhalten die Regel zusätzlich explizit
        if ($ApplyTo -eq "SubOnly" -or $ApplyTo -eq "Both") {
            foreach ($dir in (Get-SubDirectoryList -FolderPath $FolderPath)) {
                try {
                    $acl = Get-Acl -LiteralPath $dir
                    if ($acl.AreAccessRulesProtected) {
                        $acl.SetAccessRule($inheritableRule)
                        Set-Acl -LiteralPath $dir -AclObject $acl
                    }
                }
                catch { $errors.Add("$dir : $($_.Exception.Message)") }
            }
        }
        
        New-PermissionReport -Action "ADD" -Path $FolderPath -User $Username -Rights $Rights -ApplyTo $ApplyTo
        
        if ($errors.Count -gt 0) {
            foreach ($err in $errors) { Write-LogEntry "Hinzufügen: $err" "WARN" }
            $details = ($errors | Select-Object -First 20) -join "`n"
            [System.Windows.MessageBox]::Show("Berechtigung wurde gesetzt, aber $($errors.Count) Ordner konnten nicht geändert werden:`n`n$details", "Teilweise erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
        
        return $true
    }
    catch {
        Write-LogEntry "Fehler beim Hinzufügen der Berechtigung: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Hinzufügen der Berechtigung: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return $false
    }
}

function New-PermissionReport {
    param(
        [string]$Action,
        [string]$Path,
        [string]$User,
        [string]$Rights,
        [string]$ApplyTo
    )
    
    try {
        # Millisekunden im Dateinamen verhindern das Überschreiben bei mehreren Änderungen pro Sekunde
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss-fff"
        $reportFile = Join-Path $Global:ReportsPath "PermissionReport_$timestamp.txt"
        
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("=== FOLDER PERMISSIONS REPORT ===")
        [void]$sb.AppendLine("Datum/Zeit: $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")")
        [void]$sb.AppendLine("Ausgeführt von: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)")
        [void]$sb.AppendLine("Aktion: $Action")
        [void]$sb.AppendLine("Ordner: $Path")
        [void]$sb.AppendLine("Benutzer: $User")
        [void]$sb.AppendLine("Rechte: $Rights")
        [void]$sb.AppendLine("Angewendet auf: $ApplyTo")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("=== AKTUELLE BERECHTIGUNGEN NACH ÄNDERUNG ===")
        
        # Aktuelle Berechtigungen hinzufügen
        $currentPerms = Get-FolderPermissions -FolderPath $Path -IncludeSubfolders ($ApplyTo -eq "Both" -or $ApplyTo -eq "SubOnly")
        foreach ($perm in $currentPerms) {
            [void]$sb.AppendLine("$($perm.Folder) | $($perm.User) | $($perm.Rights) | $($perm.AccessType) | Vererbt: $($perm.Inherited)")
        }
        
        $sb.ToString() | Out-File -LiteralPath $reportFile -Encoding UTF8
    }
    catch {
        Write-LogEntry "Report konnte nicht erstellt werden: $($_.Exception.Message)" "ERROR"
    }
}

# WPF GUI XAML - Windows 11 Design
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Easy Folder Permissions Manager V0.0.2" Height="900" Width="1700"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        Background="#F3F3F3" FontFamily="Segoe UI" FontSize="14">
    <Grid Margin="15">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="2*"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Vereinfachte Styles -->
        <Grid.Resources>
            <Style TargetType="Button">
                <Setter Property="Background" Value="#0078D4"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
            </Style>
        </Grid.Resources>
        
        <!-- Ordner Auswahl -->
        <GroupBox Header="Ordner Auswahl" Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,0,0,10">
            <Grid Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Name="txtFolderPath" Grid.Column="0" Margin="0,0,10,0" Height="32" VerticalContentAlignment="Center"/>
                <Button Name="btnBrowse" Grid.Column="1" Content="Durchsuchen" Width="110" Height="32" Margin="0,0,8,0"/>
                <Button Name="btnLoadPermissions" Grid.Column="2" Content="Laden" Width="90" Height="32" Margin="0,0,8,0"/>
                <Button Name="btnBackup" Grid.Column="3" Content="Backup" Width="90" Height="32" Background="#107C10"
                        ToolTip="XML-Dokumentation + wiederherstellbares icacls-Backup (.acl) im Report-Ordner"/>
            </Grid>
        </GroupBox>
        
        <!-- Filter und Optionen -->
        <GroupBox Header="Filter und Optionen" Grid.Row="1" Grid.Column="0" Margin="0,0,10,10">
            <StackPanel Margin="10" Orientation="Vertical">
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                    <CheckBox Name="chkIncludeSubfolders" Content="Unterordner einbeziehen" Margin="0,0,25,0" VerticalAlignment="Center"/>
                    <CheckBox Name="chkShowInheritedOnly" Content="Nur vererbte Berechtigungen" VerticalAlignment="Center"/>
                </StackPanel>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Label Content="Benutzer:" Grid.Column="0" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBox Name="txtUserFilter" Grid.Column="1" Height="28" Margin="0,0,15,0" VerticalContentAlignment="Center"/>
                    <Label Content="Rechte:" Grid.Column="2" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBox Name="txtRightsFilter" Grid.Column="3" Height="28" Margin="0,0,15,0" VerticalContentAlignment="Center"/>
                    <Button Name="btnApplyFilter" Grid.Column="4" Content="Filter" Width="70" Height="28"/>
                </Grid>
            </StackPanel>
        </GroupBox>
        
        <!-- Berechtigungen Anzeige -->
        <GroupBox Header="Aktuelle Berechtigungen" Grid.Row="2" Grid.Column="0" Margin="0,0,10,10">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DataGrid Name="dgPermissions" Grid.Row="0" AutoGenerateColumns="True" IsReadOnly="True" EnableRowVirtualization="True" 
                         GridLinesVisibility="Horizontal" HeadersVisibility="All" MinHeight="300"/>
                <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
                    <TextBlock Name="txtPermissionCount" Text="Einträge: 0" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="#666"/>
                    <Button Name="btnRemoveSelected" Content="Entfernen" Width="120" Height="28" Background="#D13438" IsEnabled="False"
                            ToolTip="Entfernt exakt die markierten, expliziten Einträge"/>
                </StackPanel>
            </Grid>
        </GroupBox>
        
        <!-- Rechte Spalte: Benutzer verwalten und Status -->
        <StackPanel Grid.Row="1" Grid.RowSpan="2" Grid.Column="1" Margin="0,0,0,10">
            <!-- Benutzer hinzufügen -->
            <GroupBox Header="Benutzer verwalten" Margin="0,0,0,10">
                <StackPanel Margin="10">
                    <Grid Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="txtUsername" Grid.Column="0" Height="32" Margin="0,0,8,0" VerticalContentAlignment="Center" 
                                ToolTip="Windows-Benutzername oder Gruppe eingeben"/>
                        <Button Name="btnBrowseUser" Grid.Column="1" Content="..." Width="32" Height="32" ToolTip="Benutzer auswählen"/>
                    </Grid>
                    
                    <ComboBox Name="cmbRights" Height="32" Margin="0,0,0,8" VerticalContentAlignment="Center">
                        <ComboBoxItem Content="Vollzugriff" Tag="FullControl"/>
                        <ComboBoxItem Content="Ändern" Tag="Modify"/>
                        <ComboBoxItem Content="Lesen+Ausführen" Tag="ReadAndExecute"/>
                        <ComboBoxItem Content="Nur Lesen" Tag="Read"/>
                        <ComboBoxItem Content="Nur Schreiben" Tag="Write"/>
                    </ComboBox>
                    
                    <StackPanel Orientation="Vertical" Margin="0,0,0,10">
                        <RadioButton Name="rbMainOnly" Content="Nur Hauptordner (inkl. Dateien)" Margin="0,0,0,4" IsChecked="True"
                                     ToolTip="Dieser Ordner und die darin liegenden Dateien - keine Unterordner"/>
                        <RadioButton Name="rbSubOnly" Content="Nur Unterordner (inkl. Inhalt)" Margin="0,0,0,4"
                                     ToolTip="Alle Unterordner samt Dateien - nicht der Hauptordner selbst"/>
                        <RadioButton Name="rbBoth" Content="Hauptordner + Unterordner + Dateien" Margin="0,0,0,4"
                                     ToolTip="Vererbbare Berechtigung: dieser Ordner, Unterordner und Dateien"/>
                    </StackPanel>
                    
                    <Button Name="btnAddPermission" Content="Hinzufügen" Height="32" Margin="0,0,0,5"/>
                    <Button Name="btnRemovePermission" Content="Entfernen" Height="32" Background="#D13438"
                            ToolTip="Entfernt alle expliziten Einträge des Kontos im gewählten Bereich"/>
                </StackPanel>
            </GroupBox>
            
            <!-- Status und Export -->
            <GroupBox Header="Status und Export">
                <StackPanel Margin="10">
                    <TextBox Name="txtStatus" Height="80" IsReadOnly="True" 
                            VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" 
                            Background="#F9F9F9" BorderBrush="#E1E1E1" Padding="8" Margin="0,0,0,8"/>
                    <Button Name="btnOpenReports" Content="Reports" Height="32" Margin="0,0,0,5" Background="#107C10"/>
                    <Button Name="btnExportCsv" Content="CSV Export" Height="32" Background="#FF8C00"/>
                </StackPanel>
            </GroupBox>
        </StackPanel>
        
        
        <!-- Buttons -->
        <StackPanel Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnRefresh" Content="Aktualisieren" Width="110" Height="32" Margin="0,0,10,0" Background="#0078D4"/>
            <Button Name="btnExit" Content="Beenden" Width="90" Height="32" Background="#D13438"/>
        </StackPanel>
        
        <!-- Footer -->
        <Border Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Background="#E8E8E8" BorderBrush="#D0D0D0" BorderThickness="0,1,0,0" Margin="0,15,0,0">
            <Grid Margin="15,8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <!-- Links: Autor und Version -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Easy Folder Permissions Manager V0.0.2" FontWeight="SemiBold" Foreground="#333" Margin="0,0,15,0"/>
                    <TextBlock Text="© 2025 Andreas Hepp | PhinIT" Foreground="#666" FontSize="12"/>
                </StackPanel>
                
                <!-- Mitte: Website Link -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Website: " Foreground="#666" FontSize="12"/>
                    <TextBlock Name="txtWebsiteLink" Text="www.phinit.de" Foreground="#0078D4" FontSize="12" 
                              TextDecorations="Underline" Cursor="Hand" ToolTip="Website öffnen"/>
                </StackPanel>
                
                <!-- Rechts: System Info -->
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Right">
                    <TextBlock Name="txtSystemInfo" Text="Windows PowerShell" Foreground="#666" FontSize="12" Margin="0,0,10,0"/>
                    <TextBlock Name="txtCurrentTime" Text="" Foreground="#666" FontSize="12"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# WPF Window erstellen
try {
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $reader.Close()
}
catch {
    Write-LogEntry "Fehler beim Laden der GUI: $($_.Exception.Message)" "ERROR"
    Write-LogEntry "XAML-Inhalt zur Diagnose: $xaml" "ERROR"
    [System.Windows.MessageBox]::Show("Die Benutzeroberfläche konnte nicht geladen werden. Details siehe $Global:LogFilePath", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    exit 1
}

# Controls referenzieren und validieren
$controlNames = @(
    "txtFolderPath", "btnBrowse", "btnLoadPermissions", "btnBackup",
    "chkIncludeSubfolders", "chkShowInheritedOnly", "txtUserFilter", "txtRightsFilter", "btnApplyFilter",
    "dgPermissions", "txtPermissionCount", "btnRemoveSelected",
    "txtUsername", "btnBrowseUser", "cmbRights", "rbMainOnly", "rbSubOnly", "rbBoth",
    "btnAddPermission", "btnRemovePermission",
    "txtStatus", "btnOpenReports", "btnExportCsv", "btnRefresh", "btnExit",
    "txtWebsiteLink", "txtSystemInfo", "txtCurrentTime"
)

foreach ($controlName in $controlNames) {
    $control = $window.FindName($controlName)
    if ($null -eq $control) {
        Write-LogEntry "Control '$controlName' konnte nicht gefunden werden!" "ERROR"
        exit 1
    }
    Set-Variable -Name $controlName -Value $control -Scope Script
}

# ========== GUI-HILFSFUNKTIONEN ==========

# Verarbeitet ausstehende UI-Nachrichten, damit Statustexte/Cursor vor langen Operationen sichtbar werden
function Update-UI {
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Show-Info {
    param([string]$Message, [string]$Title = "Hinweis")
    [System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
}

# Gefilterte Daten im DataGrid anzeigen (zentral statt mehrfach dupliziert)
function Update-PermissionGrid {
    $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly ([bool]$chkShowInheritedOnly.IsChecked)
    $dgPermissions.ItemsSource = $filteredData
    $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
}

# Berechtigungen neu einlesen und anzeigen
function Invoke-PermissionReload {
    $Global:PermissionsData = Get-FolderPermissions -FolderPath $Global:CurrentFolder -IncludeSubfolders ([bool]$chkIncludeSubfolders.IsChecked)
    Update-PermissionGrid
}

function Get-SelectedApplyTo {
    if ($rbSubOnly.IsChecked) { return "SubOnly" }
    if ($rbBoth.IsChecked) { return "Both" }
    return "MainOnly"
}

function Get-ApplyToDisplayText {
    param([string]$ApplyTo)
    switch ($ApplyTo) {
        "SubOnly" { return "Nur Unterordner (inkl. deren Inhalt)" }
        "Both"    { return "Hauptordner, Unterordner und Dateien" }
        default   { return "Nur Hauptordner (inkl. Dateien darin)" }
    }
}

# Ordnerpfad normalisieren und prüfen
function Test-FolderInput {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return (Test-Path -LiteralPath $Path.Trim() -PathType Container)
}

# ========== EVENT HANDLERS ==========

$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Ordner für Berechtigungsverwaltung auswählen"
    
    # PS2EXE: ShowDialog() Rückgabewert in Variable speichern (verhindert MessageBox)
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolderPath.Text = $folderBrowser.SelectedPath
        $txtStatus.Text = "Ordner ausgewählt: $($folderBrowser.SelectedPath) - bitte 'Laden' klicken."
    }
    $folderBrowser.Dispose()
})

$btnLoadPermissions.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtFolderPath.Text)) {
        Show-Info "Bitte wählen Sie zuerst einen Ordner aus."
        return
    }
    
    if (-not (Test-FolderInput -Path $txtFolderPath.Text)) {
        [System.Windows.MessageBox]::Show("Der angegebene Ordner existiert nicht oder ist kein Ordner.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return
    }
    
    $Global:CurrentFolder = $txtFolderPath.Text.Trim()
    
    $txtStatus.Text = "Lade Berechtigungen..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    Update-UI
    
    try {
        Invoke-PermissionReload
        $txtStatus.Text = "Berechtigungen erfolgreich geladen: $Global:CurrentFolder"
    }
    catch {
        Write-LogEntry "Fehler beim Laden: $($_.Exception.Message)" "ERROR"
        $txtStatus.Text = "Fehler beim Laden der Berechtigungen: $($_.Exception.Message)"
    }
    finally {
        $window.Cursor = $null
    }
})

$btnAddPermission.Add_Click({
    if ([string]::IsNullOrEmpty($Global:CurrentFolder)) {
        Show-Info "Bitte wählen Sie zuerst einen Ordner aus und laden Sie die Berechtigungen."
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtUsername.Text)) {
        Show-Info "Bitte geben Sie einen Benutzernamen ein."
        return
    }
    
    if ($null -eq $cmbRights.SelectedItem) {
        Show-Info "Bitte wählen Sie die Rechte aus."
        return
    }
    
    $username = $txtUsername.Text.Trim()
    $rights = $cmbRights.SelectedItem.Tag
    $applyTo = Get-SelectedApplyTo
    
    # Sicherheitsabfrage vor der Änderung (Änderungen sind sofort wirksam)
    $confirm = [System.Windows.MessageBox]::Show(
        "Folgende Berechtigung wird gesetzt:`n`nOrdner: $Global:CurrentFolder`nBenutzer/Gruppe: $username`nRechte: $($cmbRights.SelectedItem.Content) ($rights)`nAnwenden auf: $(Get-ApplyToDisplayText $applyTo)`n`nBestehende explizite Allow-Einträge dieses Kontos werden dabei ersetzt. Fortfahren?",
        "Berechtigung hinzufügen", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
    
    $txtStatus.Text = "Füge Berechtigung hinzu..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    Update-UI
    
    try {
        $success = Add-FolderPermission -FolderPath $Global:CurrentFolder -Username $username -Rights $rights -ApplyTo $applyTo
        
        if ($success) {
            $txtStatus.Text = "Berechtigung erfolgreich hinzugefügt für: $username ($rights) - $(Get-ApplyToDisplayText $applyTo)"
            $txtUsername.Text = ""
            Invoke-PermissionReload
        } else {
            $txtStatus.Text = "Berechtigung konnte nicht hinzugefügt werden."
        }
    }
    finally {
        $window.Cursor = $null
    }
})

$btnOpenReports.Add_Click({
    if (Test-Path -LiteralPath $Global:ReportsPath) {
        # Pfad in Anführungszeichen, sonst bricht explorer.exe bei Leerzeichen im Benutzerpfad ab
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$Global:ReportsPath`""
    } else {
        Show-Info "Reports-Ordner wurde noch nicht erstellt."
    }
})

$btnRefresh.Add_Click({
    if (-not [string]::IsNullOrEmpty($Global:CurrentFolder)) {
        $txtStatus.Text = "Aktualisiere Berechtigungen..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Update-UI
        
        try {
            Invoke-PermissionReload
            $txtStatus.Text = "Berechtigungen erfolgreich aktualisiert"
        }
        finally {
            $window.Cursor = $null
        }
    }
})

$btnExit.Add_Click({
    $window.Close()
})

$btnBackup.Add_Click({
    if ([string]::IsNullOrEmpty($Global:CurrentFolder)) {
        Show-Info "Bitte wählen Sie zuerst einen Ordner aus und laden Sie die Berechtigungen."
        return
    }
    
    $txtStatus.Text = "Erstelle Backup..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    Update-UI
    
    try {
        $backupFile = Backup-FolderPermissions -FolderPath $Global:CurrentFolder
        if ($backupFile) {
            $txtStatus.Text = "Backup erstellt: $(Split-Path $backupFile -Leaf) (+ .acl für icacls /restore)"
        }
    }
    finally {
        $window.Cursor = $null
    }
})

$btnApplyFilter.Add_Click({
    if ($Global:PermissionsData.Count -gt 0) {
        Update-PermissionGrid
        $txtStatus.Text = "Filter angewendet"
    }
})

$btnRemovePermission.Add_Click({
    if ([string]::IsNullOrEmpty($Global:CurrentFolder)) {
        Show-Info "Bitte wählen Sie zuerst einen Ordner aus und laden Sie die Berechtigungen."
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtUsername.Text)) {
        Show-Info "Bitte geben Sie einen Benutzernamen ein."
        return
    }
    
    $username = $txtUsername.Text.Trim()
    $applyTo = Get-SelectedApplyTo
    
    $result = [System.Windows.MessageBox]::Show("Möchten Sie alle expliziten Berechtigungen für '$username' wirklich entfernen?`n`nOrdner: $Global:CurrentFolder`nBereich: $(Get-ApplyToDisplayText $applyTo)`n`nHinweis: Vererbte Einträge bleiben bestehen und müssen am Ursprungsordner entfernt werden.", "Bestätigung", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $txtStatus.Text = "Entferne Berechtigung..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Update-UI
        
        try {
            $success = Remove-FolderPermission -FolderPath $Global:CurrentFolder -Username $username -ApplyTo $applyTo
            
            if ($success) {
                $txtStatus.Text = "Berechtigung erfolgreich entfernt für: $username - $(Get-ApplyToDisplayText $applyTo)"
                $txtUsername.Text = ""
                Invoke-PermissionReload
            }
        }
        finally {
            $window.Cursor = $null
        }
    }
})

$btnExportCsv.Add_Click({
    if ($Global:PermissionsData.Count -eq 0) {
        Show-Info "Keine Daten zum Exportieren vorhanden."
        return
    }
    
    # Erweiterte Export-Optionen
    $exportChoice = [System.Windows.MessageBox]::Show("Welches Format möchten Sie verwenden?`n`nJa = CSV-Datei`nNein = HTML-Report`nAbbrechen = Vorgang abbrechen", "Export-Format wählen", [System.Windows.MessageBoxButton]::YesNoCancel, [System.Windows.MessageBoxImage]::Question)
    
    if ($exportChoice -eq [System.Windows.MessageBoxResult]::Cancel) {
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    
    if ($exportChoice -eq [System.Windows.MessageBoxResult]::Yes) {
        # CSV Export (Semikolon für deutsches Excel)
        $saveDialog.Filter = "CSV Dateien (*.csv)|*.csv"
        $saveDialog.FileName = "FolderPermissions_$timestamp.csv"
        
        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $Global:PermissionsData |
                    Select-Object Folder, User, @{Name = "Rights"; Expression = { "$($_.Rights)" }}, @{Name = "AccessType"; Expression = { "$($_.AccessType)" }}, Inherited,
                                  @{Name = "InheritanceFlags"; Expression = { "$($_.InheritanceFlags)" }}, @{Name = "PropagationFlags"; Expression = { "$($_.PropagationFlags)" }} |
                    Export-Csv -LiteralPath $saveDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                $txtStatus.Text = "CSV erfolgreich exportiert: $(Split-Path $saveDialog.FileName -Leaf)"
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler beim CSV-Export: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            }
        }
    } else {
        # HTML Export
        $saveDialog.Filter = "HTML Dateien (*.html)|*.html"
        $saveDialog.FileName = "FolderPermissions_Report_$timestamp.html"
        
        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                Export-PermissionsToHtml -Permissions $Global:PermissionsData -FilePath $saveDialog.FileName
                $txtStatus.Text = "HTML-Report erfolgreich exportiert: $(Split-Path $saveDialog.FileName -Leaf)"
                
                # Optional: HTML-Datei öffnen
                $openResult = [System.Windows.MessageBox]::Show("HTML-Report wurde erstellt. Möchten Sie ihn jetzt öffnen?", "Report erstellt", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                if ($openResult -eq [System.Windows.MessageBoxResult]::Yes) {
                    Invoke-Item -LiteralPath $saveDialog.FileName
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler beim HTML-Export: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            }
        }
    }
    $saveDialog.Dispose()
})

$btnBrowseUser.Add_Click({
    try {
        $txtStatus.Text = "Lade Benutzer und Gruppen..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Update-UI
        
        $users = Get-SystemUsers
        $window.Cursor = $null
        
        if ($users.Count -eq 0) {
            Show-Info "Keine Benutzer oder Gruppen gefunden." "Information"
            return
        }
        
        $selectedUser = Show-PrincipalPicker -Principals $users -Owner $window
        
        if ($selectedUser) {
            $txtUsername.Text = $selectedUser.Name
            $txtStatus.Text = "Benutzer ausgewählt: $($selectedUser.FullName)"
        } else {
            $txtStatus.Text = "Keine Auswahl getroffen"
        }
    }
    catch {
        Write-LogEntry "Fehler beim Laden der Benutzer: $($_.Exception.Message)" "ERROR"
        $txtStatus.Text = "Fehler beim Laden der Benutzer"
    }
    finally {
        $window.Cursor = $null
    }
})

# DataGrid Selection Event
$dgPermissions.Add_SelectionChanged({
    $btnRemoveSelected.IsEnabled = $dgPermissions.SelectedItems.Count -gt 0
})

# Ausgewählte Berechtigungen entfernen
$btnRemoveSelected.Add_Click({
    if ($dgPermissions.SelectedItems.Count -eq 0) {
        Show-Info "Bitte wählen Sie mindestens eine Berechtigung aus."
        return
    }
    
    $selectedCount = $dgPermissions.SelectedItems.Count
    $result = [System.Windows.MessageBox]::Show("Möchten Sie $selectedCount ausgewählte Berechtigung(en) wirklich entfernen?`n`nEs werden nur die exakt ausgewählten, expliziten Einträge entfernt. Vererbte Einträge werden übersprungen.", "Bestätigung", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $txtStatus.Text = "Entferne $selectedCount Berechtigung(en)..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Update-UI
        
        try {
            $selectedPermissions = @($dgPermissions.SelectedItems)
            
            $removeResult = Remove-SelectedPermissions -SelectedPermissions $selectedPermissions -FolderPath $Global:CurrentFolder
            
            $txtStatus.Text = "$($removeResult.Success) Berechtigung(en) entfernt, $($removeResult.Errors) Fehler/übersprungen"
            if ($removeResult.Errors -gt 0) {
                $errorDetails = ($removeResult.ErrorDetails | Select-Object -First 20) -join "`n"
                [System.Windows.MessageBox]::Show("Einige Berechtigungen konnten nicht entfernt werden:`n`n$errorDetails", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            }
            
            if ($removeResult.Success -gt 0) {
                Invoke-PermissionReload
            }
        }
        catch {
            Write-LogEntry "Fehler beim Entfernen der Berechtigungen: $($_.Exception.Message)" "ERROR"
            $txtStatus.Text = "Fehler beim Entfernen der Berechtigungen"
        }
        finally {
            $window.Cursor = $null
        }
    }
})

# Filter bei Eingabe automatisch anwenden
$applyFilterHandler = {
    if ($Global:PermissionsData.Count -gt 0) {
        Update-PermissionGrid
    }
}
$txtUserFilter.Add_TextChanged($applyFilterHandler)
$txtRightsFilter.Add_TextChanged($applyFilterHandler)
$chkShowInheritedOnly.Add_Checked($applyFilterHandler)
$chkShowInheritedOnly.Add_Unchecked($applyFilterHandler)

# Footer-Funktionalität
$txtWebsiteLink.Add_MouseLeftButtonUp({
    try {
        Start-Process "https://www.phinit.de"
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Öffnen der Website: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
    }
})

# System-Info und Zeit aktualisieren
function Update-FooterInfo {
    try {
        # System-Info
        $psVersion = $PSVersionTable.PSVersion.ToString()
        $osInfo = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption
        $txtSystemInfo.Text = "PowerShell $psVersion | $osInfo"
    }
    catch {
        $txtSystemInfo.Text = "PowerShell $($PSVersionTable.PSVersion)"
    }
    $txtCurrentTime.Text = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
}

# Timer für Zeit-Update
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({
    $txtCurrentTime.Text = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
})
$timer.Start()

# Timer beim Schließen stoppen (keine Ticks mehr nach dem Fenster-Ende)
$window.Add_Closed({
    $timer.Stop()
})

# Standardwerte setzen
$cmbRights.SelectedIndex = 0
$txtStatus.Text = "Bereit. Wählen Sie einen Ordner aus und laden Sie die Berechtigungen."
$txtPermissionCount.Text = "Einträge: 0"

# Footer initialisieren
Update-FooterInfo

Write-LogEntry "easyFPManager V$Global:AppVersion gestartet von $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" "INFO"

# Window anzeigen (VOID verhindert "False" MessageBox bei -noConsole)
[VOID]$window.ShowDialog()
