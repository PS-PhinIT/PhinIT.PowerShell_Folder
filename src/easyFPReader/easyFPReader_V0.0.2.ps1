#Requires -Version 5.1

# PS2EXE OPTIMIERUNGEN - KRITISCH FÜR GUI-EXE
$ProgressPreference = 'SilentlyContinue'  # Verhindert blinkende Progress-Fenster bei -noConsole

<#
.SYNOPSIS
    easyFolderPermissions Reader V0.0.1 - Ordnerberechtigungen-Analyse mit WPF GUI
    
.DESCRIPTION
    Dieses Script bietet eine grafische Benutzeroberfläche zur Analyse von Ordnerberechtigungen
    mit Export-Funktionen für HTML und SharePoint-Integration.
    
.FEATURES
    - WPF GUI für einfache Bedienung
    - Ordner-Auswahl mit Baumstruktur-Anzeige
    - Rekursive Berechtigungsanalyse
    - HTML-Export der Berechtigungen
    - SharePoint-Export mit OnPrem-Benutzer-Mapping
    - CSV und HTML Export für SharePoint-Daten
    
.AUTHOR
    PhinIT Solutions
    
.VERSION
    0.0.1
    
.DATE
    2024-10-24
    
.NOTES
    WICHTIG FÜR EXE-KONVERTIERUNG MIT PS2EXE:
    - $ProgressPreference = 'SilentlyContinue' am Anfang (verhindert Progress-Fenster)
    - Visual Styles VOR GUI-Erstellung aktiviert
    - ShowDialog() mit [VOID] versehen (verhindert "False" MessageBox)
    - UTF8-Encoding für deutsche Umlaute
    
    KOMPILIERUNG:
    ps2exe .\easyFPReader_V0.0.1.ps1 .\easyFPReader.exe -noConsole -STA -x64 `
           -iconFile .\icon.ico -title "Easy Folder Permissions Reader" `
           -version "0.0.1.0" -company "PhinIT" -copyright "© 2025 Andreas Hepp"
#>

# PS2EXE: Visual Styles MÜSSEN VOR allen GUI-Objekten aktiviert werden!
[System.Windows.Forms.Application]::EnableVisualStyles()

# Assembly-Imports für WPF (Out-Null verhindert MessageBoxen bei -noConsole)
Add-Type -AssemblyName PresentationFramework | Out-Null
Add-Type -AssemblyName PresentationCore | Out-Null
Add-Type -AssemblyName WindowsBase | Out-Null
Add-Type -AssemblyName System.Windows.Forms | Out-Null

# ========== LOGGING-SYSTEM (Optional - deaktivieren Sie $enableLogging für keine Logs) ==========
$enableLogging = $false  # Setzen Sie auf $true um Logfile zu erstellen
$logPath = ""

function Initialize-Logging {
    if ($enableLogging) {
        $global:logPath = Join-Path $env:TEMP "easyFolderPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        "Logging gestartet: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $global:logPath -Encoding UTF8
    }
}

function Write-Log {
    param([string]$Message)
    if ($enableLogging -and $global:logPath) {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message" | Out-File -FilePath $global:logPath -Encoding UTF8 -Append
    }
}

Initialize-Logging

# Globale Variablen
$Global:SelectedFolders = @()
$Global:PermissionsData = @()
$Global:SharePointData = @()
$Global:UserMappings = @()
$Global:CustomSharePointData = @()

# XAML für die WPF GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="easyFolderPermissions Reader V0.0.1" 
        Height="800" Width="1200"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2C3E50" CornerRadius="5" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="easyFolderPermissions Reader V0.0.1" 
                          FontSize="20" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center"/>
                <TextBlock Text="Ordnerberechtigungen analysieren und exportieren" 
                          FontSize="12" Foreground="#BDC3C7" HorizontalAlignment="Center" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Name="MainTabControl">
            
            <!-- Tab 1: Ordner-Berechtigungen -->
            <TabItem Header="📁 Ordner-Berechtigungen" FontSize="14">
                <Grid Margin="10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="300"/>
                        <ColumnDefinition Width="5"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Linke Spalte: Ordner-Auswahl -->
                    <Border Grid.Column="0" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5">
                        <StackPanel Margin="10">
                            <TextBlock Text="Ordner-Auswahl" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                            
                            <Button Name="BtnSelectFolder" Content="📂 Ordner hinzufügen" 
                                   Height="35" Margin="0,0,0,10" Background="#3498DB" Foreground="White" 
                                   BorderThickness="0" FontWeight="Bold"/>
                            
                            <TextBlock Text="Ausgewählte Ordner:" FontWeight="Bold" Margin="0,10,0,5"/>
                            <ListBox Name="LstSelectedFolders" Height="200" Margin="0,0,0,10"/>
                            
                            <Button Name="BtnRemoveFolder" Content="❌ Entfernen" 
                                   Height="30" Margin="0,0,0,10" Background="#E74C3C" Foreground="White" 
                                   BorderThickness="0"/>
                            
                            <Separator Margin="0,10"/>
                            
                            <Button Name="BtnAnalyzePermissions" Content="🔍 Berechtigungen analysieren" 
                                   Height="40" Margin="0,10,0,10" Background="#27AE60" Foreground="White" 
                                   BorderThickness="0" FontWeight="Bold" FontSize="12"/>
                            
                            <Button Name="BtnExportHTML" Content="📄 HTML Export" 
                                   Height="35" Margin="0,0,0,5" Background="#F39C12" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                        </StackPanel>
                    </Border>
                    
                    <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="#BDC3C7"/>
                    
                    <!-- Rechte Spalte: Ergebnisse -->
                    <Border Grid.Column="2" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5">
                        <Grid Margin="10">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Berechtigungen-Übersicht" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                            
                            <TreeView Grid.Row="1" Name="TreePermissions" FontFamily="Consolas" FontSize="11"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
            <!-- Tab 3: Benutzer-Anpassung -->
            <TabItem Header="👤 SharePoint Benutzer-Mapping" FontSize="14">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Anpassungs-Konfiguration -->
                    <Border Grid.Row="0" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,10">
                        <StackPanel>
                            <TextBlock Text="Manuelle UPN-Anpassung" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                            
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                
                                <TextBlock Grid.Column="0" Text="OnPrem:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <TextBox Grid.Column="1" Name="TxtOnPremUser" Height="25" Margin="0,0,10,0"/>
                                
                                <TextBlock Grid.Column="2" Text="→ EntraID UPN:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <TextBox Grid.Column="3" Name="TxtEntraIDUPN" Height="25" Margin="0,0,10,0"/>
                                
                                <Button Grid.Column="4" Name="BtnAddMapping" Content="➕ Hinzufügen" 
                                       Height="25" Width="100" Background="#27AE60" Foreground="White" BorderThickness="0"/>
                            </Grid>
                            
                            <TextBlock Text="Hier können Sie falsche automatische Zuordnungen manuell korrigieren oder neue hinzufügen." 
                                      FontSize="10" Foreground="Gray" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Benutzer-Mapping Tabelle -->
                    <Border Grid.Row="1" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBlock Text="Manuelle Benutzer-Zuordnungen" FontWeight="Bold" FontSize="14" VerticalAlignment="Center"/>
                                <Button Name="BtnLoadFromMapping" Content="🔄 UPN-Mapping aktualisieren" 
                                       Height="25" Width="150" Margin="20,0,10,0" Background="#3498DB" Foreground="White" BorderThickness="0"/>
                                <Button Name="BtnClearMappings" Content="🗑️ Alle löschen" 
                                       Height="25" Width="100" Background="#E74C3C" Foreground="White" BorderThickness="0"/>
                            </StackPanel>
                            
                            <DataGrid Grid.Row="1" Name="DgUserMappings" AutoGenerateColumns="False" 
                                     CanUserAddRows="False" CanUserDeleteRows="True" GridLinesVisibility="Horizontal"
                                     HeadersVisibility="Column" AlternatingRowBackground="#F8F9FA">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="OnPrem Benutzer" Binding="{Binding OnPremUser}" Width="250" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="EntraID UPN" Binding="{Binding EntraIDUPN}" Width="300"/>
                                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Quelle" Binding="{Binding Source}" Width="100" IsReadOnly="True"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                    
                    <!-- Export und Aktionen -->
                    <Border Grid.Row="2" Padding="15" Margin="0,10,0,0">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Name="BtnApplyMappings" Content="✅ Zuordnungen anwenden" 
                                   Height="35" Width="180" Margin="0,0,10,0" Background="#8E44AD" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                            <Button Name="BtnExportCustomCSV" Content="📊 Angepasste CSV exportieren" 
                                   Height="35" Width="200" Background="#16A085" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Tab 2: SharePoint Export -->
            <TabItem Header="☁️ SharePoint Permissions-Mapping" FontSize="14">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- SharePoint Konfiguration -->
                    <Border Grid.Row="0" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,10">
                        <StackPanel>
                            <TextBlock Text="SharePoint / EntraID UPN-Mapping" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                            
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <!-- Tenant Domain -->
                                <Grid Grid.Row="0" Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <TextBlock Grid.Column="0" Text="Tenant Domain:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <TextBox Grid.Column="1" Name="TxtTenantDomain" Height="25" 
                                            Text="contoso.onmicrosoft.com" Margin="0,0,10,0"/>
                                    <Button Grid.Column="2" Name="BtnCreateMapping" Content="🔄 UPN Mapping erstellen" 
                                           Height="25" Width="150" Background="#8E44AD" Foreground="White" BorderThickness="0"/>
                                </Grid>
                                
                                <!-- SharePoint Site URL -->
                                <Grid Grid.Row="1" Margin="0,0,0,5">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <TextBlock Grid.Column="0" Text="SharePoint Site URL:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <TextBox Grid.Column="1" Name="TxtSharePointSiteURL" Height="25" 
                                            Text="https://contoso.sharepoint.com/sites/Documents/Shared Documents"/>
                                </Grid>
                            </Grid>
                            
                            <TextBlock Text="Beispiele: Tenant Domain: contoso.onmicrosoft.com | Site URL: https://contoso.sharepoint.com/sites/Documents/Shared Documents (Ordner werden direkt hier erstellt)" 
                                      FontSize="10" Foreground="Gray" Margin="0,5,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- SharePoint Benutzer-Mapping -->
                    <Border Grid.Row="1" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="OnPrem zu EntraID UPN-Mapping" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                            
                            <DataGrid Grid.Row="1" Name="DgSharePointMapping" AutoGenerateColumns="False" 
                                     CanUserAddRows="False" CanUserDeleteRows="True" GridLinesVisibility="Horizontal"
                                     HeadersVisibility="Column" AlternatingRowBackground="#F8F9FA">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="OnPrem Benutzer" Binding="{Binding OnPremUser}" Width="200"/>
                                    <DataGridTextColumn Header="EntraID UPN" Binding="{Binding SharePointUPN}" Width="250"/>
                                    <DataGridTextColumn Header="Berechtigung" Binding="{Binding Permission}" Width="150"/>
                                    <DataGridTextColumn Header="SharePoint Pfad" Binding="{Binding SharePointPath}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                    
                    <!-- Export Buttons -->
                    <Border Grid.Row="2" Padding="15" Margin="0,10,0,0">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Name="BtnExportUPNCSV" Content="📊 UPN-Mapping CSV Export" 
                                   Height="35" Width="200" Background="#16A085" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>            
        </TabControl>
        
        <!-- Status Bar -->
        <Border Grid.Row="2" Background="#34495E" CornerRadius="3" Padding="10" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Name="TxtStatus" Text="Bereit..." Foreground="White" VerticalAlignment="Center"/>
                <ProgressBar Grid.Column="1" Name="ProgressBar" Width="200" Height="15" Visibility="Collapsed"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# TreeView wird jetzt direkt mit System.Windows.Controls.TreeViewItem befüllt

# Klasse für SharePoint-Mapping
class SharePointMapping {
    [string]$OnPremUser
    [string]$SharePointUPN
    [string]$Permission
    [string]$FolderPath
    [string]$SharePointPath
    
    SharePointMapping([string]$onprem, [string]$sharepoint, [string]$permission, [string]$folder, [string]$spPath) {
        $this.OnPremUser = $onprem
        $this.SharePointUPN = $sharepoint
        $this.Permission = $permission
        $this.FolderPath = $folder
        $this.SharePointPath = $spPath
    }
}

# Klasse für manuelle Benutzer-Mappings
class UserMapping {
    [string]$OnPremUser
    [string]$EntraIDUPN
    [string]$Status
    [string]$Source
    
    UserMapping([string]$onprem, [string]$entraid, [string]$status, [string]$source) {
        $this.OnPremUser = $onprem
        $this.EntraIDUPN = $entraid
        $this.Status = $status
        $this.Source = $source
    }
}

# Funktion: Ordner-Berechtigungen analysieren
function Get-FolderPermissions {
    param(
        [string]$FolderPath,
        [bool]$Recursive = $true
    )
    
    $results = @()
    
    try {
        # Hauptordner analysieren
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        
        $folderInfo = [PSCustomObject]@{
            Path = $FolderPath
            Type = "Folder"
            Permissions = @()
        }
        
        foreach ($access in $acl.Access) {
            $permission = [PSCustomObject]@{
                Identity = $access.IdentityReference.Value
                Rights = $access.FileSystemRights
                AccessType = $access.AccessControlType
                Inherited = $access.IsInherited
            }
            $folderInfo.Permissions += $permission
        }
        
        $results += $folderInfo
        
        # Unterordner rekursiv analysieren
        if ($Recursive) {
            $subFolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
            foreach ($subFolder in $subFolders) {
                $subResults = Get-FolderPermissions -FolderPath $subFolder.FullName -Recursive $true
                $results += $subResults
            }
        }
    }
    catch {
        Write-Log "Fehler beim Analysieren von $FolderPath : $($_.Exception.Message)"
    }
    
    return $results
}

# Funktion: TreeView mit Berechtigungen füllen
function Update-PermissionsTreeView {
    param($TreeView, $PermissionsData)
    
    $TreeView.Items.Clear()
    
    foreach ($folder in $PermissionsData) {
        # Hauptordner als TreeViewItem erstellen
        $folderItem = New-Object System.Windows.Controls.TreeViewItem
        $folderItem.Header = "📁 $($folder.Path)"
        $folderItem.Foreground = "Blue"
        $folderItem.FontWeight = "Bold"
        
        foreach ($permission in $folder.Permissions) {
            $permText = "$($permission.Identity) - $($permission.Rights) ($($permission.AccessType))"
            if ($permission.Inherited) { $permText += " [Inherited]" }
            
            $permItem = New-Object System.Windows.Controls.TreeViewItem
            $permItem.Header = $permText
            
            # Farbe basierend auf AccessType setzen
            switch ($permission.AccessType) {
                "Allow" { $permItem.Foreground = "Green" }
                "Deny" { $permItem.Foreground = "Red" }
                default { $permItem.Foreground = "Black" }
            }
            
            $folderItem.Items.Add($permItem)
        }
        
        $TreeView.Items.Add($folderItem)
    }
}

# Funktion: HTML-Export erstellen
function Export-PermissionsToHTML {
    param(
        [array]$PermissionsData,
        [string]$OutputPath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Ordnerberechtigungen Report</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #2C3E50; color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        
        /* Baumstruktur Styles */
        .tree { margin: 20px 0; }
        .tree-item { margin: 2px 0; }
        .tree-folder { 
            cursor: pointer; 
            padding: 8px; 
            background: #f8f9fa; 
            border-left: 4px solid #3498DB; 
            margin: 5px 0;
            border-radius: 3px;
        }
        .tree-folder:hover { background: #e9ecef; }
        .tree-folder.root { background: #3498DB; color: white; font-weight: bold; }
        .tree-folder.subfolder { margin-left: 20px; background: #e3f2fd; }
        .tree-folder.deep { margin-left: 40px; background: #f3e5f5; }
        
        .permissions { 
            padding: 15px; 
            margin-left: 20px; 
            border-left: 2px solid #ddd; 
            background: white;
            display: none;
        }
        .permissions.show { display: block; }
        
        .permission { 
            margin: 5px 0; 
            padding: 8px; 
            background: #f9f9f9; 
            border-left: 4px solid #3498DB; 
            border-radius: 3px;
        }
        .allow { border-left-color: #27AE60; }
        .deny { border-left-color: #E74C3C; }
        .inherited { opacity: 0.7; font-style: italic; }
        .timestamp { text-align: right; color: #7F8C8D; font-size: 12px; }
        
        .toggle-icon { 
            display: inline-block; 
            width: 16px; 
            margin-right: 8px; 
            font-weight: bold;
        }
        .folder-path { font-family: 'Courier New', monospace; font-size: 12px; }
    </style>
    <script>
        function toggleFolder(element) {
            const permissions = element.nextElementSibling;
            const icon = element.querySelector('.toggle-icon');
            
            if (permissions && permissions.classList.contains('permissions')) {
                permissions.classList.toggle('show');
                icon.textContent = permissions.classList.contains('show') ? '📂' : '📁';
            }
        }
        
        function expandAll() {
            document.querySelectorAll('.permissions').forEach(p => p.classList.add('show'));
            document.querySelectorAll('.toggle-icon').forEach(i => i.textContent = '📂');
        }
        
        function collapseAll() {
            document.querySelectorAll('.permissions').forEach(p => p.classList.remove('show'));
            document.querySelectorAll('.toggle-icon').forEach(i => i.textContent = '📁');
        }
    </script>
</head>
<body>
    <div class="header">
        <h1>📁 Ordnerberechtigungen Report</h1>
        <p>Detaillierte Analyse der Dateisystem-Berechtigungen mit Baumstruktur</p>
        <div style="margin: 10px 0;">
            <button onclick="expandAll()" style="padding: 8px 16px; margin-right: 10px; background: #27AE60; color: white; border: none; border-radius: 3px; cursor: pointer;">📂 Alle öffnen</button>
            <button onclick="collapseAll()" style="padding: 8px 16px; background: #E74C3C; color: white; border: none; border-radius: 3px; cursor: pointer;">📁 Alle schließen</button>
        </div>
        <div class="timestamp">Erstellt am: $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")</div>
    </div>
    
    <div class="tree">
"@
    
    # Ordner nach Pfadtiefe sortieren für Baumstruktur
    $sortedFolders = $PermissionsData | Sort-Object { $_.Path.Split('\').Count }, Path
    
    foreach ($folder in $sortedFolders) {
        $pathDepth = $folder.Path.Split('\').Count
        $folderName = Split-Path $folder.Path -Leaf
        $parentPath = Split-Path $folder.Path -Parent
        
        # CSS-Klasse basierend auf Tiefe
        $cssClass = "tree-folder"
        if ($pathDepth -le 3) { $cssClass += " root" }
        elseif ($pathDepth -le 5) { $cssClass += " subfolder" }
        else { $cssClass += " deep" }
        
        $html += @"
        <div class="tree-item">
            <div class="$cssClass" onclick="toggleFolder(this)">
                <span class="toggle-icon">📁</span>
                <strong>$folderName</strong>
                <div class="folder-path">$($folder.Path)</div>
            </div>
            <div class="permissions">
"@
        
        foreach ($permission in $folder.Permissions) {
            $permissionClass = "permission"
            if ($permission.AccessType -eq "Allow") { $permissionClass += " allow" }
            if ($permission.AccessType -eq "Deny") { $permissionClass += " deny" }
            if ($permission.Inherited) { $permissionClass += " inherited" }
            
            $inheritedText = if ($permission.Inherited) { " [Inherited]" } else { "" }
            
            $html += @"
                <div class="$permissionClass">
                    <strong>$($permission.Identity)</strong> - $($permission.Rights) ($($permission.AccessType))$inheritedText
                </div>
"@
        }
        
        $html += @"
            </div>
        </div>
"@
    }
    
    $html += @"
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
}

# Funktion: OnPrem zu EntraID UPN konvertieren
function Convert-OnPremToEntraID {
    param(
        [string]$OnPremUser,
        [string]$TenantDomain
    )
    
    # Konvertierung von verschiedenen OnPrem-Formaten zu EntraID UPN
    if ($OnPremUser -match "^(.+)\\(.+)$") {
        # Format: DOMAIN\username -> username@tenant.domain
        $username = $Matches[2]
        return "$username@$TenantDomain"
    }
    elseif ($OnPremUser -match "^(.+)@(.+)$") {
        # Bereits UPN-Format - prüfen ob Tenant-Domain angepasst werden muss
        $username = $Matches[1]
        return "$username@$TenantDomain"
    }
    else {
        # Nur Username -> username@tenant.domain
        return "$OnPremUser@$TenantDomain"
    }
}

# Funktion: Lokalen Ordnerpfad zu SharePoint-Pfad konvertieren
function Convert-LocalPathToSharePoint {
    param(
        [string]$LocalPath,
        [string]$SharePointSiteURL,
        [array]$AllFolderPaths
    )
    
    try {
        # SharePoint Site URL normalisieren
        $normalizedSiteURL = $SharePointSiteURL.TrimEnd('/')
        
        # Alle Pfade normalisieren für Vergleich
        $normalizedLocalPath = $LocalPath.Replace('\', '/').TrimEnd('/')
        $normalizedAllPaths = $AllFolderPaths | ForEach-Object { $_.Replace('\', '/').TrimEnd('/') }
        
        # Den ursprünglich ausgewählten Ordner (Root) finden - der kürzeste Pfad
        $rootPath = ($normalizedAllPaths | Sort-Object Length)[0]
        
        if ($normalizedLocalPath -eq $rootPath) {
            # Das ist der ausgewählte Hauptordner - sein INHALT kommt direkt in die SharePoint-Site
            return $normalizedSiteURL
        }
        else {
            # Unterordner - vollständige Hierarchie beibehalten
            $relativePath = $normalizedLocalPath.Substring($rootPath.Length).TrimStart('/')
            return "$normalizedSiteURL/$relativePath"
        }
    }
    catch {
        # Fallback: Direkt die angegebene URL verwenden
        return $SharePointSiteURL.TrimEnd('/')
    }
}

# Funktion: UPN-Mapping erstellen
function New-UPNMapping {
    param(
        [array]$PermissionsData,
        [string]$TenantDomain,
        [string]$SharePointSiteURL = ""
    )
    
    $mappings = @()
    
    # Alle Ordnerpfade für intelligente SharePoint-Pfad-Erstellung sammeln
    $allFolderPaths = $PermissionsData | ForEach-Object { $_.Path }
    
    foreach ($folder in $PermissionsData) {
        Write-Log "DEBUG: Verarbeite Ordner: $($folder.Path) mit $($folder.Permissions.Count) Berechtigungen"
        
        foreach ($permission in $folder.Permissions) {
            Write-Log "DEBUG: Berechtigung - Identity: '$($permission.Identity)', AccessType: '$($permission.AccessType)', Rights: '$($permission.Rights)'"
            
            # Erweiterte Benutzer-Filterung mit Debug-Ausgabe - Lokale/System-Accounts ausschließen
            $systemAccountPatterns = @(
                "^NT AUTHORITY",           # NT-Autorität (englisch)
                "^NT-AUTORITÄT",          # NT-Autorität (deutsch)
                "^BUILTIN",               # Vordefiniert (englisch)
                "^VORDEFINIERT",          # Vordefiniert (deutsch)
                "^Everyone$",             # Jeder
                "^Jeder$",                # Jeder (deutsch)
                "^Authenticated Users",    # Authentifizierte Benutzer
                "^Authentifizierte Benutzer", # Authentifizierte Benutzer (deutsch)
                "^SYSTEM$",               # System
                "^LOCAL SERVICE",         # Lokaler Dienst
                "^LOKALER DIENST",        # Lokaler Dienst (deutsch)
                "^NETWORK SERVICE",       # Netzwerkdienst
                "^NETZWERKDIENST",       # Netzwerkdienst (deutsch)
                "^ANONYMOUS LOGON",       # Anonyme Anmeldung
                "^ANONYME ANMELDUNG",     # Anonyme Anmeldung (deutsch)
                "^S-1-",                  # SID-basierte Accounts
                "^IIS_IUSRS",            # IIS Benutzer
                "^IUSR",                 # Internet Benutzer
                "^\w+\$",                # Computer-Accounts (enden mit $)
                "^Creator Owner",         # Ersteller-Besitzer
                "^Ersteller-Besitzer"    # Ersteller-Besitzer (deutsch)
            )
            
            $isSystemAccount = $false
            foreach ($pattern in $systemAccountPatterns) {
                if ($permission.Identity -match $pattern) {
                    $isSystemAccount = $true
                    break
                }
            }
            
            $isAllowAccess = $permission.AccessType -eq "Allow"
            
            Write-Log "DEBUG: IsSystemAccount: $isSystemAccount, IsAllowAccess: $isAllowAccess"
            
            if (-not $isSystemAccount -and $isAllowAccess) {
                Write-Log "DEBUG: Gültiger Benutzer gefunden: $($permission.Identity)"
                
                $entraIDUPN = Convert-OnPremToEntraID -OnPremUser $permission.Identity -TenantDomain $TenantDomain
                
                # SharePoint-Pfad erstellen wenn Site-URL angegeben ist
                $sharePointPath = ""
                if (-not [string]::IsNullOrWhiteSpace($SharePointSiteURL)) {
                    $sharePointPath = Convert-LocalPathToSharePoint -LocalPath $folder.Path -SharePointSiteURL $SharePointSiteURL -AllFolderPaths $allFolderPaths
                }
                
                $mapping = [SharePointMapping]::new(
                    $permission.Identity,
                    $entraIDUPN,
                    $permission.Rights.ToString(),
                    $folder.Path,
                    $sharePointPath
                )
                
                $mappings += $mapping
                Write-Log "DEBUG: Mapping hinzugefügt für: $($permission.Identity) -> $entraIDUPN"
            } else {
                Write-Log "DEBUG: Benutzer übersprungen: $($permission.Identity) (System: $isSystemAccount, Allow: $isAllowAccess)"
            }
        }
    }
    
    return $mappings
}

# Funktion: CSV-Export für UPN-Mapping
function Export-UPNMappingToCSV {
    param(
        [array]$UPNMappingData,
        [string]$OutputPath
    )
    
    # CSV mit deutschen Spaltennamen und Semikolon-Trennung für Excel (nur SharePoint-relevante Daten)
    $UPNMappingData | Select-Object @{Name="OnPrem_Benutzer";Expression={$_.OnPremUser}},
                                   @{Name="EntraID_UPN";Expression={$_.SharePointUPN}},
                                   @{Name="Berechtigung";Expression={$_.Permission}},
                                   @{Name="SharePoint_Pfad";Expression={$_.SharePointPath}} |
                      Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

# HTML-Export für SharePoint wurde entfernt - nur CSV-Export verfügbar

# Funktion: Benutzer-Mappings aus UPN-Mapping laden
function Import-UserMappingsFromUPN {
    param([array]$SharePointData)
    
    $Global:UserMappings = @()
    
    # Eindeutige Benutzer aus SharePoint-Daten extrahieren
    $uniqueUsers = $SharePointData | Select-Object OnPremUser, SharePointUPN -Unique
    
    foreach ($user in $uniqueUsers) {
        $mapping = [UserMapping]::new(
            $user.OnPremUser,
            $user.SharePointUPN,
            "Automatisch",
            "UPN-Mapping"
        )
        $Global:UserMappings += $mapping
    }
}

# Funktion: Manuelles Benutzer-Mapping hinzufügen
function Add-ManualUserMapping {
    param(
        [string]$OnPremUser,
        [string]$EntraIDUPN
    )
    
    # Prüfen ob Benutzer bereits existiert
    $existing = $Global:UserMappings | Where-Object { $_.OnPremUser -eq $OnPremUser }
    
    if ($existing) {
        # Vorhandenes Mapping aktualisieren
        $existing.EntraIDUPN = $EntraIDUPN
        $existing.Status = "Manuell"
        $existing.Source = "Benutzer"
        return "Aktualisiert"
    } else {
        # Neues Mapping hinzufügen
        $mapping = [UserMapping]::new(
            $OnPremUser,
            $EntraIDUPN,
            "Manuell",
            "Benutzer"
        )
        $Global:UserMappings += $mapping
        return "Hinzugefügt"
    }
}

# Funktion: Angepasste SharePoint-Daten mit manuellen Mappings erstellen
function New-CustomSharePointMapping {
    param([array]$OriginalSharePointData)
    
    $customData = @()
    
    foreach ($original in $OriginalSharePointData) {
        # Prüfen ob manuelles Mapping existiert
        $manualMapping = $Global:UserMappings | Where-Object { $_.OnPremUser -eq $original.OnPremUser }
        
        if ($manualMapping) {
            # Manuelles Mapping verwenden
            $customMapping = [SharePointMapping]::new(
                $original.OnPremUser,
                $manualMapping.EntraIDUPN,
                $original.Permission,
                $original.FolderPath,
                $original.SharePointPath
            )
        } else {
            # Original-Mapping verwenden
            $customMapping = $original
        }
        
        $customData += $customMapping
    }
    
    return $customData
}

# Funktion: CSV-Export für angepasste Mappings
function Export-CustomMappingToCSV {
    param(
        [array]$CustomMappingData,
        [string]$OutputPath
    )
    
    # CSV mit deutschen Spaltennamen und Anpassungs-Info (nur SharePoint-relevante Daten)
    $CustomMappingData | Select-Object @{Name="OnPrem_Benutzer";Expression={$_.OnPremUser}},
                                      @{Name="EntraID_UPN";Expression={$_.SharePointUPN}},
                                      @{Name="Berechtigung";Expression={$_.Permission}},
                                      @{Name="SharePoint_Pfad";Expression={$_.SharePointPath}},
                                      @{Name="Anpassung";Expression={
                                          $manual = $Global:UserMappings | Where-Object { $_.OnPremUser -eq $_.OnPremUser }
                                          if ($manual -and $manual.Status -eq "Manuell") { "Manuell angepasst" } else { "Automatisch" }
                                      }} |
                         Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

# Hauptfunktion: GUI erstellen und anzeigen
function Show-MainWindow {
    # XAML laden
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Controls referenzieren
    $btnSelectFolder = $window.FindName("BtnSelectFolder")
    $lstSelectedFolders = $window.FindName("LstSelectedFolders")
    $btnRemoveFolder = $window.FindName("BtnRemoveFolder")
    $btnAnalyzePermissions = $window.FindName("BtnAnalyzePermissions")
    $btnExportHTML = $window.FindName("BtnExportHTML")
    $treePermissions = $window.FindName("TreePermissions")
    
    $txtTenantDomain = $window.FindName("TxtTenantDomain")
    $txtSharePointSiteURL = $window.FindName("TxtSharePointSiteURL")
    $btnCreateMapping = $window.FindName("BtnCreateMapping")
    $dgSharePointMapping = $window.FindName("DgSharePointMapping")
    $btnExportUPNCSV = $window.FindName("BtnExportUPNCSV")
    
    # Tab 3: Benutzer-Anpassung Controls
    $txtOnPremUser = $window.FindName("TxtOnPremUser")
    $txtEntraIDUPN = $window.FindName("TxtEntraIDUPN")
    $btnAddMapping = $window.FindName("BtnAddMapping")
    $btnLoadFromMapping = $window.FindName("BtnLoadFromMapping")
    $btnClearMappings = $window.FindName("BtnClearMappings")
    $dgUserMappings = $window.FindName("DgUserMappings")
    $btnApplyMappings = $window.FindName("BtnApplyMappings")
    $btnExportCustomCSV = $window.FindName("BtnExportCustomCSV")
    
    $txtStatus = $window.FindName("TxtStatus")
    $progressBar = $window.FindName("ProgressBar")
    
    # Event Handlers
    
    # Ordner hinzufügen
    $btnSelectFolder.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Ordner für Berechtigungsanalyse auswählen"
        
        # PS2EXE: ShowDialog() Rückgabewert in Variable speichern (verhindert MessageBox)
        $result = $folderDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $folderDialog.SelectedPath
            if ($Global:SelectedFolders -notcontains $selectedPath) {
                $Global:SelectedFolders += $selectedPath
                $lstSelectedFolders.Items.Add($selectedPath)
                $txtStatus.Text = "Ordner hinzugefügt: $selectedPath"
            }
        }
    })
    
    # Ordner entfernen
    $btnRemoveFolder.Add_Click({
        if ($lstSelectedFolders.SelectedItem) {
            $selectedPath = $lstSelectedFolders.SelectedItem
            $Global:SelectedFolders = $Global:SelectedFolders | Where-Object { $_ -ne $selectedPath }
            $lstSelectedFolders.Items.Remove($selectedPath)
            
            # Berechtigungsdaten für entfernten Ordner auch löschen
            $Global:PermissionsData = $Global:PermissionsData | Where-Object { $_.Path -ne $selectedPath }
            
            # TreeView aktualisieren
            Update-PermissionsTreeView -TreeView $treePermissions -PermissionsData $Global:PermissionsData
            
            # SharePoint-Daten und Benutzer-Mappings zurücksetzen
            $Global:SharePointData = @()
            $Global:UserMappings = @()
            $Global:CustomSharePointData = @()
            
            # UI zurücksetzen
            $dgSharePointMapping.ItemsSource = $null
            $dgUserMappings.ItemsSource = $null
            $btnExportUPNCSV.IsEnabled = $false
            $btnApplyMappings.IsEnabled = $false
            $btnExportCustomCSV.IsEnabled = $false
            
            $txtStatus.Text = "Ordner entfernt: $selectedPath (Berechtigungsdaten aktualisiert)"
        }
    })
    
    # Berechtigungen analysieren
    $btnAnalyzePermissions.Add_Click({
        if ($Global:SelectedFolders.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens einen Ordner aus.", "Fehler", "OK", "Warning")
            return
        }
        
        $txtStatus.Text = "Analysiere Berechtigungen..."
        $progressBar.Visibility = "Visible"
        
        # Alle Daten komplett zurücksetzen
        $Global:PermissionsData = @()
        $Global:SharePointData = @()
        $Global:UserMappings = @()
        $Global:CustomSharePointData = @()
        
        # UI zurücksetzen
        $dgSharePointMapping.ItemsSource = $null
        $dgUserMappings.ItemsSource = $null
        $btnExportUPNCSV.IsEnabled = $false
        $btnApplyMappings.IsEnabled = $false
        $btnExportCustomCSV.IsEnabled = $false
        
        foreach ($folder in $Global:SelectedFolders) {
            $txtStatus.Text = "Analysiere: $folder"
            $folderPermissions = Get-FolderPermissions -FolderPath $folder -Recursive $true
            $Global:PermissionsData += $folderPermissions
        }
        
        Update-PermissionsTreeView -TreeView $treePermissions -PermissionsData $Global:PermissionsData
        
        $btnExportHTML.IsEnabled = $true
        $progressBar.Visibility = "Collapsed"
        $txtStatus.Text = "Analyse abgeschlossen. $($Global:PermissionsData.Count) Ordner analysiert. Bereit für UPN-Mapping."
    })
    
    # HTML Export
    $btnExportHTML.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "HTML Dateien (*.html)|*.html"
        $saveDialog.FileName = "Ordnerberechtigungen_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        # PS2EXE: ShowDialog() Rückgabewert in Variable speichern (verhindert MessageBox)
        $result = $saveDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-PermissionsToHTML -PermissionsData $Global:PermissionsData -OutputPath $saveDialog.FileName
            $txtStatus.Text = "HTML-Export erstellt: $($saveDialog.FileName)"
            [System.Windows.MessageBox]::Show("HTML-Export erfolgreich erstellt!", "Export", "OK", "Information")
        }
    })
    
    # UPN-Mapping erstellen
    $btnCreateMapping.Add_Click({
        try {
            Write-Log "DEBUG: UPN-Mapping Button geklickt"
            
            if ([string]::IsNullOrWhiteSpace($txtTenantDomain.Text)) {
                [System.Windows.MessageBox]::Show("Bitte geben Sie eine Tenant-Domain ein.", "Fehler", "OK", "Warning")
                return
            }
            
            if ($Global:PermissionsData.Count -eq 0) {
                [System.Windows.MessageBox]::Show("Bitte analysieren Sie zuerst die Ordnerberechtigungen.", "Fehler", "OK", "Warning")
                return
            }
            
            $txtStatus.Text = "Erstelle UPN-Mapping..."
            Write-Log "DEBUG: Starte UPN-Mapping mit $($Global:PermissionsData.Count) Ordnern"
            
            $Global:SharePointData = New-UPNMapping -PermissionsData $Global:PermissionsData -TenantDomain $txtTenantDomain.Text -SharePointSiteURL $txtSharePointSiteURL.Text
            
            Write-Log "DEBUG: UPN-Mapping erstellt mit $($Global:SharePointData.Count) Einträgen"
            
            $dgSharePointMapping.ItemsSource = $Global:SharePointData
            
            # Automatisch Benutzer-Mappings für Tab 3 erstellen
            Import-UserMappingsFromUPN -SharePointData $Global:SharePointData
            $dgUserMappings.ItemsSource = $null
            $dgUserMappings.ItemsSource = $Global:UserMappings
            
            $btnExportUPNCSV.IsEnabled = $true
            $btnApplyMappings.IsEnabled = $true
            
            $txtStatus.Text = "UPN-Mapping erstellt. $($Global:SharePointData.Count) Benutzer-Mappings gefunden und in Tab 3 geladen."
        }
        catch {
            $txtStatus.Text = "Fehler beim Erstellen des UPN-Mappings: $($_.Exception.Message)"
            Write-Log "DEBUG: Fehler - $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Fehler beim Erstellen des UPN-Mappings: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
    })
    
    # UPN-Mapping CSV Export
    $btnExportUPNCSV.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Dateien (*.csv)|*.csv"
        $saveDialog.FileName = "UPN_Mapping_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        # PS2EXE: ShowDialog() Rückgabewert in Variable speichern (verhindert MessageBox)
        $result = $saveDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # Verwende die aktuellen (möglicherweise angepassten) SharePoint-Daten
            Export-UPNMappingToCSV -UPNMappingData $Global:SharePointData -OutputPath $saveDialog.FileName
            $txtStatus.Text = "CSV-Export erstellt: $($saveDialog.FileName)"
            
            # Prüfen ob manuelle Anpassungen enthalten sind
            $hasManualChanges = $Global:UserMappings | Where-Object { $_.Status -eq "Manuell" }
            if ($hasManualChanges) {
                [System.Windows.MessageBox]::Show("UPN-Mapping CSV mit manuellen Anpassungen erfolgreich erstellt!", "Export", "OK", "Information")
            } else {
                [System.Windows.MessageBox]::Show("UPN-Mapping CSV erfolgreich erstellt!", "Export", "OK", "Information")
            }
        }
    })
    
    # HTML-Export für SharePoint wurde entfernt
    
    # === TAB 3: BENUTZER-ANPASSUNG EVENT HANDLERS ===
    
    # Manuelles Mapping hinzufügen
    $btnAddMapping.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtOnPremUser.Text) -or [string]::IsNullOrWhiteSpace($txtEntraIDUPN.Text)) {
            [System.Windows.MessageBox]::Show("Bitte füllen Sie beide Felder aus.", "Fehler", "OK", "Warning")
            return
        }
        
        $result = Add-ManualUserMapping -OnPremUser $txtOnPremUser.Text -EntraIDUPN $txtEntraIDUPN.Text
        
        # DataGrid aktualisieren
        $dgUserMappings.ItemsSource = $null
        $dgUserMappings.ItemsSource = $Global:UserMappings
        
        # Eingabefelder leeren
        $txtOnPremUser.Text = ""
        $txtEntraIDUPN.Text = ""
        
        $btnApplyMappings.IsEnabled = $true
        $txtStatus.Text = "Benutzer-Mapping $result"
    })
    
    # Mappings aus UPN-Mapping neu laden/aktualisieren
    $btnLoadFromMapping.Add_Click({
        if ($Global:SharePointData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte erstellen Sie zuerst ein UPN-Mapping in Tab 2.", "Fehler", "OK", "Warning")
            return
        }
        
        # Bestehende manuelle Änderungen beibehalten
        $existingManual = $Global:UserMappings | Where-Object { $_.Status -eq "Manuell" }
        
        # Neue automatische Mappings laden
        Import-UserMappingsFromUPN -SharePointData $Global:SharePointData
        
        # Manuelle Änderungen wieder hinzufügen
        foreach ($manual in $existingManual) {
            $existing = $Global:UserMappings | Where-Object { $_.OnPremUser -eq $manual.OnPremUser }
            if ($existing) {
                $existing.EntraIDUPN = $manual.EntraIDUPN
                $existing.Status = "Manuell"
                $existing.Source = "Benutzer"
            }
        }
        
        # DataGrid aktualisieren
        $dgUserMappings.ItemsSource = $null
        $dgUserMappings.ItemsSource = $Global:UserMappings
        
        $btnApplyMappings.IsEnabled = $true
        $txtStatus.Text = "$($Global:UserMappings.Count) Benutzer-Mappings aktualisiert (manuelle Änderungen beibehalten)"
    })
    
    # Alle Mappings löschen
    $btnClearMappings.Add_Click({
        $result = [System.Windows.MessageBox]::Show("Möchten Sie wirklich alle Benutzer-Mappings löschen?", "Bestätigung", "YesNo", "Question")
        if ($result -eq "Yes") {
            $Global:UserMappings = @()
            $dgUserMappings.ItemsSource = $null
            $btnApplyMappings.IsEnabled = $false
            $btnExportCustomCSV.IsEnabled = $false
            $txtStatus.Text = "Alle Benutzer-Mappings gelöscht"
        }
    })
    
    # Zuordnungen anwenden
    $btnApplyMappings.Add_Click({
        if ($Global:SharePointData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte erstellen Sie zuerst ein UPN-Mapping in Tab 2.", "Fehler", "OK", "Warning")
            return
        }
        
        $txtStatus.Text = "Wende manuelle Zuordnungen an..."
        $Global:CustomSharePointData = New-CustomSharePointMapping -OriginalSharePointData $Global:SharePointData
        
        # Tab 2 mit angepassten UPNs aktualisieren
        $Global:SharePointData = $Global:CustomSharePointData
        $dgSharePointMapping.ItemsSource = $null
        $dgSharePointMapping.ItemsSource = $Global:SharePointData
        
        $btnExportCustomCSV.IsEnabled = $true
        $txtStatus.Text = "Manuelle Zuordnungen angewendet und in Tab 2 aktualisiert. $($Global:CustomSharePointData.Count) Einträge bereit für Export."
    })
    
    # Angepasste CSV exportieren
    $btnExportCustomCSV.Add_Click({
        if (-not $Global:CustomSharePointData -or $Global:CustomSharePointData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte wenden Sie zuerst die Zuordnungen an.", "Fehler", "OK", "Warning")
            return
        }
        
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Dateien (*.csv)|*.csv"
        $saveDialog.FileName = "Angepasste_UPN_Mappings_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        # PS2EXE: ShowDialog() Rückgabewert in Variable speichern (verhindert MessageBox)
        $result = $saveDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-CustomMappingToCSV -CustomMappingData $Global:CustomSharePointData -OutputPath $saveDialog.FileName
            $txtStatus.Text = "Angepasste CSV erstellt: $($saveDialog.FileName)"
            [System.Windows.MessageBox]::Show("Angepasste UPN-Mapping CSV erfolgreich erstellt!", "Export", "OK", "Information")
        }
    })
    
    # Fenster anzeigen (VOID verhindert "False" MessageBox bei -noConsole)
    [VOID]$window.ShowDialog()
}

# Script starten
try {
    Show-MainWindow
}
catch {
    [System.Windows.MessageBox]::Show("Fehler beim Starten der Anwendung: $($_.Exception.Message)", "Fehler", "OK", "Error")
}

