#Requires -Version 5.1
<#
.SYNOPSIS
    easyFolder SharePoint Online Permissions Manager V0.0.2 (BETA) - CSV Import und SPO Berechtigungsverwaltung
    
.DESCRIPTION
    Diese Anwendung importiert CSV-Dateien aus dem easyFPReader und verwaltet SharePoint Online Berechtigungen
    mit einer integrierten WPF GUI. Bietet Verbindung zu SharePoint Online Sites und Benutzer-Validierung.
    
.FEATURES
    - CSV Import von easyFPReader Mapping-Dateien
    - WPF GUI für SharePoint Online Verwaltung
    - SharePoint Online Site-Verbindung mit Admin-Credentials
    - Benutzer-Validierung gegen EntraID
    - Automatische Berechtigungszuweisung auf SPO Sites
    - Admin wird automatisch als Site Admin hinterlegt
    
.AUTHOR
    PhinIT Solutions
    
.VERSION
    0.0.2
    
.DATE
    2026-10-14
    
.REQUIREMENTS
    - Windows PowerShell 5.1 mit PnP.PowerShell 1.12.x  ODER  PowerShell 7.4+ mit PnP.PowerShell 2.x/3.x
    - Eigene Entra ID App-Registrierung (Client ID) - die frühere "PnP Management Shell"-App existiert nicht mehr
    - SharePoint Online Admin-Berechtigung
    - Datei ist UTF-8 mit BOM gespeichert (Umlaute/Emojis unter Windows PowerShell 5.1)
#>

# PnP PowerShell Update-Check deaktivieren
$env:PNPPOWERSHELL_UPDATECHECK = "Off"

# Assembly-Imports für WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Globale Variablen
$Global:ImportedCSVData = @()
$Global:SPOConnection = $null
$Global:SPOSiteURL = ""
$Global:SPOWebUrl = ""
$Global:SPOUserCache = $null
$Global:SPORoleNameCache = @{}
$Global:SPOListCache = @{}
$Global:ProcessedSPOFolders = @{}
$Global:AppRegistration = @{
    ClientId = $null
    TenantId = $null
    SiteURL = $null
    AdminUPN = $null
}

# Registry-Pfad für Einstellungen
$Global:RegistryPath = "HKCU:\Software\PhinIT\easyFolder-SPOPermissions"

# Klassen für Datenstrukturen
class SPOPermissionEntry {
    [string]$OnPremUser
    [string]$EntraIDUPN
    [string]$Permission
    [string]$SharePointPath
    [string]$ValidationStatus
    [string]$ApplyStatus
    
    SPOPermissionEntry([string]$onprem, [string]$entraid, [string]$permission, [string]$spPath) {
        $this.OnPremUser = $onprem
        $this.EntraIDUPN = $entraid
        $this.Permission = $permission
        $this.SharePointPath = $spPath
        $this.ValidationStatus = "Nicht geprüft"
        $this.ApplyStatus = "Ausstehend"
    }
}

# XAML für die WPF GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="easyFolder SharePoint Online Permissions Manager V0.0.2 (BETA)" 
        Height="900" Width="1400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#1B4F72" CornerRadius="5" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="easyFolder SharePoint Online Permissions Manager V0.0.2 (BETA)" 
                          FontSize="22" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center"/>
                <TextBlock Text="CSV Import und SharePoint Online Berechtigungsverwaltung" 
                          FontSize="13" Foreground="#AED6F1" HorizontalAlignment="Center" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Name="MainTabControl">
            
            <!-- Tab 1: CSV Import -->
            <TabItem Header="📊 CSV Import" FontSize="14">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Import Bereich -->
                    <Border Grid.Row="0" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,10">
                        <StackPanel>
                            <TextBlock Text="CSV-Datei Import" FontWeight="Bold" FontSize="16" Margin="0,0,0,15"/>
                            
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                
                                <TextBox Grid.Column="0" Name="TxtCSVFilePath" Height="30" 
                                        Text="Wählen Sie eine CSV-Datei aus dem easyFPReader..." 
                                        IsReadOnly="True" Background="#F8F9FA" Margin="0,0,10,0"/>
                                <Button Grid.Column="1" Name="BtnSelectCSV" Content="📂 CSV auswählen" 
                                       Height="30" Width="120" Margin="0,0,10,0" Background="#3498DB" Foreground="White" 
                                       BorderThickness="0" FontWeight="Bold"/>
                                <Button Grid.Column="2" Name="BtnImportCSV" Content="📥 Importieren" 
                                       Height="30" Width="100" Background="#27AE60" Foreground="White" 
                                       BorderThickness="0" FontWeight="Bold" IsEnabled="False"/>
                            </Grid>
                            
                            <TextBlock Text="Unterstützte Formate: UPN-Mapping CSV, Angepasste UPN-Mapping CSV aus easyFPReader" 
                                      FontSize="11" Foreground="Gray" Margin="0,8,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Importierte Daten Anzeige -->
                    <Border Grid.Row="1" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBlock Text="Importierte Berechtigungen" FontWeight="Bold" FontSize="14" VerticalAlignment="Center"/>
                                <TextBlock Name="TxtImportCount" Text="(0 Einträge)" FontSize="12" Foreground="Gray" 
                                          VerticalAlignment="Center" Margin="10,0,0,0"/>
                                <Button Name="BtnClearImport" Content="🗑️ Leeren" 
                                       Height="25" Width="80" Margin="20,0,0,0" Background="#E74C3C" Foreground="White" 
                                       BorderThickness="0" IsEnabled="False"/>
                            </StackPanel>
                            
                            <DataGrid Grid.Row="1" Name="DgImportedData" AutoGenerateColumns="False" 
                                     CanUserAddRows="False" CanUserDeleteRows="True" GridLinesVisibility="Horizontal"
                                     HeadersVisibility="Column" AlternatingRowBackground="#F8F9FA">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="OnPrem Benutzer" Binding="{Binding OnPremUser}" Width="200"/>
                                    <DataGridTextColumn Header="EntraID UPN" Binding="{Binding EntraIDUPN}" Width="250"/>
                                    <DataGridTextColumn Header="Berechtigung" Binding="{Binding Permission}" Width="150"/>
                                    <DataGridTextColumn Header="SharePoint Pfad" Binding="{Binding SharePointPath}" Width="*"/>
                                    <DataGridTextColumn Header="Status" Binding="{Binding ValidationStatus}" Width="120"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                    
                    <!-- Import Aktionen -->
                    <Border Grid.Row="2" Padding="15" Margin="0,10,0,0">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Name="BtnValidateUsers" Content="✅ Benutzer validieren" 
                                   Height="35" Width="160" Margin="0,0,10,0" Background="#8E44AD" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                            <Button Name="BtnProceedToConnection" Content="➡️ Weiter zu Verbindung" 
                                   Height="35" Width="180" Background="#16A085" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>
            
            <!-- Tab 2: SharePoint Verbindung -->
            <TabItem Header="🔗 SharePoint Verbindung" FontSize="14">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Verbindungseinstellungen -->
                    <Border Grid.Row="0" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="20" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock Text="SharePoint Online Verbindungseinstellungen" FontWeight="Bold" FontSize="16" Margin="0,0,0,15"/>
                            
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <!-- SharePoint Site URL -->
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Site URL:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                <TextBox Grid.Row="0" Grid.Column="1" Name="TxtSPOSiteURL" Height="30" Margin="0,0,0,10"
                                        Text="https://contoso.sharepoint.com/sites/YourSite"/>
                                
                                <!-- Admin UPN -->
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Admin UPN:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                <TextBox Grid.Row="1" Grid.Column="1" Name="TxtAdminUPN" Height="30" Margin="0,0,0,10"
                                        Text="admin@contoso.onmicrosoft.com"/>
                                
                                <!-- Tenant Domain -->
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Tenant Domain:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <TextBox Grid.Row="2" Grid.Column="1" Name="TxtTenantDomain" Height="30"
                                        Text="contoso.onmicrosoft.com"/>
                            </Grid>
                            
                            <TextBlock Text="Der Admin-Benutzer wird als Site-Collection-Administrator hinzugefügt, sofern die Option in Schritt 3 aktiviert ist." 
                                      FontSize="11" Foreground="Gray" Margin="0,10,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Verbindungsstatus -->
                    <Border Grid.Row="1" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock Text="Verbindungsstatus:" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <Ellipse Name="ConnectionStatusIndicator" Width="12" Height="12" Fill="#E74C3C" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <TextBlock Name="TxtConnectionStatus" Text="Nicht verbunden" VerticalAlignment="Center" Foreground="#E74C3C"/>
                            </StackPanel>
                            
                            <Button Grid.Column="1" Name="BtnAppSetup" Content="🔧 App-Setup" 
                                   Height="30" Width="100" Margin="0,0,10,0" Background="#9B59B6" Foreground="White" 
                                   BorderThickness="0"/>
                            <Button Grid.Column="2" Name="BtnConnect" Content="🔗 Verbinden" 
                                   Height="30" Width="100" Background="#27AE60" Foreground="White" 
                                   BorderThickness="0"/>
                        </Grid>
                    </Border>
                    
                    <!-- Site Informationen -->
                    <Border Grid.Row="2" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Site Informationen" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                            
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                                <TextBlock Name="TxtSiteInfo" Text="Verbinden Sie sich mit einer SharePoint Site, um Informationen anzuzeigen..." 
                                          FontFamily="Consolas" FontSize="11" Background="#F8F9FA" Padding="10" 
                                          TextWrapping="Wrap"/>
                            </ScrollViewer>
                        </Grid>
                    </Border>
                    
                    <!-- Verbindungsaktionen -->
                    <Border Grid.Row="3" Padding="15" Margin="0,10,0,0">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button Name="BtnDisconnect" Content="❌ Trennen" 
                                   Height="35" Width="100" Margin="0,0,10,0" Background="#E74C3C" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                            <Button Name="BtnProceedToPermissions" Content="➡️ Weiter zu Berechtigungen" 
                                   Height="35" Width="200" Background="#16A085" Foreground="White" 
                                   BorderThickness="0" IsEnabled="False"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>
            
            <!-- Tab 3: Berechtigungen anwenden -->
            <TabItem Header="🔐 Berechtigungen anwenden" FontSize="14">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Anwendungsoptionen -->
                    <Border Grid.Row="0" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,10">
                        <StackPanel>
                            <TextBlock Text="Berechtigungen auf SharePoint Site anwenden" FontWeight="Bold" FontSize="16" Margin="0,0,0,15"/>
                            
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                
                                <StackPanel Grid.Column="0">
                                    <CheckBox Name="ChkCreateFolders" Content="Ordner automatisch erstellen (falls nicht vorhanden)" 
                                             IsChecked="True" Margin="0,0,0,5"/>
                                    <CheckBox Name="ChkBreakInheritance" Content="Bestehende Berechtigungen am Zielordner ersetzen (sonst werden sie beim Unterbrechen der Vererbung kopiert)" 
                                             IsChecked="True" Margin="0,0,0,5"
                                             ToolTip="Ordner-Berechtigungen unterbrechen immer die Vererbung. Aktiviert: geerbte Berechtigungen werden einmalig pro Ordner entfernt."/>
                                    <CheckBox Name="ChkAddAdminToSite" Content="Admin-Benutzer als Site-Administrator hinzufügen" 
                                             IsChecked="True" Margin="0,0,0,5"/>
                                </StackPanel>
                                
                                <Button Grid.Column="1" Name="BtnValidateAllUsers" Content="👥 Alle Benutzer validieren" 
                                       Height="35" Width="160" Margin="10,0,10,0" Background="#9B59B6" Foreground="White" 
                                       BorderThickness="0"/>
                                <Button Grid.Column="2" Name="BtnApplyPermissions" Content="🚀 Berechtigungen anwenden" 
                                       Height="35" Width="180" Background="#E67E22" Foreground="White" 
                                       BorderThickness="0" IsEnabled="False"/>
                            </Grid>
                        </StackPanel>
                    </Border>
                    
                    <!-- Berechtigungen Übersicht -->
                    <Border Grid.Row="1" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBlock Text="Berechtigungen Status" FontWeight="Bold" FontSize="14" VerticalAlignment="Center"/>
                                <TextBlock Name="TxtPermissionCount" Text="(0 Einträge)" FontSize="12" Foreground="Gray" 
                                          VerticalAlignment="Center" Margin="10,0,0,0"/>
                                <Button Name="BtnRefreshStatus" Content="🔄 Status aktualisieren" 
                                       Height="25" Width="130" Margin="20,0,0,0" Background="#3498DB" Foreground="White" 
                                       BorderThickness="0"/>
                            </StackPanel>
                            
                            <DataGrid Grid.Row="1" Name="DgPermissionStatus" AutoGenerateColumns="False" 
                                     CanUserAddRows="False" CanUserDeleteRows="False" GridLinesVisibility="Horizontal"
                                     HeadersVisibility="Column" AlternatingRowBackground="#F8F9FA">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="OnPrem Benutzer" Binding="{Binding OnPremUser}" Width="180"/>
                                    <DataGridTextColumn Header="EntraID UPN" Binding="{Binding EntraIDUPN}" Width="220"/>
                                    <DataGridTextColumn Header="Berechtigung" Binding="{Binding Permission}" Width="120"/>
                                    <DataGridTextColumn Header="SharePoint Pfad" Binding="{Binding SharePointPath}" Width="*"/>
                                    <DataGridTextColumn Header="Validierung" Binding="{Binding ValidationStatus}" Width="100"/>
                                    <DataGridTextColumn Header="Anwendung" Binding="{Binding ApplyStatus}" Width="100"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                    
                    <!-- Fortschritt und Aktionen -->
                    <Border Grid.Row="2" Padding="15" Margin="0,10,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <!-- Fortschrittsbalken -->
                            <StackPanel Grid.Row="0" Margin="0,0,0,10">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <ProgressBar Grid.Column="0" Name="ProgressBarPermissions" Height="20" 
                                               Background="#ECF0F1" Foreground="#27AE60" Margin="0,0,10,0"/>
                                    <TextBlock Grid.Column="1" Name="TxtProgressPercent" Text="0%" 
                                              VerticalAlignment="Center" FontWeight="Bold"/>
                                </Grid>
                                <TextBlock Name="TxtCurrentAction" Text="Bereit für Berechtigungsanwendung..." 
                                          FontSize="11" Foreground="Gray" Margin="0,5,0,0"/>
                            </StackPanel>
                            
                            <!-- Aktionsbuttons -->
                            <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center">
                                <Button Name="BtnExportResults" Content="📊 Ergebnisse exportieren" 
                                       Height="35" Width="160" Margin="0,0,10,0" Background="#34495E" Foreground="White" 
                                       BorderThickness="0" IsEnabled="False"/>
                                <Button Name="BtnResetAll" Content="🔄 Alles zurücksetzen" 
                                       Height="35" Width="140" Background="#E74C3C" Foreground="White" 
                                       BorderThickness="0"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Status Bar -->
        <Border Grid.Row="2" Background="#2C3E50" CornerRadius="3" Padding="12" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Name="TxtStatus" Text="Bereit - Wählen Sie eine CSV-Datei zum Import aus..." 
                          Foreground="White" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Name="TxtModuleStatus" Text="PnP.PowerShell: Nicht geladen" 
                          Foreground="#F39C12" VerticalAlignment="Center" Margin="0,0,15,0" FontSize="11"/>
                <ProgressBar Grid.Column="2" Name="ProgressBarMain" Width="200" Height="15" Visibility="Collapsed"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Funktion: Einstellungen aus Registry laden
function Load-SettingsFromRegistry {
    try {
        if (Test-Path $Global:RegistryPath) {
            $clientId = Get-ItemProperty -Path $Global:RegistryPath -Name "ClientId" -ErrorAction SilentlyContinue
            $tenantId = Get-ItemProperty -Path $Global:RegistryPath -Name "TenantId" -ErrorAction SilentlyContinue
            $siteURL = Get-ItemProperty -Path $Global:RegistryPath -Name "SiteURL" -ErrorAction SilentlyContinue
            $adminUPN = Get-ItemProperty -Path $Global:RegistryPath -Name "AdminUPN" -ErrorAction SilentlyContinue
            
            if ($clientId) { $Global:AppRegistration.ClientId = $clientId.ClientId }
            if ($tenantId) { $Global:AppRegistration.TenantId = $tenantId.TenantId }
            if ($siteURL) { $Global:AppRegistration.SiteURL = $siteURL.SiteURL }
            if ($adminUPN) { $Global:AppRegistration.AdminUPN = $adminUPN.AdminUPN }
            
            return @{
                Success = $true
                Message = "Einstellungen aus Registry geladen"
                HasClientId = (-not [string]::IsNullOrWhiteSpace($Global:AppRegistration.ClientId))
            }
        } else {
            return @{
                Success = $true
                Message = "Keine gespeicherten Einstellungen gefunden"
                HasClientId = $false
            }
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Fehler beim Laden der Einstellungen: $($_.Exception.Message)"
            HasClientId = $false
        }
    }
}

# Funktion: Einstellungen in Registry speichern
function Save-SettingsToRegistry {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [string]$SiteURL,
        [string]$AdminUPN
    )
    
    try {
        # Registry-Pfad erstellen falls nicht vorhanden
        if (-not (Test-Path $Global:RegistryPath)) {
            New-Item -Path $Global:RegistryPath -Force | Out-Null
        }
        
        # Einstellungen speichern
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
            Set-ItemProperty -Path $Global:RegistryPath -Name "ClientId" -Value $ClientId
            $Global:AppRegistration.ClientId = $ClientId
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            Set-ItemProperty -Path $Global:RegistryPath -Name "TenantId" -Value $TenantId
            $Global:AppRegistration.TenantId = $TenantId
        }
        if (-not [string]::IsNullOrWhiteSpace($SiteURL)) {
            Set-ItemProperty -Path $Global:RegistryPath -Name "SiteURL" -Value $SiteURL
            $Global:AppRegistration.SiteURL = $SiteURL
        }
        if (-not [string]::IsNullOrWhiteSpace($AdminUPN)) {
            Set-ItemProperty -Path $Global:RegistryPath -Name "AdminUPN" -Value $AdminUPN
            $Global:AppRegistration.AdminUPN = $AdminUPN
        }
        
        return @{
            Success = $true
            Message = "Einstellungen erfolgreich gespeichert"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Fehler beim Speichern der Einstellungen: $($_.Exception.Message)"
        }
    }
}

# Funktion: App-Registrierung Setup Dialog
function Show-AppRegistrationDialog {
    $dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SharePoint App-Registrierung Setup" 
        Height="600" Width="700"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2C3E50" CornerRadius="5" Padding="15" Margin="0,0,0,20">
            <StackPanel>
                <TextBlock Text="🔧 SharePoint App-Registrierung Setup" 
                          FontSize="18" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center"/>
                <TextBlock Text="Erstellen Sie eine eigene App-Registrierung für zuverlässige SharePoint-Verbindungen" 
                          FontSize="12" Foreground="#BDC3C7" HorizontalAlignment="Center" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Anleitung -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,20">
            <StackPanel>
                <TextBlock Text="📋 Schritt-für-Schritt Anleitung:" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
                
                <TextBlock Text="1. Gehen Sie zu https://entra.microsoft.com" FontWeight="Bold" TextWrapping="Wrap" Margin="0,0,0,10"/>
                
                <TextBlock Text="2. Melden Sie sich mit Ihrem Admin-Account an" TextWrapping="Wrap" Margin="0,0,0,5"/>
                <TextBlock Text="3. Navigieren Sie zu: App registrations → New registration" TextWrapping="Wrap" Margin="0,0,0,5"/>
                <TextBlock Text="4. Füllen Sie aus:" FontWeight="Bold" TextWrapping="Wrap" Margin="0,0,0,5"/>
                <TextBlock Text="   • Name: easyFolder SPO Permissions Manager" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Supported account types: Accounts in this organizational directory only" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Redirect URI: Public client/native → https://login.microsoftonline.com/common/oauth2/nativeclient" TextWrapping="Wrap" Margin="20,0,0,10"/>
                
                <TextBlock Text="5. Nach der Erstellung:" FontWeight="Bold" TextWrapping="Wrap" Margin="0,0,0,5"/>
                <TextBlock Text="   • Kopieren Sie die Application (client) ID" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Kopieren Sie die Directory (tenant) ID" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Gehen Sie zu API permissions" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Add a permission → SharePoint → Delegated permissions" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Wählen Sie: AllSites.FullControl" TextWrapping="Wrap" Margin="20,0,0,5"/>
                <TextBlock Text="   • Klicken Sie Grant admin consent for [Ihr Tenant]" TextWrapping="Wrap" Margin="20,0,0,10"/>
                
                <Border Background="#F8F9FA" BorderBrush="#E9ECEF" BorderThickness="1" CornerRadius="3" Padding="10" Margin="0,10,0,0">
                    <StackPanel>
                        <TextBlock Text="💡 Tipp: Diese Einstellungen werden lokal gespeichert und beim nächsten Start automatisch geladen." 
                                  FontStyle="Italic" TextWrapping="Wrap" Margin="0,0,0,10"/>
                        <Button Name="BtnOpenEntra" Content="🌐 Entra Admin Center öffnen" 
                               Height="30" Width="200" Background="#0078D4" Foreground="White" 
                               BorderThickness="0" HorizontalAlignment="Center"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>
        
        <!-- Eingabefelder -->
        <Border Grid.Row="2" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="150"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Client ID:" VerticalAlignment="Center" FontWeight="Bold"/>
                <TextBox Grid.Row="0" Grid.Column="1" Name="TxtClientId" Height="25" Margin="10,5,0,5"/>
                
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Tenant ID:" VerticalAlignment="Center" FontWeight="Bold"/>
                <TextBox Grid.Row="1" Grid.Column="1" Name="TxtTenantId" Height="25" Margin="10,5,0,5"/>
                
                <TextBlock Grid.Row="2" Grid.Column="0" Text="SharePoint Site URL:" VerticalAlignment="Center" FontWeight="Bold"/>
                <TextBox Grid.Row="2" Grid.Column="1" Name="TxtSiteURL" Height="25" Margin="10,5,0,5"/>
                
                <TextBlock Grid.Row="3" Grid.Column="0" Text="Admin UPN:" VerticalAlignment="Center" FontWeight="Bold"/>
                <TextBox Grid.Row="3" Grid.Column="1" Name="TxtAdminUPN" Height="25" Margin="10,5,0,5"/>
            </Grid>
        </Border>
        
        <!-- Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="BtnSaveAndConnect" Content="💾 Speichern und Verbinden" 
                   Height="35" Width="180" Margin="0,0,10,0" Background="#27AE60" Foreground="White" 
                   BorderThickness="0" FontWeight="Bold"/>
            <Button Name="BtnCancel" Content="❌ Abbrechen" 
                   Height="35" Width="100" Background="#E74C3C" Foreground="White" 
                   BorderThickness="0" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$dialogXaml)
        $dialog = [Windows.Markup.XamlReader]::Load($reader)
        
        # Controls referenzieren
        $txtClientId = $dialog.FindName("TxtClientId")
        $txtTenantId = $dialog.FindName("TxtTenantId")
        $txtSiteURL = $dialog.FindName("TxtSiteURL")
        $txtAdminUPN = $dialog.FindName("TxtAdminUPN")
        $btnSaveAndConnect = $dialog.FindName("BtnSaveAndConnect")
        $btnCancel = $dialog.FindName("BtnCancel")
        $btnOpenEntra = $dialog.FindName("BtnOpenEntra")
        
        # Vorhandene Werte laden
        $txtClientId.Text = [string]$Global:AppRegistration.ClientId
        $txtTenantId.Text = [string]$Global:AppRegistration.TenantId
        $txtSiteURL.Text = [string]$Global:AppRegistration.SiteURL
        $txtAdminUPN.Text = [string]$Global:AppRegistration.AdminUPN
        
        # Event Handlers
        $btnSaveAndConnect.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtClientId.Text) -or 
                [string]::IsNullOrWhiteSpace($txtTenantId.Text) -or
                [string]::IsNullOrWhiteSpace($txtSiteURL.Text) -or
                [string]::IsNullOrWhiteSpace($txtAdminUPN.Text)) {
                
                [System.Windows.MessageBox]::Show("Bitte füllen Sie alle Felder aus.", "Eingabe erforderlich", "OK", "Warning")
                return
            }
            
            # Einstellungen speichern
            $saveResult = Save-SettingsToRegistry -ClientId $txtClientId.Text -TenantId $txtTenantId.Text -SiteURL $txtSiteURL.Text -AdminUPN $txtAdminUPN.Text
            
            if ($saveResult.Success) {
                $dialog.DialogResult = $true
                $dialog.Close()
            } else {
                [System.Windows.MessageBox]::Show($saveResult.Message, "Speicherfehler", "OK", "Error")
            }
        })
        
        $btnCancel.Add_Click({
            $dialog.DialogResult = $false
            $dialog.Close()
        })
        
        $btnOpenEntra.Add_Click({
            try {
                Start-Process "https://entra.microsoft.com"
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler beim Öffnen des Browsers: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        })
        
        # Dialog anzeigen (ShowDialog liefert Nullable[bool] - in echtes bool umwandeln)
        return [bool]$dialog.ShowDialog()
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Öffnen des Setup-Dialogs: $($_.Exception.Message)", "Fehler", "OK", "Error")
        return $false
    }
}

# Funktion: PnP.PowerShell Modul prüfen und laden
function Test-PnPModule {
    try {
        $loaded = Get-Module -Name "PnP.PowerShell"
        if ($loaded) {
            return @{
                IsAvailable = $true
                Version = $loaded.Version.ToString()
                Message = "PnP.PowerShell v$($loaded.Version) geladen"
            }
        }
        
        $module = Get-Module -Name "PnP.PowerShell" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if ($module) {
            # Kein -Force: ein erneutes Laden würde eine bestehende Verbindung zurücksetzen
            Import-Module "PnP.PowerShell" -ErrorAction Stop
            return @{
                IsAvailable = $true
                Version = $module.Version.ToString()
                Message = "PnP.PowerShell v$($module.Version) geladen"
            }
        } else {
            return @{
                IsAvailable = $false
                Version = "Nicht installiert"
                Message = "PnP.PowerShell Modul nicht gefunden. Installieren Sie es mit: Install-Module PnP.PowerShell -Scope CurrentUser  (Windows PowerShell 5.1: -RequiredVersion 1.12.0)"
            }
        }
    }
    catch {
        return @{
            IsAvailable = $false
            Version = "Fehler"
            Message = "Fehler beim Laden: $($_.Exception.Message)"
        }
    }
}

# Funktion: CSV-Datei importieren
function Import-CSVPermissions {
    param(
        [string]$FilePath
    )
    
    try {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
            throw "Datei nicht gefunden: $FilePath"
        }
        
        # CSV mit Semikolon-Trennung importieren (Standard für deutsche Excel-Exporte)
        $csvData = Import-Csv -LiteralPath $FilePath -Delimiter ";" -Encoding UTF8
        
        # Spalten-Mapping für verschiedene CSV-Formate aus easyFPReader
        $mappedData = New-Object System.Collections.Generic.List[object]
        
        foreach ($row in $csvData) {
            $entry = $null
            
            # Format 1: UPN-Mapping CSV (Standard)
            if ($row.PSObject.Properties.Name -contains "OnPrem_Benutzer") {
                $entry = [SPOPermissionEntry]::new(
                    $row.OnPrem_Benutzer,
                    $row.EntraID_UPN,
                    $row.Berechtigung,
                    $row.SharePoint_Pfad
                )
            }
            # Format 2: Angepasste UPN-Mapping CSV (mit Anpassung-Spalte)
            elseif ($row.PSObject.Properties.Name -contains "OnPremUser") {
                $entry = [SPOPermissionEntry]::new(
                    $row.OnPremUser,
                    $row.SharePointUPN,
                    $row.Permission,
                    $row.SharePointPath
                )
            }
            # Format 3: Direkte Spalten ohne Prefix
            elseif ($row.PSObject.Properties.Name -contains "EntraID UPN") {
                $entry = [SPOPermissionEntry]::new(
                    $row."OnPrem Benutzer",
                    $row."EntraID UPN",
                    $row."Berechtigung",
                    $row."SharePoint Pfad"
                )
            }
            
            # Zeilen ohne UPN überspringen (können nicht angewendet werden)
            if ($entry -and -not [string]::IsNullOrWhiteSpace($entry.EntraIDUPN)) {
                $entry.EntraIDUPN = $entry.EntraIDUPN.Trim()
                $mappedData.Add($entry)
            }
        }
        
        if ($mappedData.Count -eq 0) {
            throw "Keine gültigen Berechtigungseinträge in der CSV-Datei gefunden. Überprüfen Sie das Format."
        }
        
        return @{
            Success = $true
            Data = $mappedData.ToArray()
            Count = $mappedData.Count
            Message = "$($mappedData.Count) Berechtigungseinträge erfolgreich importiert"
        }
    }
    catch {
        return @{
            Success = $false
            Data = @()
            Count = 0
            Message = "Fehler beim CSV-Import: $($_.Exception.Message)"
        }
    }
}

# Funktion: SharePoint Online URL vor dem Verbinden prüfen (Format, HTTPS, *.sharepoint.com)
function Test-SPOConnection {
    param(
        [string]$SiteURL,
        [string]$AdminUPN
    )
    
    try {
        # Prüfen ob PnP.PowerShell verfügbar ist
        $moduleCheck = Test-PnPModule
        if (-not $moduleCheck.IsAvailable) {
            throw $moduleCheck.Message
        }
        
        # URL prüfen (ohne Authentifizierung)
        $uri = $null
        if (-not [System.Uri]::TryCreate($SiteURL.Trim(), [System.UriKind]::Absolute, [ref]$uri)) {
            throw "Ungültige URL. Format: https://tenant.sharepoint.com/sites/sitename"
        }
        if ($uri.Scheme -ne "https") {
            throw "SharePoint URL muss HTTPS verwenden"
        }
        
        # Basis-URL Validierung
        if (-not ($uri.Host -like "*.sharepoint.com" -or $uri.Host -like "*.sharepoint.us" -or $uri.Host -like "*.sharepoint.de" -or $uri.Host -like "*.sharepoint.cn")) {
            throw "Ungültige SharePoint Online URL. Format: https://tenant.sharepoint.com/sites/sitename"
        }
        
        return @{
            Success = $true
            Message = "URL-Format ist gültig. Bereit für Verbindung."
            SiteURL = $SiteURL
            AdminUPN = $AdminUPN
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Verbindungstest fehlgeschlagen: $($_.Exception.Message)"
            SiteURL = $SiteURL
            AdminUPN = $AdminUPN
        }
    }
}

# Funktion: SharePoint Online Verbindung herstellen
function Connect-SPOSite {
    param(
        [string]$SiteURL,
        [string]$AdminUPN
    )
    
    try {
        # Bestehende Verbindung trennen und Caches leeren
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        } catch { }
        $Global:SPOWebUrl = ""
        Reset-SPOSessionCache
        
        # App-Registrierung verwenden
        Write-Host "Versuche SharePoint-Verbindung mit App-Registrierung..." -ForegroundColor Yellow
        
        # Prüfen ob App-Registrierung konfiguriert ist
        if ([string]::IsNullOrWhiteSpace($Global:AppRegistration.ClientId)) {
            throw "Keine App-Registrierung konfiguriert. Bitte führen Sie das Setup durch."
        }
        
        Write-Host "Verwende Client ID: $($Global:AppRegistration.ClientId)" -ForegroundColor Green
        Write-Host "Tenant ID: $($Global:AppRegistration.TenantId)" -ForegroundColor Green
        
        $connectionSuccess = $false
        
        # Methode 1: Interactive Login mit App-Registrierung
        try {
            Write-Host "Verwende Interactive Login mit App-Registrierung..." -ForegroundColor Cyan
            Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $Global:AppRegistration.ClientId -ErrorAction Stop
            Write-Host "Interactive Login erfolgreich!" -ForegroundColor Green
            $connectionSuccess = $true
        }
        catch {
            Write-Host "Interactive Login fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Red
            
            # Methode 2: Device Login mit App-Registrierung
            try {
                Write-Host "Versuche Device Login mit App-Registrierung..." -ForegroundColor Yellow
                Write-Host "Gehen Sie zu https://microsoft.com/devicelogin und geben Sie den Code ein." -ForegroundColor Cyan
                Connect-PnPOnline -Url $SiteURL -DeviceLogin -ClientId $Global:AppRegistration.ClientId -ErrorAction Stop
                Write-Host "Device Login erfolgreich!" -ForegroundColor Green
                $connectionSuccess = $true
            }
            catch {
                Write-Host "Device Login fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Red
                
                # Methode 3: Web Login als Fallback (ohne Client ID) - nur PnP.PowerShell 1.x, in 2.x entfernt
                try {
                    Write-Host "Versuche Web Login als Fallback..." -ForegroundColor Yellow
                    Connect-PnPOnline -Url $SiteURL -UseWebLogin -ErrorAction Stop
                    Write-Host "Web Login erfolgreich!" -ForegroundColor Green
                    $connectionSuccess = $true
                }
                catch {
                    $errorMsg = @"
Alle Verbindungsmethoden fehlgeschlagen!

Mögliche Ursachen:
1. App-Registrierung hat keine SharePoint-Berechtigungen
2. Admin Consent wurde nicht erteilt
3. Client ID ist ungültig
4. SharePoint Site URL ist falsch

Überprüfen Sie:
- API permissions → SharePoint → AllSites.FullControl
- Grant admin consent wurde geklickt
- Client ID ist korrekt kopiert
- Site URL Format: https://tenant.sharepoint.com/sites/sitename

Letzter Fehler: $($_.Exception.Message)
"@
                    throw $errorMsg
                }
            }
        }
        
        # Verbindung testen
        $web = Get-PnPWeb -ErrorAction Stop
        $site = Get-PnPSite -ErrorAction Stop
        
        # Web-URL merken: wird zum Umrechnen der CSV-URLs in site-relative Pfade benötigt
        $Global:SPOWebUrl = $web.Url
        
        # Site-Informationen sammeln
        $siteInfo = @{
            Title = $web.Title
            URL = $web.Url
            Description = $web.Description
            Created = $web.Created
            LastModified = $web.LastItemModifiedDate
            Owner = $site.Owner.Email
            StorageUsed = [math]::Round($site.Usage.Storage / 1MB, 2)
            StorageQuota = [math]::Round($site.Usage.StorageWarningLevel / 1MB, 2)
        }
        
        return @{
            Success = $true
            Message = "Erfolgreich mit SharePoint Site verbunden"
            SiteInfo = $siteInfo
            IsConnected = $true
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Verbindung fehlgeschlagen: $($_.Exception.Message)"
            SiteInfo = $null
            IsConnected = $false
        }
    }
}

# Funktion: Benutzer gegen EntraID/SharePoint validieren (nur lesend - legt keine Benutzer an)
function Test-EntraIDUser {
    param(
        [string]$UserUPN
    )
    
    try {
        if ([string]::IsNullOrWhiteSpace($UserUPN) -or $UserUPN -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            return @{
                IsValid = $false
                Status = "Invalid"
                UserInfo = $null
                Message = "Ungültiges UPN-Format"
            }
        }
        
        # Site-Benutzer nur einmal pro Verbindung laden (statt Get-PnPUser je Eintrag)
        if ($null -eq $Global:SPOUserCache) {
            $Global:SPOUserCache = @(Get-PnPUser -ErrorAction Stop)
        }
        
        $claimsLogin = "i:0#.f|membership|$UserUPN"
        $user = $Global:SPOUserCache | Where-Object {
            $_.LoginName -eq $claimsLogin -or $_.Email -eq $UserUPN
        } | Select-Object -First 1
        
        if ($user) {
            return @{
                IsValid = $true
                Status = "Found"
                UserInfo = @{
                    LoginName = $user.LoginName
                    Email = $user.Email
                    Title = $user.Title
                    IsSiteAdmin = $user.IsSiteAdmin
                }
                Message = "Benutzer in der Site gefunden"
            }
        }
        
        return @{
            IsValid = $false
            Status = "NotInSite"
            UserInfo = $null
            Message = "Noch nicht in der Site - wird beim Anwenden per EnsureUser (New-PnPUser) gegen Entra ID aufgelöst"
        }
    }
    catch {
        return @{
            IsValid = $false
            Status = "Error"
            UserInfo = $null
            Message = "Fehler bei Benutzervalidierung: $($_.Exception.Message)"
        }
    }
}

# Funktion: Pfad/URL aus der CSV in einen site-relativen Pfad umwandeln
# Der easyFPReader exportiert vollständige URLs (https://tenant.sharepoint.com/sites/x/Bibliothek/Ordner).
# PnP-Cmdlets erwarten site-relative Pfade (Bibliothek/Ordner). Leerer Rückgabewert = Web-Root.
function ConvertTo-SPOSiteRelativePath {
    param(
        [string]$Path,
        [string]$WebUrl = $Global:SPOWebUrl
    )
    
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    
    $relativePath = $Path.Trim().Replace('\', '/')
    $webUri = $null
    $hasWebUri = -not [string]::IsNullOrWhiteSpace($WebUrl) -and
                 [System.Uri]::TryCreate($WebUrl.TrimEnd('/'), [System.UriKind]::Absolute, [ref]$webUri)
    $webPath = if ($hasWebUri) { [System.Uri]::UnescapeDataString($webUri.AbsolutePath).TrimEnd('/') } else { "" }
    
    if ($relativePath -match '^(?i)https?://([^/]+)(/.*)?$') {
        # Bewusst KEIN System.Uri: Der Reader schreibt unkodierte Ordnernamen. Ein "#" im Namen
        # würde sonst als Fragment abgeschnitten und die Berechtigung am Elternordner landen.
        $pathHost = $Matches[1]
        $relativePath = "$($Matches[2])"
        if (-not $hasWebUri) {
            throw "Die URL der verbundenen Site ist unbekannt - bitte neu verbinden."
        }
        if ($pathHost -ne $webUri.Authority -and $pathHost -ne $webUri.Host) {
            throw "Pfad '$Path' gehört nicht zum verbundenen Tenant ($($webUri.Host))."
        }
        # Aus dem Browser kopierte URLs sind prozentkodiert (%20) - nur diese dekodieren
        if ($relativePath -notmatch ' ' -and $relativePath -match '%[0-9A-Fa-f]{2}') {
            $relativePath = [System.Uri]::UnescapeDataString($relativePath)
        }
        
        $isInWeb = ($webPath -eq "") -or (
            $relativePath.StartsWith($webPath, [System.StringComparison]::OrdinalIgnoreCase) -and
            ($relativePath.Length -eq $webPath.Length -or $relativePath[$webPath.Length] -eq '/')
        )
        if (-not $isInWeb) {
            throw "Pfad '$Path' liegt nicht in der verbundenen Site '$WebUrl'."
        }
        $relativePath = $relativePath.Substring($webPath.Length)
    }
    elseif ($webPath -and $relativePath.StartsWith("$webPath/", [System.StringComparison]::OrdinalIgnoreCase)) {
        # Server-relativer Pfad (/sites/x/Bibliothek/Ordner)
        $relativePath = $relativePath.Substring($webPath.Length)
    }
    
    return $relativePath.Trim('/')
}

# Funktion: Bibliothek (Liste) zu einem site-relativen Pfad ermitteln (mit Cache)
function Get-SPOLibraryForPath {
    param([string]$SiteRelativePath)
    
    $libraryUrl = ($SiteRelativePath -split '/')[0]
    if ([string]::IsNullOrWhiteSpace($libraryUrl)) {
        throw "Im Pfad ist keine Dokumentbibliothek enthalten."
    }
    
    $cacheKey = $libraryUrl.ToLowerInvariant()
    if (-not $Global:SPOListCache.ContainsKey($cacheKey)) {
        # Get-PnPList -Identity akzeptiert Titel oder URL-Namen (z. B. "Shared Documents" / "Freigegebene Dokumente")
        $list = Get-PnPList -Identity $libraryUrl -ErrorAction Stop
        if (-not $list) {
            throw "Dokumentbibliothek '$libraryUrl' wurde in der Site nicht gefunden."
        }
        $Global:SPOListCache[$cacheKey] = $list
    }
    
    return $Global:SPOListCache[$cacheKey]
}

# Funktion: NTFS-Rechte auf SharePoint-Rollentyp abbilden
# Reihenfolge ist wichtig: "Write, ReadAndExecute" muss als Schreibrecht erkannt werden, nicht als Lesen.
function Get-SPORoleTypeForRights {
    param([string]$Rights)
    
    $value = "$Rights"
    if ($value -match 'FullControl|Full Control|GenericAll|^\s*268435456\s*$') { return "Administrator" }
    if ($value -match 'Modify|Change|^\s*-536805376\s*$') { return "Editor" }
    if ($value -match 'Write|AppendData|CreateFiles|CreateDirectories|GenericWrite|^\s*1073741824\s*$') { return "Contributor" }
    return "Reader"
}

# Funktion: Lokalisierten Namen einer Berechtigungsstufe ermitteln (Full Control/Vollzugriff, Edit/Bearbeiten, ...)
function Get-SPORoleName {
    param(
        [ValidateSet("Administrator", "Editor", "Contributor", "Reader")]
        [string]$RoleType
    )
    
    if ($Global:SPORoleNameCache.ContainsKey($RoleType)) {
        return $Global:SPORoleNameCache[$RoleType]
    }
    
    $fallbackNames = @{
        Administrator = "Full Control"
        Editor = "Edit"
        Contributor = "Contribute"
        Reader = "Read"
    }
    $roleName = $fallbackNames[$RoleType]
    
    try {
        $role = Get-PnPRoleDefinition -ErrorAction Stop |
            Where-Object { "$($_.RoleTypeKind)" -eq $RoleType } |
            Select-Object -First 1
        if ($role) { $roleName = $role.Name }
    } catch {
        Write-Host "Rollendefinitionen konnten nicht gelesen werden, verwende '$roleName': $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    $Global:SPORoleNameCache[$RoleType] = $roleName
    return $roleName
}

# Funktion: Session-Caches zurücksetzen (bei Verbindungswechsel)
function Reset-SPOSessionCache {
    $Global:SPOUserCache = $null
    $Global:SPORoleNameCache = @{}
    $Global:SPOListCache = @{}
    $Global:ProcessedSPOFolders = @{}
}

# Funktion: SharePoint Ordner erstellen
function New-SPOFolder {
    param(
        [string]$FolderPath
    )
    
    try {
        $relativePath = ConvertTo-SPOSiteRelativePath -Path $FolderPath
        
        # Web-Root oder Bibliotheks-Root muss nicht angelegt werden
        if ([string]::IsNullOrWhiteSpace($relativePath) -or $relativePath -notmatch '/') {
            return @{
                Success = $true
                Message = "Root-/Bibliotheksordner (wird nicht angelegt)"
                FolderPath = $relativePath
            }
        }
        
        # Bibliothek prüfen (sonst würde ein Ordner im Web statt in der Bibliothek entstehen)
        Get-SPOLibraryForPath -SiteRelativePath $relativePath | Out-Null
        
        # Resolve-PnPFolder legt fehlende Ordner rekursiv an und gibt vorhandene unverändert zurück
        Resolve-PnPFolder -SiteRelativePath $relativePath -ErrorAction Stop | Out-Null
        
        return @{
            Success = $true
            Message = "Ordner vorhanden bzw. erstellt"
            FolderPath = $relativePath
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Fehler beim Erstellen des Ordners: $($_.Exception.Message)"
            FolderPath = $FolderPath
        }
    }
}

# Funktion: SharePoint Berechtigungen setzen
function Set-SPOPermission {
    param(
        [string]$UserUPN,
        [string]$FolderPath,
        [string]$Permission,
        # $true = vorhandene (geerbte) Berechtigungen am Ziel einmalig pro Lauf entfernen
        [bool]$BreakInheritance = $true
    )
    
    try {
        $roleType = Get-SPORoleTypeForRights -Rights $Permission
        $permissionLevel = Get-SPORoleName -RoleType $roleType
        
        $relativePath = ConvertTo-SPOSiteRelativePath -Path $FolderPath
        
        # Benutzer in der Site sicherstellen (EnsureUser, idempotent - schlägt fehl, wenn der UPN in Entra ID nicht existiert)
        try {
            New-PnPUser -LoginName $UserUPN -ErrorAction Stop | Out-Null
        } catch {
            throw "Benutzer konnte nicht zur Site hinzugefügt werden: $($_.Exception.Message)"
        }
        
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
            # Web-Root: Rolle direkt am Web vergeben (ClearExisting wird hier bewusst nicht angewendet)
            Set-PnPWebPermission -User $UserUPN -AddRole $permissionLevel -ErrorAction Stop
            $message = "Site-Berechtigung '$permissionLevel' für $UserUPN gesetzt"
        }
        elseif ($relativePath -notmatch '/') {
            # Bibliotheks-Root: Vererbung mit Kopie der bestehenden Rollen unterbrechen (nie leeren)
            $list = Get-SPOLibraryForPath -SiteRelativePath $relativePath
            Set-PnPList -Identity $list -BreakRoleInheritance -CopyRoleAssignments -ErrorAction Stop | Out-Null
            Set-PnPListPermission -Identity $list -User $UserUPN -AddRole $permissionLevel -ErrorAction Stop
            $message = "Bibliotheks-Berechtigung '$permissionLevel' für $UserUPN gesetzt"
        }
        else {
            $list = Get-SPOLibraryForPath -SiteRelativePath $relativePath
            $folderKey = $relativePath.ToLowerInvariant()
            
            # -ClearExisting nur beim ERSTEN Eintrag je Ordner - sonst würde jeder weitere Benutzer
            # die zuvor gesetzten Berechtigungen wieder entfernen.
            if ($BreakInheritance -and -not $Global:ProcessedSPOFolders.ContainsKey($folderKey)) {
                Set-PnPFolderPermission -List $list -Identity $relativePath -User $UserUPN -AddRole $permissionLevel -ClearExisting -ErrorAction Stop
                $message = "Ordner-Berechtigung '$permissionLevel' für $UserUPN gesetzt (Vererbung unterbrochen, bestehende Berechtigungen ersetzt)"
            } else {
                Set-PnPFolderPermission -List $list -Identity $relativePath -User $UserUPN -AddRole $permissionLevel -ErrorAction Stop
                $message = "Ordner-Berechtigung '$permissionLevel' für $UserUPN hinzugefügt"
            }
            $Global:ProcessedSPOFolders[$folderKey] = $true
        }
        
        return @{
            Success = $true
            Message = $message
            UserUPN = $UserUPN
            FolderPath = $relativePath
            PermissionLevel = $permissionLevel
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Fehler beim Setzen der Berechtigung: $($_.Exception.Message)"
            UserUPN = $UserUPN
            FolderPath = $FolderPath
            PermissionLevel = $Permission
        }
    }
}

# Funktion: Admin als Site Administrator hinzufügen
function Add-SPOSiteAdmin {
    param(
        [string]$AdminUPN
    )
    
    try {
        # Add-PnPSiteCollectionAdmin ist idempotent; eine zusätzliche Rolle "Full Control" ist nicht nötig
        # (und würde auf lokalisierten Sites mit "Vollzugriff" fehlschlagen)
        Add-PnPSiteCollectionAdmin -Owners $AdminUPN -ErrorAction Stop
        
        return @{
            Success = $true
            Message = "$AdminUPN ist Site-Collection-Administrator"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Fehler beim Hinzufügen des Site Administrators: $($_.Exception.Message)"
        }
    }
}

# Hauptfunktion: GUI erstellen und anzeigen
function Show-MainWindow {
    # XAML laden
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Controls referenzieren
    # Tab 1: CSV Import
    $txtCSVFilePath = $window.FindName("TxtCSVFilePath")
    $btnSelectCSV = $window.FindName("BtnSelectCSV")
    $btnImportCSV = $window.FindName("BtnImportCSV")
    $dgImportedData = $window.FindName("DgImportedData")
    $txtImportCount = $window.FindName("TxtImportCount")
    $btnClearImport = $window.FindName("BtnClearImport")
    $btnValidateUsers = $window.FindName("BtnValidateUsers")
    $btnProceedToConnection = $window.FindName("BtnProceedToConnection")
    
    # Tab 2: SharePoint Verbindung
    $txtSPOSiteURL = $window.FindName("TxtSPOSiteURL")
    $txtAdminUPN = $window.FindName("TxtAdminUPN")
    $connectionStatusIndicator = $window.FindName("ConnectionStatusIndicator")
    $txtConnectionStatus = $window.FindName("TxtConnectionStatus")
    $btnAppSetup = $window.FindName("BtnAppSetup")
    $btnConnect = $window.FindName("BtnConnect")
    $txtSiteInfo = $window.FindName("TxtSiteInfo")
    $btnDisconnect = $window.FindName("BtnDisconnect")
    $btnProceedToPermissions = $window.FindName("BtnProceedToPermissions")
    
    # Tab 3: Berechtigungen anwenden
    $chkCreateFolders = $window.FindName("ChkCreateFolders")
    $chkBreakInheritance = $window.FindName("ChkBreakInheritance")
    $chkAddAdminToSite = $window.FindName("ChkAddAdminToSite")
    $btnValidateAllUsers = $window.FindName("BtnValidateAllUsers")
    $btnApplyPermissions = $window.FindName("BtnApplyPermissions")
    $dgPermissionStatus = $window.FindName("DgPermissionStatus")
    $txtPermissionCount = $window.FindName("TxtPermissionCount")
    $btnRefreshStatus = $window.FindName("BtnRefreshStatus")
    $progressBarPermissions = $window.FindName("ProgressBarPermissions")
    $txtProgressPercent = $window.FindName("TxtProgressPercent")
    $txtCurrentAction = $window.FindName("TxtCurrentAction")
    $btnExportResults = $window.FindName("BtnExportResults")
    $btnResetAll = $window.FindName("BtnResetAll")
    
    # Status Bar
    $txtStatus = $window.FindName("TxtStatus")
    $txtModuleStatus = $window.FindName("TxtModuleStatus")
    $progressBarMain = $window.FindName("ProgressBarMain")
    $mainTabControl = $window.FindName("MainTabControl")
    
    # Hilfsfunktionen für UI-Updates
    function Update-UI {
        # Dispatcher abarbeiten lassen, damit Fortschritt/Status während langer Schleifen sichtbar werden
        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }
    
    function Get-ValidationStatusText {
        param($ValidationResult)
        switch ($ValidationResult.Status) {
            "Found"     { "✅ Gültig" }
            "NotInSite" { "⚠️ Nicht in Site" }
            default     { "❌ $($ValidationResult.Message)" }
        }
    }
    
    function Update-StatusBar {
        param([string]$Message, [string]$ModuleStatus = $null)
        $txtStatus.Text = $Message
        if ($ModuleStatus) { $txtModuleStatus.Text = $ModuleStatus }
        Update-UI
    }
    
    function Update-ConnectionStatus {
        param([bool]$IsConnected, [string]$Message)
        if ($IsConnected) {
            $connectionStatusIndicator.Fill = "#27AE60"
            $txtConnectionStatus.Text = "Verbunden"
            $txtConnectionStatus.Foreground = "#27AE60"
            $btnDisconnect.IsEnabled = $true
            $btnProceedToPermissions.IsEnabled = $true
        } else {
            $connectionStatusIndicator.Fill = "#E74C3C"
            $txtConnectionStatus.Text = $Message
            $txtConnectionStatus.Foreground = "#E74C3C"
            $btnDisconnect.IsEnabled = $false
            $btnProceedToPermissions.IsEnabled = $false
        }
    }
    
    function Update-ProgressBar {
        param([int]$Value, [int]$Maximum, [string]$Action = "")
        $percentage = if ($Maximum -gt 0) { [math]::Round(($Value / $Maximum) * 100, 1) } else { 0 }
        $progressBarPermissions.Value = $percentage
        $txtProgressPercent.Text = "$percentage%"
        if ($Action) { $txtCurrentAction.Text = $Action }
    }
    
    # Initialisierung: PnP.PowerShell Modul prüfen und Einstellungen laden
    $window.Add_Loaded({
        Update-StatusBar -Message "Initialisiere Anwendung..."
        
        # Einstellungen aus Registry laden
        $settingsResult = Load-SettingsFromRegistry
        Write-Host $settingsResult.Message -ForegroundColor $(if ($settingsResult.Success) { "Green" } else { "Yellow" })
        
        # PnP.PowerShell Modul prüfen
        $moduleCheck = Test-PnPModule
        Update-StatusBar -Message "Bereit - Wählen Sie eine CSV-Datei zum Import aus..." -ModuleStatus "PnP.PowerShell: $($moduleCheck.Version)"
        
        if (-not $moduleCheck.IsAvailable) {
            [System.Windows.MessageBox]::Show($moduleCheck.Message, "Modul-Warnung", "OK", "Warning")
        }
        
        # Gespeicherte Werte in GUI laden (falls vorhanden)
        if (-not [string]::IsNullOrWhiteSpace($Global:AppRegistration.SiteURL)) {
            $txtSPOSiteURL.Text = $Global:AppRegistration.SiteURL
        }
        if (-not [string]::IsNullOrWhiteSpace($Global:AppRegistration.AdminUPN)) {
            $txtAdminUPN.Text = $Global:AppRegistration.AdminUPN
        }
        
        # App-Registrierung Setup anbieten falls nicht konfiguriert
        if (-not $settingsResult.HasClientId) {
            $result = [System.Windows.MessageBox]::Show(
                "Keine App-Registrierung gefunden. Möchten Sie jetzt das Setup durchführen?`n`nDies ist erforderlich für eine zuverlässige SharePoint-Verbindung.",
                "App-Registrierung Setup",
                "YesNo",
                "Question"
            )
            
            if ($result -eq "Yes") {
                $setupResult = Show-AppRegistrationDialog
                if ($setupResult) {
                    Update-StatusBar -Message "App-Registrierung konfiguriert - Bereit für SharePoint-Verbindung"
                }
            }
        } else {
            Update-StatusBar -Message "App-Registrierung geladen - Bereit für SharePoint-Verbindung"
        }
    })
    
    # === TAB 1: CSV IMPORT EVENT HANDLERS ===
    
    # CSV-Datei auswählen
    $btnSelectCSV.Add_Click({
        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Filter = "CSV Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
        $openDialog.Title = "CSV-Datei aus easyFPReader auswählen"
        
        if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtCSVFilePath.Text = $openDialog.FileName
            $btnImportCSV.IsEnabled = $true
            Update-StatusBar -Message "CSV-Datei ausgewählt: $($openDialog.FileName)"
        }
    })
    
    # CSV importieren
    $btnImportCSV.Add_Click({
        try {
            Update-StatusBar -Message "Importiere CSV-Datei..."
            $progressBarMain.Visibility = "Visible"
            
            $importResult = Import-CSVPermissions -FilePath $txtCSVFilePath.Text
            
            if ($importResult.Success) {
                $Global:ImportedCSVData = @($importResult.Data)
                $dgImportedData.ItemsSource = $Global:ImportedCSVData
                $dgPermissionStatus.ItemsSource = $null
                $btnApplyPermissions.IsEnabled = $false
                $txtImportCount.Text = "($($importResult.Count) Einträge)"
                
                $btnClearImport.IsEnabled = $true
                $btnValidateUsers.IsEnabled = $true
                $btnProceedToConnection.IsEnabled = $true
                
                Update-StatusBar -Message $importResult.Message
                [System.Windows.MessageBox]::Show($importResult.Message, "Import erfolgreich", "OK", "Information")
            } else {
                Update-StatusBar -Message $importResult.Message
                [System.Windows.MessageBox]::Show($importResult.Message, "Import-Fehler", "OK", "Error")
            }
        }
        catch {
            Update-StatusBar -Message "Unerwarteter Fehler beim Import: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Unerwarteter Fehler: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
        finally {
            $progressBarMain.Visibility = "Collapsed"
        }
    })
    
    # Import leeren
    $btnClearImport.Add_Click({
        $result = [System.Windows.MessageBox]::Show("Möchten Sie wirklich alle importierten Daten löschen?", "Bestätigung", "YesNo", "Question")
        if ($result -eq "Yes") {
            $Global:ImportedCSVData = @()
            $dgImportedData.ItemsSource = $null
            $txtImportCount.Text = "(0 Einträge)"
            
            $btnClearImport.IsEnabled = $false
            $btnValidateUsers.IsEnabled = $false
            $btnProceedToConnection.IsEnabled = $false
            
            Update-StatusBar -Message "Importierte Daten gelöscht"
        }
    })
    
    # Benutzer validieren (Tab 1)
    $btnValidateUsers.Add_Click({
        if ($Global:ImportedCSVData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Keine Daten zum Validieren vorhanden.", "Fehler", "OK", "Warning")
            return
        }
        
        if ($Global:SPOConnection -eq $null) {
            [System.Windows.MessageBox]::Show("Bitte stellen Sie zuerst eine SharePoint-Verbindung her.", "Fehler", "OK", "Warning")
            $mainTabControl.SelectedIndex = 1  # Wechsel zu Tab 2
            return
        }
        
        $progressBarMain.Visibility = "Visible"
        Update-StatusBar -Message "Validiere Benutzer gegen EntraID..."
        
        $validatedCount = 0
        try {
            foreach ($entry in $Global:ImportedCSVData) {
                $validationResult = Test-EntraIDUser -UserUPN $entry.EntraIDUPN
                $entry.ValidationStatus = Get-ValidationStatusText -ValidationResult $validationResult
                $validatedCount++
                if (($validatedCount % 10) -eq 0) { Update-UI }
            }
        }
        finally {
            $dgImportedData.Items.Refresh()
            $progressBarMain.Visibility = "Collapsed"
        }
        
        Update-StatusBar -Message "$validatedCount Benutzer validiert"
        [System.Windows.MessageBox]::Show("Benutzervalidierung abgeschlossen.", "Validierung", "OK", "Information")
    })
    
    # Weiter zu Verbindung
    $btnProceedToConnection.Add_Click({
        $mainTabControl.SelectedIndex = 1  # Wechsel zu Tab 2
        Update-StatusBar -Message "Konfigurieren Sie die SharePoint-Verbindung"
    })
    
    # === TAB 2: SHAREPOINT VERBINDUNG EVENT HANDLERS ===
    
    # App-Setup öffnen
    $btnAppSetup.Add_Click({
        $setupResult = Show-AppRegistrationDialog
        if ($setupResult) {
            # GUI mit neuen Werten aktualisieren
            if (-not [string]::IsNullOrWhiteSpace($Global:AppRegistration.SiteURL)) {
                $txtSPOSiteURL.Text = $Global:AppRegistration.SiteURL
            }
            if (-not [string]::IsNullOrWhiteSpace($Global:AppRegistration.AdminUPN)) {
                $txtAdminUPN.Text = $Global:AppRegistration.AdminUPN
            }
            
            Update-StatusBar -Message "App-Registrierung aktualisiert - Bereit für Verbindung"
            [System.Windows.MessageBox]::Show("App-Registrierung erfolgreich konfiguriert!", "Setup abgeschlossen", "OK", "Information")
        }
    })
    
    # SharePoint verbinden
    $btnConnect.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtSPOSiteURL.Text) -or [string]::IsNullOrWhiteSpace($txtAdminUPN.Text)) {
            [System.Windows.MessageBox]::Show("Bitte füllen Sie alle Verbindungsfelder aus.", "Fehler", "OK", "Warning")
            return
        }
        
        # Prüfen ob App-Registrierung konfiguriert ist
        if ([string]::IsNullOrWhiteSpace($Global:AppRegistration.ClientId)) {
            $result = [System.Windows.MessageBox]::Show(
                "Keine App-Registrierung konfiguriert. Möchten Sie jetzt das Setup durchführen?",
                "App-Registrierung erforderlich",
                "YesNo",
                "Question"
            )
            
            if ($result -eq "Yes") {
                $setupResult = Show-AppRegistrationDialog
                if (-not $setupResult) {
                    return
                }
            } else {
                return
            }
        }
        
        # URL-Format prüfen, bevor eine Anmeldung gestartet wird
        $urlCheck = Test-SPOConnection -SiteURL $txtSPOSiteURL.Text -AdminUPN $txtAdminUPN.Text
        if (-not $urlCheck.Success) {
            [System.Windows.MessageBox]::Show($urlCheck.Message, "Ungültige Eingabe", "OK", "Warning")
            return
        }
        
        # Aktuelle Werte speichern (falls geändert)
        if ($txtSPOSiteURL.Text -ne $Global:AppRegistration.SiteURL -or $txtAdminUPN.Text -ne $Global:AppRegistration.AdminUPN) {
            Save-SettingsToRegistry -ClientId $Global:AppRegistration.ClientId -TenantId $Global:AppRegistration.TenantId -SiteURL $txtSPOSiteURL.Text -AdminUPN $txtAdminUPN.Text
        }
        
        Update-StatusBar -Message "Verbinde mit SharePoint Online..."
        $progressBarMain.Visibility = "Visible"
        
        try {
            $connectResult = Connect-SPOSite -SiteURL $txtSPOSiteURL.Text -AdminUPN $txtAdminUPN.Text
            
            if ($connectResult.Success) {
                $Global:SPOConnection = $connectResult
                $Global:SPOSiteURL = $txtSPOSiteURL.Text
                
                Update-ConnectionStatus -IsConnected $true -Message "Verbunden"
                
                # Site-Informationen anzeigen
                $siteInfoText = @"
Site-Titel: $($connectResult.SiteInfo.Title)
URL: $($connectResult.SiteInfo.URL)
Beschreibung: $($connectResult.SiteInfo.Description)
Erstellt: $($connectResult.SiteInfo.Created)
Letzte Änderung: $($connectResult.SiteInfo.LastModified)
Besitzer: $($connectResult.SiteInfo.Owner)
Speicher verwendet: $($connectResult.SiteInfo.StorageUsed) MB
Speicher-Quota: $($connectResult.SiteInfo.StorageQuota) MB

Verbindung erfolgreich hergestellt!
"@
                $txtSiteInfo.Text = $siteInfoText
                
                Update-StatusBar -Message "SharePoint-Verbindung hergestellt"
                [System.Windows.MessageBox]::Show("Erfolgreich mit SharePoint verbunden!", "Verbindung", "OK", "Information")
            } else {
                Update-ConnectionStatus -IsConnected $false -Message "Verbindung fehlgeschlagen"
                $txtSiteInfo.Text = "Verbindungsfehler: $($connectResult.Message)"
                Update-StatusBar -Message $connectResult.Message
                [System.Windows.MessageBox]::Show($connectResult.Message, "Verbindungsfehler", "OK", "Error")
            }
        }
        catch {
            Update-ConnectionStatus -IsConnected $false -Message "Unerwarteter Fehler"
            $txtSiteInfo.Text = "Unerwarteter Fehler: $($_.Exception.Message)"
            Update-StatusBar -Message "Unerwarteter Verbindungsfehler: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Unerwarteter Fehler: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
        finally {
            $progressBarMain.Visibility = "Collapsed"
        }
    })
    
    # SharePoint trennen
    $btnDisconnect.Add_Click({
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            $Global:SPOConnection = $null
            $Global:SPOSiteURL = ""
            $Global:SPOWebUrl = ""
            Reset-SPOSessionCache
            $btnApplyPermissions.IsEnabled = $false
            
            Update-ConnectionStatus -IsConnected $false -Message "Nicht verbunden"
            $txtSiteInfo.Text = "Verbinden Sie sich mit einer SharePoint Site, um Informationen anzuzeigen..."
            
            Update-StatusBar -Message "SharePoint-Verbindung getrennt"
        }
        catch {
            Update-StatusBar -Message "Fehler beim Trennen: $($_.Exception.Message)"
        }
    })
    
    # Weiter zu Berechtigungen
    $btnProceedToPermissions.Add_Click({
        if ($Global:ImportedCSVData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte importieren Sie zuerst CSV-Daten.", "Fehler", "OK", "Warning")
            $mainTabControl.SelectedIndex = 0  # Zurück zu Tab 1
            return
        }
        
        # Daten für Tab 3 vorbereiten
        $dgPermissionStatus.ItemsSource = $Global:ImportedCSVData
        $txtPermissionCount.Text = "($($Global:ImportedCSVData.Count) Einträge)"
        
        $mainTabControl.SelectedIndex = 2  # Wechsel zu Tab 3
        Update-StatusBar -Message "Bereit für Berechtigungsanwendung"
    })
    
    # === TAB 3: BERECHTIGUNGEN ANWENDEN EVENT HANDLERS ===
    
    # Alle Benutzer validieren
    $btnValidateAllUsers.Add_Click({
        if ($Global:ImportedCSVData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Keine Daten zum Validieren vorhanden.", "Fehler", "OK", "Warning")
            return
        }
        
        if ($Global:SPOConnection -eq $null) {
            [System.Windows.MessageBox]::Show("Keine SharePoint-Verbindung vorhanden.", "Fehler", "OK", "Warning")
            return
        }
        
        Update-StatusBar -Message "Validiere alle Benutzer..."
        
        $totalUsers = $Global:ImportedCSVData.Count
        $currentUser = 0
        
        foreach ($entry in $Global:ImportedCSVData) {
            $currentUser++
            Update-ProgressBar -Value $currentUser -Maximum $totalUsers -Action "Validiere: $($entry.EntraIDUPN)"
            
            $validationResult = Test-EntraIDUser -UserUPN $entry.EntraIDUPN
            $entry.ValidationStatus = Get-ValidationStatusText -ValidationResult $validationResult
            
            Update-UI
        }
        
        $dgPermissionStatus.Items.Refresh()
        
        Update-ProgressBar -Value 0 -Maximum 100 -Action "Validierung abgeschlossen"
        $btnApplyPermissions.IsEnabled = $true
        
        Update-StatusBar -Message "Benutzervalidierung abgeschlossen"
        [System.Windows.MessageBox]::Show("Alle Benutzer wurden validiert.", "Validierung", "OK", "Information")
    })
    
    # Berechtigungen anwenden
    $btnApplyPermissions.Add_Click({
        if ($Global:ImportedCSVData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Keine Daten zum Anwenden vorhanden.", "Fehler", "OK", "Warning")
            return
        }
        
        if ($Global:SPOConnection -eq $null) {
            [System.Windows.MessageBox]::Show("Keine SharePoint-Verbindung vorhanden.", "Fehler", "OK", "Warning")
            return
        }
        
        $replaceText = if ($chkBreakInheritance.IsChecked) { "Bestehende Berechtigungen an Zielordnern werden ERSETZT." } else { "Bestehende Berechtigungen an Zielordnern werden kopiert und ergänzt." }
        $result = [System.Windows.MessageBox]::Show("Möchten Sie die Berechtigungen jetzt auf die SharePoint Site anwenden?`n`n$replaceText`n`nDieser Vorgang kann nicht automatisch rückgängig gemacht werden.", "Berechtigungen anwenden", "YesNo", "Warning")
        if ($result -ne "Yes") { return }
        
        Update-StatusBar -Message "Wende Berechtigungen an..."
        
        # Übergeordnete Ordner immer vor Unterordnern verarbeiten: Das erstmalige Unterbrechen der Vererbung
        # (BreakRoleInheritance mit clearSubscopes) würde sonst bereits gesetzte Unterordner-Berechtigungen löschen.
        # Innerhalb einer Ebene bleibt die CSV-Reihenfolge erhalten (Index als zweiter Sortierschlüssel).
        $sortIndex = 0
        $orderedEntries = @($Global:ImportedCSVData | ForEach-Object {
            [PSCustomObject]@{
                Entry = $_
                Depth = @(("$($_.SharePointPath)".Trim().TrimEnd('/', '\')) -split '[/\\]').Count
                Index = $sortIndex++
            }
        } | Sort-Object Depth, Index | ForEach-Object { $_.Entry })
        
        # Pro Lauf neu: welche Ordner bereits bereinigt wurden (ClearExisting nur einmal je Ordner)
        $Global:ProcessedSPOFolders = @{}
        $btnApplyPermissions.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        $totalEntries = $Global:ImportedCSVData.Count
        $currentEntry = 0
        $successCount = 0
        $errorCount = 0
        
        # Admin als Site Administrator hinzufügen (falls gewünscht)
        if ($chkAddAdminToSite.IsChecked) {
            Update-ProgressBar -Value 0 -Maximum $totalEntries -Action "Füge Admin als Site Administrator hinzu..."
            Update-UI
            $adminResult = Add-SPOSiteAdmin -AdminUPN $txtAdminUPN.Text
            if (-not $adminResult.Success) {
                [System.Windows.MessageBox]::Show("Warnung: $($adminResult.Message)", "Admin-Setup", "OK", "Warning")
            }
        }
        
        try {
            foreach ($entry in $orderedEntries) {
                $currentEntry++
                Update-ProgressBar -Value $currentEntry -Maximum $totalEntries -Action "Verarbeite: $($entry.EntraIDUPN)"
                Update-UI
            
                try {
                    # Ordner erstellen (falls gewünscht)
                    if ($chkCreateFolders.IsChecked -and -not [string]::IsNullOrWhiteSpace($entry.SharePointPath)) {
                        $folderResult = New-SPOFolder -FolderPath $entry.SharePointPath
                        if (-not $folderResult.Success) {
                            $entry.ApplyStatus = "❌ Ordner-Fehler: $($folderResult.Message)"
                            $errorCount++
                            $dgPermissionStatus.Items.Refresh()
                            continue
                        }
                    }
                
                    # Berechtigung setzen
                    $permissionResult = Set-SPOPermission -UserUPN $entry.EntraIDUPN -FolderPath $entry.SharePointPath -Permission $entry.Permission -BreakInheritance ([bool]$chkBreakInheritance.IsChecked)
                
                    if ($permissionResult.Success) {
                        $entry.ApplyStatus = "✅ Erfolgreich"
                        $successCount++
                    } else {
                        $entry.ApplyStatus = "❌ Fehler: $($permissionResult.Message)"
                        $errorCount++
                    }
                }
                catch {
                    $entry.ApplyStatus = "❌ Unerwarteter Fehler: $($_.Exception.Message)"
                    $errorCount++
                }
            
                # UI aktualisieren
                $dgPermissionStatus.Items.Refresh()
            }
        }
        finally {
            $window.Cursor = $null
            $btnApplyPermissions.IsEnabled = $true
        }
        
        Update-ProgressBar -Value $totalEntries -Maximum $totalEntries -Action "Berechtigungsanwendung abgeschlossen"
        $btnExportResults.IsEnabled = $true
        
        $summaryMessage = "Berechtigungsanwendung abgeschlossen:`n`n✅ Erfolgreich: $successCount`n❌ Fehler: $errorCount`n📊 Gesamt: $totalEntries"
        Update-StatusBar -Message "Berechtigungen angewendet: $successCount erfolgreich, $errorCount Fehler"
        [System.Windows.MessageBox]::Show($summaryMessage, "Anwendung abgeschlossen", "OK", "Information")
    })
    
    # Status aktualisieren
    $btnRefreshStatus.Add_Click({
        $dgPermissionStatus.ItemsSource = $null
        $dgPermissionStatus.ItemsSource = $Global:ImportedCSVData
        Update-StatusBar -Message "Status aktualisiert"
    })
    
    # Ergebnisse exportieren
    $btnExportResults.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Dateien (*.csv)|*.csv"
        $saveDialog.FileName = "SPO_Permissions_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $Global:ImportedCSVData | Select-Object OnPremUser, EntraIDUPN, Permission, SharePointPath, ValidationStatus, ApplyStatus |
                    Export-Csv -LiteralPath $saveDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                
                Update-StatusBar -Message "Ergebnisse exportiert: $($saveDialog.FileName)"
                [System.Windows.MessageBox]::Show("Ergebnisse erfolgreich exportiert!", "Export", "OK", "Information")
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler beim Export: $($_.Exception.Message)", "Export-Fehler", "OK", "Error")
            }
        }
    })
    
    # Alles zurücksetzen
    $btnResetAll.Add_Click({
        $result = [System.Windows.MessageBox]::Show("Möchten Sie wirklich alle Daten und Verbindungen zurücksetzen?", "Zurücksetzen", "YesNo", "Question")
        if ($result -eq "Yes") {
            # Verbindung trennen
            try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
            
            # Globale Variablen zurücksetzen
            $Global:ImportedCSVData = @()
            $Global:SPOConnection = $null
            $Global:SPOSiteURL = ""
            $Global:SPOWebUrl = ""
            Reset-SPOSessionCache
            
            # UI zurücksetzen
            $txtCSVFilePath.Text = "Wählen Sie eine CSV-Datei aus dem easyFPReader..."
            $dgImportedData.ItemsSource = $null
            $dgPermissionStatus.ItemsSource = $null
            $txtImportCount.Text = "(0 Einträge)"
            $txtPermissionCount.Text = "(0 Einträge)"
            
            Update-ConnectionStatus -IsConnected $false -Message "Nicht verbunden"
            $txtSiteInfo.Text = "Verbinden Sie sich mit einer SharePoint Site, um Informationen anzuzeigen..."
            
            # Buttons zurücksetzen
            $btnImportCSV.IsEnabled = $false
            $btnClearImport.IsEnabled = $false
            $btnValidateUsers.IsEnabled = $false
            $btnProceedToConnection.IsEnabled = $false
            $btnApplyPermissions.IsEnabled = $false
            $btnExportResults.IsEnabled = $false
            
            Update-ProgressBar -Value 0 -Maximum 100 -Action "Bereit für neue Berechtigungsanwendung..."
            
            $mainTabControl.SelectedIndex = 0  # Zurück zu Tab 1
            Update-StatusBar -Message "Alle Daten zurückgesetzt - Bereit für neuen Import"
        }
    })
    
    # Fenster anzeigen
    $window.ShowDialog() | Out-Null
}

# Script starten
try {
    Write-Host "Starte easyFolder SharePoint Online Permissions Manager V0.0.2 (BETA)..." -ForegroundColor Green
    Show-MainWindow
}
catch {
    Write-Error "Fehler beim Starten der Anwendung: $($_.Exception.Message)"
    Read-Host "Drücken Sie Enter zum Beenden"
}