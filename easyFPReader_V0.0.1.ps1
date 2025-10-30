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

# SIG # Begin signature block
# MIIoiQYJKoZIhvcNAQcCoIIoejCCKHYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCmWEk8kO6D2exW
# zMIF2ZYUZFQ8/IxPcTLHPt0znyeWWaCCILswggXJMIIEsaADAgECAhAbtY8lKt8j
# AEkoya49fu0nMA0GCSqGSIb3DQEBDAUAMH4xCzAJBgNVBAYTAlBMMSIwIAYDVQQK
# ExlVbml6ZXRvIFRlY2hub2xvZ2llcyBTLkEuMScwJQYDVQQLEx5DZXJ0dW0gQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkxIjAgBgNVBAMTGUNlcnR1bSBUcnVzdGVkIE5l
# dHdvcmsgQ0EwHhcNMjEwNTMxMDY0MzA2WhcNMjkwOTE3MDY0MzA2WjCBgDELMAkG
# A1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAl
# BgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMb
# Q2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEAvfl4+ObVgAxknYYblmRnPyI6HnUBfe/7XGeMycxca6mR5rlC
# 5SBLm9qbe7mZXdmbgEvXhEArJ9PoujC7Pgkap0mV7ytAJMKXx6fumyXvqAoAl4Va
# qp3cKcniNQfrcE1K1sGzVrihQTib0fsxf4/gX+GxPw+OFklg1waNGPmqJhCrKtPQ
# 0WeNG0a+RzDVLnLRxWPa52N5RH5LYySJhi40PylMUosqp8DikSiJucBb+R3Z5yet
# /5oCl8HGUJKbAiy9qbk0WQq/hEr/3/6zn+vZnuCYI+yma3cWKtvMrTscpIfcRnNe
# GWJoRVfkkIJCu0LW8GHgwaM9ZqNd9BjuiMmNF0UpmTJ1AjHuKSbIawLmtWJFfzcV
# WiNoidQ+3k4nsPBADLxNF8tNorMe0AZa3faTz1d1mfX6hhpneLO/lv403L3nUlbl
# s+V1e9dBkQXcXWnjlQ1DufyDljmVe2yAWk8TcsbXfSl6RLpSpCrVQUYJIP4ioLZb
# MI28iQzV13D4h1L92u+sUS4Hs07+0AnacO+Y+lbmbdu1V0vc5SwlFcieLnhO+Nqc
# noYsylfzGuXIkosagpZ6w7xQEmnYDlpGizrrJvojybawgb5CAKT41v4wLsfSRvbl
# jnX98sy50IdbzAYQYLuDNbdeZ95H7JlI8aShFf6tjGKOOVVPORa5sWOd/7cCAwEA
# AaOCAT4wggE6MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFLahVDkCw6A/joq8
# +tT4HKbROg79MB8GA1UdIwQYMBaAFAh2zcsH/yT2xc3tu5C84oQ3RnX3MA4GA1Ud
# DwEB/wQEAwIBBjAvBgNVHR8EKDAmMCSgIqAghh5odHRwOi8vY3JsLmNlcnR1bS5w
# bC9jdG5jYS5jcmwwawYIKwYBBQUHAQEEXzBdMCgGCCsGAQUFBzABhhxodHRwOi8v
# c3ViY2Eub2NzcC1jZXJ0dW0uY29tMDEGCCsGAQUFBzAChiVodHRwOi8vcmVwb3Np
# dG9yeS5jZXJ0dW0ucGwvY3RuY2EuY2VyMDkGA1UdIAQyMDAwLgYEVR0gADAmMCQG
# CCsGAQUFBwIBFhhodHRwOi8vd3d3LmNlcnR1bS5wbC9DUFMwDQYJKoZIhvcNAQEM
# BQADggEBAFHCoVgWIhCL/IYx1MIy01z4S6Ivaj5N+KsIHu3V6PrnCA3st8YeDrJ1
# BXqxC/rXdGoABh+kzqrya33YEcARCNQOTWHFOqj6seHjmOriY/1B9ZN9DbxdkjuR
# mmW60F9MvkyNaAMQFtXx0ASKhTP5N+dbLiZpQjy6zbzUeulNndrnQ/tjUoCFBMQl
# lVXwfqefAcVbKPjgzoZwpic7Ofs4LphTZSJ1Ldf23SIikZbr3WjtP6MZl9M7JYjs
# NhI9qX7OAo0FmpKnJ25FspxihjcNpDOO16hO0EoXQ0zF8ads0h5YbBRRfopUofbv
# n3l6XYGaFpAP4bvxSgD5+d2+7arszgowggaDMIIEa6ADAgECAhEAnpwE9lWotKcC
# bUmMbHiNqjANBgkqhkiG9w0BAQwFADBWMQswCQYDVQQGEwJQTDEhMB8GA1UEChMY
# QXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0dW0gVGltZXN0
# YW1waW5nIDIwMjEgQ0EwHhcNMjUwMTA5MDg0MDQzWhcNMzYwMTA3MDg0MDQzWjBQ
# MQswCQYDVQQGEwJQTDEhMB8GA1UECgwYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MR4wHAYDVQQDDBVDZXJ0dW0gVGltZXN0YW1wIDIwMjUwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDHKV9n+Kwr3ZBF5UCLWOQ/NdbblAvQeGMjfCi/bibT
# 71hPkwKV4UvQt1MuOwoaUCYtsLhw8jrmOmoz2HoHKKzEpiS3A1rA3ssXUZMnSrbi
# iVpDj+5MtnbXSVEJKbccuHbmwcjl39N4W72zccoC/neKAuwO1DJ+9SO+YkHncRiV
# 95idWhxRAcDYv47hc9GEFZtTFxQXLbrL4N7N90BqLle3ayznzccEPQ+E6H6p00zE
# 9HUp++3bZTF4PfyPRnKCLc5ezAzEqqbbU5F/nujx69T1mm02jltlFXnTMF1vlake
# QXWYpGIjtrR7WP7tIMZnk78nrYSfeAp8le+/W/5+qr7tqQZufW9invsRTcfk7P+m
# nKjJLuSbwqgxelvCBryz9r51bT0561aR2c+joFygqW7n4FPCnMLOj40X4ot7wP2u
# 8kLRDVHbhsHq5SGLqr8DbFq14ws2ALS3tYa2GGiA7wX79rS5oDMnSY/xmJO5cupu
# SvqpylzO7jzcLOwWiqCrq05AXp51SRrj9xRt8KdZWpDdWhWmE8MFiFtmQ0AqODLJ
# Bn1hQAx3FvD/pte6pE1Bil0BOVC2Snbeq/3NylDwvDdAg/0CZRJsQIaydHswJwyY
# BlYUDyaQK2yUS57hobnYx/vStMvTB96ii4jGV3UkZh3GvwdDCsZkbJXaU8ATF/z6
# DwIDAQABo4IBUDCCAUwwdQYIKwYBBQUHAQEEaTBnMDsGCCsGAQUFBzAChi9odHRw
# Oi8vc3ViY2EucmVwb3NpdG9yeS5jZXJ0dW0ucGwvY3RzY2EyMDIxLmNlcjAoBggr
# BgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAfBgNVHSMEGDAW
# gBS+VAIvv0Bsc0POrAklTp5DRBru4DAMBgNVHRMBAf8EAjAAMDkGA1UdHwQyMDAw
# LqAsoCqGKGh0dHA6Ly9zdWJjYS5jcmwuY2VydHVtLnBsL2N0c2NhMjAyMS5jcmww
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMCIGA1UdIAQb
# MBkwCAYGZ4EMAQQCMA0GCyqEaAGG9ncCBQELMB0GA1UdDgQWBBSBjAagKFP8AD/b
# fp5KwR8i7LISiTANBgkqhkiG9w0BAQwFAAOCAgEAmQ8ZDBvrBUPnaL87AYc4Jlmf
# H1ZP5yt65MtzYu8fbmsL3d3cvYs+Enbtfu9f2wMehzSyved3Rc59a04O8NN7plw4
# PXg71wfSE4MRFM1EuqL63zq9uTjm/9tA73r1aCdWmkprKp0aLoZolUN0qGcvr9+Q
# G8VIJVMcuSqFeEvRrLEKK2xVkMSdTTbDhseUjI4vN+BrXm5z45EA3aDpSiZQuoNd
# 4RFnDzddbgfcCQPaY2UyXqzNBjnuz6AyHnFzKtNlCevkMBgh4dIDt/0DGGDOaTEA
# WZtUEqK5AlHd0PBnd40Lnog4UATU3Bt6GHfeDmWEHFTjHKsmn9Q8wiGj906bVgL8
# 35tfEH9EgYDklqrOUxWxDf1cOA7ds/r8pIc2vjLQ9tOSkm9WXVbnTeLG3Q57frTg
# CvTObd/qf3UzE97nTNOU7vOMZEo41AgmhuEbGsyQIDM/V6fJQX1RnzzJNoqfTTkU
# zUoP2tlNHnNsjFo2YV+5yZcoaawmNWmR7TywUXG2/vFgJaG0bfEoodeeXp7A4I4H
# aDDpfRa7ypgJEPeTwHuBRJpj9N+1xtri+6BzHPwsAAvUJm58PGoVsteHAXwvpg4N
# VgvUk3BKbl7xFulWU1KHqH/sk7T0CFBQ5ohuKPmFf1oqAP4AO9a3Yg2wBMwEg1zP
# Oh6xbUXskzs9iSa9yGwwgga5MIIEoaADAgECAhEAmaOACiZVO2Wr3G6EprPqOTAN
# BgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8g
# VGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9u
# IEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAy
# MB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIxOFowVjELMAkGA1UEBhMCUEwx
# ITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2Vy
# dHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5cTbq96y34vuTmflN4mSAfgLKT
# vggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7VS5+djSoMcbvIKck6+hI1shs
# ylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1PH9ud0IF+njvMk2xqbNTIPsnW
# tw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOouu9Tj1yHIohzuC8KNqfcYf7Z4
# /iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv8aGUsRdaCtVD2bSlbfsq7Biq
# ljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtMLK+Wo837Q4QOZgYqVWQ4x6cM
# 7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9lDV2nT8mFSkcSkAExzd4prHw
# YjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/JHuurfTI5XDYO962WZayx7AC
# Ff5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkKbWpQ5boufUnq1UiYPIAHlezf
# 4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADaCi2JSplKShBSND36E/ENVv8u
# rPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAWd18Jx5n858JSqPECAwEAAaOC
# AVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFN10XUwA23ufoHTKsW73
# PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB
# /wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAwBgNVHR8EKTAnMCWgI6Ahhh9o
# dHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAo
# BggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEF
# BQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYD
# VR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVt
# LnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhYD+WPUCiaU58Q7EP89DttyZqG
# Yn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUStJl490L94C9LGF3vjzzH8Jq3
# iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChDUyuQy6rGDxLUUAsO0eqeLNhL
# Vsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiRsWrhWM2f8pXdd3x2mbJCKKtl
# 2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7bWRLDm0CdY9rNLqyA3ahe8Wlx
# VWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mATwZWwSD+B7eMcZNhpn8zJ+6MT
# yE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3/bFAEloMU+vUBfSouCReZwSL
# o8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESYkOh1/w1tVxTpV2Na3PR7nxYV
# lPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR+x+zPF/2DaGgK2W1eEJfo2qy
# rBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C+xN4YaNjt2ywzOr+tKyEVAot
# nyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qwt4HOUBCrW602NCmvO1nm+/80
# nLy5r0AZvCQxaQ4wgga5MIIEoaADAgECAhEA5/9pxzs1zkuRJth0fGilhzANBgkq
# hkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVj
# aG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1
# dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMB4X
# DTIxMDUxOTA1MzIwN1oXDTM2MDUxODA1MzIwN1owVjELMAkGA1UEBhMCUEwxITAf
# BgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVt
# IFRpbWVzdGFtcGluZyAyMDIxIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA6RIfBDXtuV16xaaVQb6KZX9Od9FtJXXTZo7b+GEof3+3g0ChWiKnO7R4
# +6MfrvLyLCWZa6GpFHjEt4t0/GiUQvnkLOBRdBqr5DOvlmTvJJs2X8ZmWgWJjC7P
# BZLYBWAs8sJl3kNXxBMX5XntjqWx1ZOuuXl0R4x+zGGSMzZ45dpvB8vLpQfZkfMC
# /1tL9KYyjU+htLH68dZJPtzhqLBVG+8ljZ1ZFilOKksS79epCeqFSeAUm2eMTGpO
# iS3gfLM6yvb8Bg6bxg5yglDGC9zbr4sB9ceIGRtCQF1N8dqTgM/dSViiUgJkcv5d
# LNJeWxGCqJYPgzKlYZTgDXfGIeZpEFmjBLwURP5ABsyKoFocMzdjrCiFbTvJn+bD
# 1kq78qZUgAQGGtd6zGJ88H4NPJ5Y2R4IargiWAmv8RyvWnHr/VA+2PrrK9eXe5q7
# M88YRdSTq9TKbqdnITUgZcjjm4ZUjteq8K331a4P0s2in0p3UubMEYa/G5w6jSWP
# UzchGLwWKYBfeSu6dIOC4LkeAPvmdZxSB1lWOb9HzVWZoM8Q/blaP4LWt6JxjkI9
# yQsYGMdCqwl7uMnPUIlcExS1mzXRxUowQref/EPaS7kYVaHHQrp4XB7nTEtQhkP0
# Z9Puz/n8zIFnUSnxDof4Yy650PAXSYmK2TcbyDoTNmmt8xAxzcMCAwEAAaOCAVUw
# ggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFL5UAi+/QGxzQ86sCSVOnkNE
# Gu7gMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB/wQE
# AwIBBjATBgNVHSUEDDAKBggrBgEFBQcDCDAwBgNVHR8EKTAnMCWgI6Ahhh9odHRw
# Oi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAoBggr
# BgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEFBQcw
# AoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYDVR0g
# BDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVtLnBs
# L0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAuJNZd8lMFf2UBwigp3qgLPBBk58BFCS3
# Q6aJDf3TISoytK0eal/JyCB88aUEd0wMNiEcNVMbK9j5Yht2whaknUE1G32k6uld
# 7wcxHmw67vUBY6pSp8QhdodY4SzRRaZWzyYlviUpyU4dXyhKhHSncYJfa1U75cXx
# Ce3sTp9uTBm3f8Bj8LkpjMUSVTtMJ6oEu5JqCYzRfc6nnoRUgwz/GVZFoOBGdrSE
# tDN7mZgcka/tS5MI47fALVvN5lZ2U8k7Dm/hTX8CWOw0uBZloZEW4HB0Xra3qE4q
# zzq/6M8gyoU/DE0k3+i7bYOrOk/7tPJg1sOhytOGUQ30PbG++0FfJioDuOFhj99b
# 151SqFlSaRQYz74y/P2XJP+cF19oqozmi0rRTkfyEJIvhIZ+M5XIFZttmVQgTxfp
# fJwMFFEoQrSrklOxpmSygppsUDJEoliC05vBLVQ+gMZyYaKvBJ4YxBMlKH5ZHkRd
# loRYlUDplk8GUa+OCMVhpDSQurU6K1ua5dmZftnvSSz2H96UrQDzA6DyiI1V3ejV
# tvn2azVAXg6NnjmuRZ+wa7Pxy0H3+V4K4rOTHlG3VYA6xfLsTunCz72T6Ot4+tkr
# DYOeaU1pPX1CBfYj6EW2+ELq46GP8KCNUQDirWLU4nOmgCat7vN0SD6RlwUiSsMe
# CiQDmZwgwrUwggbpMIIE0aADAgECAhBiOsZKIV2oSfsf25d4iu6HMA0GCSqGSIb3
# DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28gRGF0YSBTeXN0
# ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBDQTAe
# Fw0yNTA3MzExMTM4MDhaFw0yNjA3MzExMTM4MDdaMIGOMQswCQYDVQQGEwJERTEb
# MBkGA1UECAwSQmFkZW4tV8O8cnR0ZW1iZXJnMRQwEgYDVQQHDAtCYWllcnNicm9u
# bjEeMBwGA1UECgwVT3BlbiBTb3VyY2UgRGV2ZWxvcGVyMSwwKgYDVQQDDCNPcGVu
# IFNvdXJjZSBEZXZlbG9wZXIsIEhlcHAgQW5kcmVhczCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAOt2txKXx2UtfBNIw2kVihIAcgPkK3lp7np/qE0evLq2
# J/L5kx8m6dUY4WrrcXPSn1+W2/PVs/XBFV4fDfwczZnQ/hYzc8Ot5YxPKLx6hZxK
# C5v8LjNIZ3SRJvMbOpjzWoQH7MLIIj64n8mou+V0CMk8UElmU2d0nxBQyau1njQP
# CLvlfInu4tDndyp3P87V5bIdWw6MkZFhWDkILTYInYicYEkut5dN9hT02t/3rXu2
# 30DEZ6S1OQtm9loo8wzvwjRoVX3IxnfpCHGW8Z9ie9I9naMAOG2YpvpoUbLG3fL/
# B6JVNNR1mm/AYaqVMtAXJpRlqvbIZyepcG0YGB+kOQLdoQCWlIp3a14Z4kg6bU9C
# U1KNR4ueA+SqLNu0QGtgBAdTfqoWvyiaeyEogstBHglrZ39y/RW8OOa50pSleSRx
# SXiGW+yH+Ps5yrOopTQpKHy0kRincuJpYXgxGdGxxKHwuVJHKXL0nWScEku0C38p
# M9sYanIKncuF0Ed7RvyNqmPP5pt+p/0ZG+zLNu/Rce0LE5FjAIRtW2hFxmYMyohk
# afzyjCCCG0p2KFFT23CoUfXx59nCU+lyWx/iyDMV4sqrcvmZdPZF7lkaIb5B4PYP
# vFFE7enApz4Niycj1gPUFlx4qTcXHIbFLJDp0ry6MYelX+SiMHV7yDH/rnWXm5d3
# AgMBAAGjggF4MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1UdHwQ2MDQwMqAwoC6GLGh0
# dHA6Ly9jY3NjYTIwMjEuY3JsLmNlcnR1bS5wbC9jY3NjYTIwMjEuY3JsMHMGCCsG
# AQUFBwEBBGcwZTAsBggrBgEFBQcwAYYgaHR0cDovL2Njc2NhMjAyMS5vY3NwLWNl
# cnR1bS5jb20wNQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBvc2l0b3J5LmNlcnR1bS5w
# bC9jY3NjYTIwMjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA23ufoHTKsW73PMAywHDN
# MB0GA1UdDgQWBBQYl6R41hwxInb9JVvqbCTp9ILCcTBLBgNVHSAERDBCMAgGBmeB
# DAEEATA2BgsqhGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIBFhlodHRwczovL3d3dy5j
# ZXJ0dW0ucGwvQ1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIH
# gDANBgkqhkiG9w0BAQsFAAOCAgEAQ4guyo7zysB7MHMBOVKKY72rdY5hrlxPci8u
# 1RgBZ9ZDGFzhnUM7iIivieAeAYLVxP922V3ag9sDVNR+mzCmu1pWCgZyBbNXykue
# KJwOfE8VdpmC/F7637i8a7Pyq6qPbcfvLSqiXtVrT4NX4NIvODW3kIqf4nGwd0h3
# 1tuJVHLkdpGmT0q4TW0gAxnNoQ+lO8uNzCrtOBk+4e1/3CZXSDnjR8SUsHrHdhnm
# qkAnYb40vf69dfDR148tToUj872yYeBUEGUsQUDgJ6HSkMVpLQz/Nb3xy9qkY33M
# 7CBWKuBVwEcbGig/yj7CABhIrY1XwRddYQhEyozUS4mXNqXydAD6Ylt143qrECD2
# s3MDQBgP2sbRHdhVgzr9+n1iztXkPHpIlnnXPkZrt89E5iGL+1PtjETrhTkr7nxj
# yMFjrbmJ8W/XglwopUTCGfopDFPlzaoFf5rH/v3uzS24yb6+dwQrvCwFA9Y9ZHy2
# ITJx7/Ll6AxWt7Lz9JCJ5xRyYeRUHs6ycB8EuMPAKyGpzdGtjWv2rkTXbkIYUjkl
# FTpquXJBc/kO5L+Quu0a0uKn4ea16SkABy052XHQqd87cSJg3rGxsagi0IAfxGM6
# 08oupufSS/q9mpQPgkDuMJ8/zdre0st8OduAoG131W+XJ7mm0gIuh2zNmSIet5RD
# oa8THmwxggckMIIHIAIBATBqMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3Nl
# Y28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25p
# bmcgMjAyMSBDQQIQYjrGSiFdqEn7H9uXeIruhzANBglghkgBZQMEAgEFAKCBhDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEi
# BCChyxgG5+tYONYnmXZnFjcVp+zlgffRcFmqb+JdLwglRzANBgkqhkiG9w0BAQEF
# AASCAgBEM8C8gWCSbKPwkS3Ok8/st3aOQZG868oKSjPs3+lMMM5tnmlVgQgS6m/g
# syR+x/x2bCaQg0NT3xhA+KqYjc1Kg+Z9Fg7rtwjYrFJT3omy+llaMQSjr3H7Sp5w
# PdbwD6pZJDvVTpDrII5H8jkuok3alq5+QW6nZk3XNb/3pD0/WJNNYwbgFHJR/pz/
# pyaW7NSUDAEHPfAFY3K6EdoRequgy2NPEfDMhmXzDVhbzT8ZJHb3iOwfzG8iN1ga
# zOnWPCatpfIgZb1Iv9j33T3kzslkQIsiamiGlmK1bBJgrnuDf8B3+HX/F3g7IjAe
# L7qnsjfcSXCaApEsy5zwrnthSqBUZsKmp2pCd0fSBktJGssP5X4q3MAAThPBg42e
# rOYX3a/oVfpKcuh4vZXWlL9xuZiyA2a1Rq7BV37MomXrVjW/N7efGi9DciifI1+6
# G26SdLVSOAsAF/704xrYE9eGpv8jrit54q1asjU+xFdDzCsDJ/dHopK1TZGhPe1t
# pPpUS1wIbjfmmkZ7p1M+xkDxFZy6FUqBA2WEcLAJWycmYPCU0BUfnncXBSBCGPcM
# k5Dbw+HRhEziZZMXNPh4q0kJEQuaSX8RWPBZ+gkftP67bSB/t5KWD6oxQA72GFOh
# kURdCIucylBxaK75LeDgysg8QmFBjY4FhItaP3O0GM1B/TnDhqGCBAQwggQABgkq
# hkiG9w0BCQYxggPxMIID7QIBATBrMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3Rh
# bXBpbmcgMjAyMSBDQQIRAJ6cBPZVqLSnAm1JjGx4jaowDQYJYIZIAWUDBAICBQCg
# ggFXMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcN
# MjUxMDI1MTc1ODQ3WjA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCDPodw1ne0rw8uJ
# D6Iw5dr3e1QPGm4rI93PF1ThjPqg1TA/BgkqhkiG9w0BCQQxMgQwn0a2doCB4skq
# qEy+x4wzvVxp2cEAOypClYuHjYcERYd76kvSdCKQCkRyAF7PFiK0MIGgBgsqhkiG
# 9w0BCRACDDGBkDCBjTCBijCBhwQUwyW4mxf8xQJgYc4rcXtFB92camowbzBapFgw
# VjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5B
# LjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBAhEAnpwE9lWo
# tKcCbUmMbHiNqjANBgkqhkiG9w0BAQEFAASCAgBwE9W7CppXsTA13kdHs7/hSmH5
# oervX7ypuSenTyW4kcFWcLh8BttAAXgXe9WVadXSEURJbOlYJAbsGBRFtCXQv3SU
# W4JeSrFxPwJk3OqJ3AB0kNqPN0I94T6b57GMTEEoZPzYSEYiZ/xD89+J2uYxn+8v
# hfzbxZUB7rzvTjbpG+XORDNuKuxUg9ARC7zGm8DKeuFtAboQLt4sFKsE+d364jsv
# ImNs0GyD9j2ECMzjqqByLoTuRhZUbdfhwcORTYC3Zxw6pBWXJ2jEsrOGzY6OdKIv
# tV/ByDnG8i+Qiwr3bOn9cUr1N5/B55EWoS+n3XGL1T87Q3QMh2ltS+RyAYjtRlpA
# 8OsgBhjAZCZojWkawoMHVvMbi68nqlTpclUuaPxJJgDtqu9+YE/13tBy1XSEyk1z
# n2Y5AdwA7gQHcDm5mE1YGdfiSeFqn26u8z+4Rqr4ShU3VZIR/wTZ9zIFvfZeCv38
# 2e2FuASY+YkOyiBTOCYBTY5DQrorHxMKYIFCWZAHG81jb+LpbixZRHFlcLGbulhr
# mptQUJ4bOYBaaxBlfMemSqEzHbbnC/dwCZkl07nl7unZyRCzRP2fWBBP30g2/YLY
# 8Ulau1IHxkUJfZVxf08QcGwB/Jkc0Jtax+c9XTLk1gSFBxBlIrE762eS1J286hw5
# 8Bo/OTWMwcYaptJPsg==
# SIG # End signature block
