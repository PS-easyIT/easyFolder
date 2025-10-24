#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Easy Folder Permissions Manager (easyFPFM) V0.0.1
    
.DESCRIPTION
    PowerShell Script mit WPF GUI zur Verwaltung von Ordnerberechtigungen
    - Anzeige aktueller Berechtigungen
    - Hinzufügen von Benutzern mit spezifischen Rechten
    - Unterstützung für Hauptordner, Unterordner oder beide
    - Automatische Report-Erstellung bei Änderungen
    
.AUTHOR
    PhinIT Development
    
.VERSION
    0.0.1
    
.NOTES
    WICHTIG FÜR EXE-KONVERTIERUNG:
    - Alle Debug-Meldungen werden AUSSCHLIESSLICH in die Log-Datei geschrieben
    - Keine Write-Host, Write-Error oder Write-Warning Ausgaben in der Shell!
    - Debug-Logfile: %USERPROFILE%\Documents\FolderPermissions_Reports\Debug.log
    - Dies ermöglicht eine saubere GUI-EXE ohne Console-Fenster
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Globale Variablen
$Global:ReportsPath = "$env:USERPROFILE\Documents\FolderPermissions_Reports"
$Global:CurrentFolder = ""
$Global:PermissionsData = @()

# Logfile-Pfad für Debug-Informationen
$Global:LogFilePath = "$env:USERPROFILE\Documents\FolderPermissions_Reports\Debug.log"

# Logging-Funktion (nur in Datei, nicht in Shell)
function Write-LogEntry {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        Add-Content -Path $Global:LogFilePath -Value $logMessage -ErrorAction SilentlyContinue
    }
    catch {
        # Fehler beim Schreiben ins Log werden ignoriert
    }
}

# Report-Ordner erstellen falls nicht vorhanden
if (-not (Test-Path $Global:ReportsPath)) {
    New-Item -Path $Global:ReportsPath -ItemType Directory -Force | Out-Null
}

# Erweiterte Funktionen
function Remove-FolderPermission {
    param(
        [string]$FolderPath,
        [string]$Username,
        [string]$ApplyTo
    )
    
    try {
        $identity = [System.Security.Principal.NTAccount]$Username
        
        switch ($ApplyTo) {
            "MainOnly" {
                $acl = Get-Acl -Path $FolderPath
                $acl.PurgeAccessRules($identity)
                Set-Acl -Path $FolderPath -AclObject $acl
                Create-PermissionReport -Action "REMOVE" -Path $FolderPath -User $Username -Rights "ALL" -ApplyTo $ApplyTo
            }
            "SubOnly" {
                $subfolders = Get-ChildItem -Path $FolderPath -Directory -Recurse
                foreach ($subfolder in $subfolders) {
                    $acl = Get-Acl -Path $subfolder.FullName
                    $acl.PurgeAccessRules($identity)
                    Set-Acl -Path $subfolder.FullName -AclObject $acl
                }
                Create-PermissionReport -Action "REMOVE" -Path $FolderPath -User $Username -Rights "ALL" -ApplyTo $ApplyTo
            }
            "Both" {
                # Hauptordner
                $acl = Get-Acl -Path $FolderPath
                $acl.PurgeAccessRules($identity)
                Set-Acl -Path $FolderPath -AclObject $acl
                
                # Unterordner
                $subfolders = Get-ChildItem -Path $FolderPath -Directory -Recurse
                foreach ($subfolder in $subfolders) {
                    $acl = Get-Acl -Path $subfolder.FullName
                    $acl.PurgeAccessRules($identity)
                    Set-Acl -Path $subfolder.FullName -AclObject $acl
                }
                Create-PermissionReport -Action "REMOVE" -Path $FolderPath -User $Username -Rights "ALL" -ApplyTo $ApplyTo
            }
        }
        
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Entfernen der Berechtigung: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

function Backup-FolderPermissions {
    param(
        [string]$FolderPath
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $backupFile = "$Global:ReportsPath\Backup_$($FolderPath.Replace(':', '').Replace('\', '_'))_$timestamp.xml"
        
        $permissions = Get-FolderPermissions -FolderPath $FolderPath -IncludeSubfolders $true
        $permissions | Export-Clixml -Path $backupFile
        
        return $backupFile
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Erstellen des Backups: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
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
    
    if (-not [string]::IsNullOrEmpty($UserFilter)) {
        $filtered = $filtered | Where-Object { $_.User -like "*$UserFilter*" }
    }
    
    if (-not [string]::IsNullOrEmpty($RightsFilter)) {
        $filtered = $filtered | Where-Object { $_.Rights -like "*$RightsFilter*" }
    }
    
    if ($ShowInheritedOnly) {
        $filtered = $filtered | Where-Object { $_.Inherited -eq $true }
    }
    
    return $filtered
}

function Get-SystemUsers {
    try {
        $users = @()
        
        # Lokale Benutzer
        $localUsers = Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -eq $true }
        foreach ($user in $localUsers) {
            $users += [PSCustomObject]@{
                Name = $user.Name
                FullName = if ($user.FullName) { "$($user.Name) ($($user.FullName))" } else { $user.Name }
                Type = "Lokaler Benutzer"
                SID = $user.SID
            }
        }
        
        # Lokale Gruppen
        $localGroups = Get-LocalGroup -ErrorAction SilentlyContinue
        foreach ($group in $localGroups) {
            $users += [PSCustomObject]@{
                Name = $group.Name
                FullName = if ($group.Description) { "$($group.Name) ($($group.Description))" } else { $group.Name }
                Type = "Lokale Gruppe"
                SID = $group.SID
            }
        }
        
        # Bekannte System-Accounts hinzufügen
        $systemAccounts = @(
            @{Name="Everyone"; FullName="Everyone (Jeder)"; Type="System-Gruppe"},
            @{Name="Authenticated Users"; FullName="Authenticated Users (Authentifizierte Benutzer)"; Type="System-Gruppe"},
            @{Name="SYSTEM"; FullName="SYSTEM (System)"; Type="System-Account"},
            @{Name="Administrators"; FullName="Administrators (Administratoren)"; Type="Lokale Gruppe"},
            @{Name="Users"; FullName="Users (Benutzer)"; Type="Lokale Gruppe"}
        )
        
        foreach ($account in $systemAccounts) {
            if (-not ($users | Where-Object { $_.Name -eq $account.Name })) {
                $users += [PSCustomObject]@{
                    Name = $account.Name
                    FullName = $account.FullName
                    Type = $account.Type
                    SID = ""
                }
            }
        }
        
        return $users | Sort-Object Type, Name
    }
    catch {
        Write-LogEntry "Fehler beim Laden der Benutzer: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Remove-SelectedPermissions {
    param(
        [array]$SelectedPermissions,
        [string]$FolderPath
    )
    
    $successCount = 0
    $errorCount = 0
    $errors = @()
    
    foreach ($permission in $SelectedPermissions) {
        try {
            $identity = [System.Security.Principal.NTAccount]$permission.User
            $acl = Get-Acl -Path $permission.Folder
            $acl.PurgeAccessRules($identity)
            Set-Acl -Path $permission.Folder -AclObject $acl
            $successCount++
        }
        catch {
            $errorCount++
            $errors += "$($permission.User) auf $($permission.Folder): $($_.Exception.Message)"
        }
    }
    
    # Report erstellen
    if ($successCount -gt 0) {
        Create-PermissionReport -Action "BULK_REMOVE" -Path $FolderPath -User "Multiple ($successCount)" -Rights "Various" -ApplyTo "Selected"
    }
    
    return @{
        Success = $successCount
        Errors = $errorCount
        ErrorDetails = $errors
    }
}

function Export-PermissionsToHtml {
    param(
        [array]$Permissions,
        [string]$FilePath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
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
        </tr>
"@
    
    foreach ($perm in $Permissions) {
        $inheritedClass = if ($perm.Inherited) { "inherited" } else { "" }
        $accessClass = if ($perm.AccessType -eq "Allow") { "allow" } else { "deny" }
        $inheritedText = if ($perm.Inherited) { "Ja" } else { "Nein" }
        
        $html += @"
        <tr class="$inheritedClass">
            <td>$($perm.Folder)</td>
            <td>$($perm.User)</td>
            <td>$($perm.Rights)</td>
            <td class="$accessClass">$($perm.AccessType)</td>
            <td>$inheritedText</td>
        </tr>
"@
    }
    
    $html += @"
    </table>
</body>
</html>
"@
    
    $html | Out-File -FilePath $FilePath -Encoding UTF8
}

# Hauptfunktionen
function Get-FolderPermissions {
    param(
        [string]$FolderPath,
        [bool]$IncludeSubfolders = $false
    )
    
    $permissions = @()
    
    try {
        if ($IncludeSubfolders) {
            $folders = Get-ChildItem -Path $FolderPath -Directory -Recurse -ErrorAction SilentlyContinue
            $folders = @($FolderPath) + $folders.FullName
        } else {
            $folders = @($FolderPath)
        }
        
        foreach ($folder in $folders) {
            $acl = Get-Acl -Path $folder -ErrorAction SilentlyContinue
            if ($acl) {
                foreach ($access in $acl.Access) {
                    $permissions += [PSCustomObject]@{
                        Folder = $folder
                        User = $access.IdentityReference.Value
                        Rights = $access.FileSystemRights
                        AccessType = $access.AccessControlType
                        Inherited = $access.IsInherited
                    }
                }
            }
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Lesen der Berechtigungen: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
    
    return $permissions
}

function Add-FolderPermission {
    param(
        [string]$FolderPath,
        [string]$Username,
        [string]$Rights,
        [string]$ApplyTo
    )
    
    try {
        $rightsEnum = [System.Security.AccessControl.FileSystemRights]$Rights
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Username, $rightsEnum, "Allow")
        
        switch ($ApplyTo) {
            "MainOnly" {
                $acl = Get-Acl -Path $FolderPath
                $acl.SetAccessRule($accessRule)
                Set-Acl -Path $FolderPath -AclObject $acl
                Create-PermissionReport -Action "ADD" -Path $FolderPath -User $Username -Rights $Rights -ApplyTo $ApplyTo
            }
            "SubOnly" {
                $subfolders = Get-ChildItem -Path $FolderPath -Directory -Recurse
                foreach ($subfolder in $subfolders) {
                    $acl = Get-Acl -Path $subfolder.FullName
                    $acl.SetAccessRule($accessRule)
                    Set-Acl -Path $subfolder.FullName -AclObject $acl
                }
                Create-PermissionReport -Action "ADD" -Path $FolderPath -User $Username -Rights $Rights -ApplyTo $ApplyTo
            }
            "Both" {
                # Hauptordner
                $acl = Get-Acl -Path $FolderPath
                $acl.SetAccessRule($accessRule)
                Set-Acl -Path $FolderPath -AclObject $acl
                
                # Unterordner
                $subfolders = Get-ChildItem -Path $FolderPath -Directory -Recurse
                foreach ($subfolder in $subfolders) {
                    $acl = Get-Acl -Path $subfolder.FullName
                    $acl.SetAccessRule($accessRule)
                    Set-Acl -Path $subfolder.FullName -AclObject $acl
                }
                Create-PermissionReport -Action "ADD" -Path $FolderPath -User $Username -Rights $Rights -ApplyTo $ApplyTo
            }
        }
        
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Hinzufügen der Berechtigung: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

function Create-PermissionReport {
    param(
        [string]$Action,
        [string]$Path,
        [string]$User,
        [string]$Rights,
        [string]$ApplyTo
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $reportFile = "$Global:ReportsPath\PermissionReport_$timestamp.txt"
    
    $report = @"
=== FOLDER PERMISSIONS REPORT ===
Datum/Zeit: $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")
Aktion: $Action
Ordner: $Path
Benutzer: $User
Rechte: $Rights
Angewendet auf: $ApplyTo

=== AKTUELLE BERECHTIGUNGEN NACH ÄNDERUNG ===
"@
    
    # Aktuelle Berechtigungen hinzufügen
    $currentPerms = Get-FolderPermissions -FolderPath $Path -IncludeSubfolders ($ApplyTo -eq "Both" -or $ApplyTo -eq "SubOnly")
    foreach ($perm in $currentPerms) {
        $report += "`n$($perm.Folder) | $($perm.User) | $($perm.Rights) | $($perm.AccessType)"
    }
    
    $report | Out-File -FilePath $reportFile -Encoding UTF8
}

# WPF GUI XAML - Windows 11 Design
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Easy Folder Permissions Manager V0.0.1" Height="900" Width="1700"
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
                <Button Name="btnBackup" Grid.Column="3" Content="Backup" Width="90" Height="32" Background="#107C10"/>
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
                <DataGrid Name="dgPermissions" Grid.Row="0" AutoGenerateColumns="True" IsReadOnly="True" 
                         GridLinesVisibility="Horizontal" HeadersVisibility="All" MinHeight="300"/>
                <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
                    <TextBlock Name="txtPermissionCount" Text="Einträge: 0" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="#666"/>
                    <Button Name="btnRemoveSelected" Content="Entfernen" Width="120" Height="28" Background="#D13438" IsEnabled="False"/>
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
                        <RadioButton Name="rbMainOnly" Content="Nur Hauptordner" Margin="0,0,0,4" IsChecked="True"/>
                        <RadioButton Name="rbSubOnly" Content="Nur Unterordner" Margin="0,0,0,4"/>
                        <RadioButton Name="rbBoth" Content="Hauptordner + Unterordner" Margin="0,0,0,4"/>
                    </StackPanel>
                    
                    <Button Name="btnAddPermission" Content="Hinzufügen" Height="32" Margin="0,0,0,5"/>
                    <Button Name="btnRemovePermission" Content="Entfernen" Height="32" Background="#D13438"/>
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
                    <TextBlock Text="Easy Folder Permissions Manager V0.0.1" FontWeight="SemiBold" Foreground="#333" Margin="0,0,15,0"/>
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
    exit 1
}

# Controls referenzieren
$txtFolderPath = $window.FindName("txtFolderPath")
$btnBrowse = $window.FindName("btnBrowse")
$btnLoadPermissions = $window.FindName("btnLoadPermissions")
$btnBackup = $window.FindName("btnBackup")
$chkIncludeSubfolders = $window.FindName("chkIncludeSubfolders")
$chkShowInheritedOnly = $window.FindName("chkShowInheritedOnly")
$txtUserFilter = $window.FindName("txtUserFilter")
$txtRightsFilter = $window.FindName("txtRightsFilter")
$btnApplyFilter = $window.FindName("btnApplyFilter")
$dgPermissions = $window.FindName("dgPermissions")
$txtPermissionCount = $window.FindName("txtPermissionCount")
$btnRemoveSelected = $window.FindName("btnRemoveSelected")
$txtUsername = $window.FindName("txtUsername")
$btnBrowseUser = $window.FindName("btnBrowseUser")
$cmbRights = $window.FindName("cmbRights")
$rbMainOnly = $window.FindName("rbMainOnly")
$rbSubOnly = $window.FindName("rbSubOnly")
$rbBoth = $window.FindName("rbBoth")
$btnAddPermission = $window.FindName("btnAddPermission")
$btnRemovePermission = $window.FindName("btnRemovePermission")
$txtStatus = $window.FindName("txtStatus")
$btnOpenReports = $window.FindName("btnOpenReports")
$btnExportCsv = $window.FindName("btnExportCsv")
$btnRefresh = $window.FindName("btnRefresh")
$btnExit = $window.FindName("btnExit")
$txtWebsiteLink = $window.FindName("txtWebsiteLink")
$txtSystemInfo = $window.FindName("txtSystemInfo")
$txtCurrentTime = $window.FindName("txtCurrentTime")

# Validierung der Controls
$controls = @{
    "txtFolderPath" = $txtFolderPath
    "btnBrowse" = $btnBrowse
    "btnLoadPermissions" = $btnLoadPermissions
    "btnBackup" = $btnBackup
    "chkIncludeSubfolders" = $chkIncludeSubfolders
    "chkShowInheritedOnly" = $chkShowInheritedOnly
    "txtUserFilter" = $txtUserFilter
    "txtRightsFilter" = $txtRightsFilter
    "btnApplyFilter" = $btnApplyFilter
    "dgPermissions" = $dgPermissions
    "txtPermissionCount" = $txtPermissionCount
    "btnRemoveSelected" = $btnRemoveSelected
    "txtUsername" = $txtUsername
    "btnBrowseUser" = $btnBrowseUser
    "cmbRights" = $cmbRights
    "rbMainOnly" = $rbMainOnly
    "rbSubOnly" = $rbSubOnly
    "rbBoth" = $rbBoth
    "btnAddPermission" = $btnAddPermission
    "btnRemovePermission" = $btnRemovePermission
    "txtStatus" = $txtStatus
    "btnOpenReports" = $btnOpenReports
    "btnExportCsv" = $btnExportCsv
    "btnRefresh" = $btnRefresh
    "btnExit" = $btnExit
    "txtWebsiteLink" = $txtWebsiteLink
    "txtSystemInfo" = $txtSystemInfo
    "txtCurrentTime" = $txtCurrentTime
}

foreach ($controlName in $controls.Keys) {
    if ($controls[$controlName] -eq $null) {
        Write-LogEntry "Control '$controlName' konnte nicht gefunden werden!" "ERROR"
        exit 1
    }
}

# Event Handlers
$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Ordner für Berechtigungsverwaltung auswählen"
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolderPath.Text = $folderBrowser.SelectedPath
        $Global:CurrentFolder = $folderBrowser.SelectedPath
        $txtStatus.Text = "Ordner ausgewählt: $($folderBrowser.SelectedPath)"
    }
})

$btnLoadPermissions.Add_Click({
    if ([string]::IsNullOrEmpty($txtFolderPath.Text)) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst einen Ordner aus.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    if (-not (Test-Path $txtFolderPath.Text)) {
        [System.Windows.MessageBox]::Show("Der angegebene Ordner existiert nicht.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    $Global:CurrentFolder = $txtFolderPath.Text
    $includeSubfolders = $chkIncludeSubfolders.IsChecked
    
    $txtStatus.Text = "Lade Berechtigungen..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Global:PermissionsData = Get-FolderPermissions -FolderPath $Global:CurrentFolder -IncludeSubfolders $includeSubfolders
        
        # Filter anwenden falls gesetzt
        $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
        
        $dgPermissions.ItemsSource = $filteredData
        $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
        $txtStatus.Text = "Berechtigungen erfolgreich geladen"
    }
    finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

$btnAddPermission.Add_Click({
    if ([string]::IsNullOrEmpty($Global:CurrentFolder)) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst einen Ordner aus und laden Sie die Berechtigungen.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    if ([string]::IsNullOrEmpty($txtUsername.Text)) {
        [System.Windows.MessageBox]::Show("Bitte geben Sie einen Benutzernamen ein.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    if ($cmbRights.SelectedItem -eq $null) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie die Rechte aus.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $username = $txtUsername.Text
    $rights = $cmbRights.SelectedItem.Tag
    
    $applyTo = "MainOnly"
    if ($rbSubOnly.IsChecked) { $applyTo = "SubOnly" }
    elseif ($rbBoth.IsChecked) { $applyTo = "Both" }
    
    $txtStatus.Text = "Füge Berechtigung hinzu..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $success = Add-FolderPermission -FolderPath $Global:CurrentFolder -Username $username -Rights $rights -ApplyTo $applyTo
        
        if ($success) {
            $txtStatus.Text = "Berechtigung erfolgreich hinzugefügt für: $username ($rights) - $applyTo"
            $txtUsername.Text = ""
            $cmbRights.SelectedIndex = -1
            
            # Berechtigungen neu laden
            $includeSubfolders = $chkIncludeSubfolders.IsChecked
            $Global:PermissionsData = Get-FolderPermissions -FolderPath $Global:CurrentFolder -IncludeSubfolders $includeSubfolders
            $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
            $dgPermissions.ItemsSource = $filteredData
            $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
        }
    }
    finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

$btnOpenReports.Add_Click({
    if (Test-Path $Global:ReportsPath) {
        Start-Process explorer.exe -ArgumentList $Global:ReportsPath
    } else {
        [System.Windows.MessageBox]::Show("Reports-Ordner wurde noch nicht erstellt.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
})

$btnRefresh.Add_Click({
    if (-not [string]::IsNullOrEmpty($Global:CurrentFolder)) {
        $txtStatus.Text = "Aktualisiere Berechtigungen..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        try {
            $includeSubfolders = $chkIncludeSubfolders.IsChecked
            $Global:PermissionsData = Get-FolderPermissions -FolderPath $Global:CurrentFolder -IncludeSubfolders $includeSubfolders
            $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
            $dgPermissions.ItemsSource = $filteredData
            $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
            $txtStatus.Text = "Berechtigungen erfolgreich aktualisiert"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    }
})

$btnExit.Add_Click({
    $window.Close()
})

# Neue Event Handler
$btnBackup.Add_Click({
    if ([string]::IsNullOrEmpty($Global:CurrentFolder)) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst einen Ordner aus.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $txtStatus.Text = "Erstelle Backup..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $backupFile = Backup-FolderPermissions -FolderPath $Global:CurrentFolder
        if ($backupFile) {
            $txtStatus.Text = "Backup erfolgreich erstellt: $(Split-Path $backupFile -Leaf)"
        }
    }
    finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

$btnApplyFilter.Add_Click({
    if ($Global:PermissionsData.Count -gt 0) {
        $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
        $dgPermissions.ItemsSource = $filteredData
        $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
        $txtStatus.Text = "Filter angewendet"
    }
})

$btnRemovePermission.Add_Click({
    if ([string]::IsNullOrEmpty($Global:CurrentFolder)) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst einen Ordner aus.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    if ([string]::IsNullOrEmpty($txtUsername.Text)) {
        [System.Windows.MessageBox]::Show("Bitte geben Sie einen Benutzernamen ein.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $username = $txtUsername.Text
    $applyTo = "MainOnly"
    if ($rbSubOnly.IsChecked) { $applyTo = "SubOnly" }
    elseif ($rbBoth.IsChecked) { $applyTo = "Both" }
    
    $result = [System.Windows.MessageBox]::Show("Möchten Sie alle Berechtigungen für '$username' wirklich entfernen?", "Bestätigung", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $txtStatus.Text = "Entferne Berechtigung..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        try {
            $success = Remove-FolderPermission -FolderPath $Global:CurrentFolder -Username $username -ApplyTo $applyTo
            
            if ($success) {
                $txtStatus.Text = "Berechtigung erfolgreich entfernt für: $username - $applyTo"
                $txtUsername.Text = ""
                
                # Berechtigungen neu laden
                $includeSubfolders = $chkIncludeSubfolders.IsChecked
                $Global:PermissionsData = Get-FolderPermissions -FolderPath $Global:CurrentFolder -IncludeSubfolders $includeSubfolders
                $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
                $dgPermissions.ItemsSource = $filteredData
                $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
            }
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    }
})

$btnExportCsv.Add_Click({
    if ($Global:PermissionsData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
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
        # CSV Export
        $saveDialog.Filter = "CSV Dateien (*.csv)|*.csv"
        $saveDialog.FileName = "FolderPermissions_$timestamp.csv"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $Global:PermissionsData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                $txtStatus.Text = "CSV erfolgreich exportiert: $(Split-Path $saveDialog.FileName -Leaf)"
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler beim CSV-Export: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    } else {
        # HTML Export
        $saveDialog.Filter = "HTML Dateien (*.html)|*.html"
        $saveDialog.FileName = "FolderPermissions_Report_$timestamp.html"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                Export-PermissionsToHtml -Permissions $Global:PermissionsData -FilePath $saveDialog.FileName
                $txtStatus.Text = "HTML-Report erfolgreich exportiert: $(Split-Path $saveDialog.FileName -Leaf)"
                
                # Optional: HTML-Datei öffnen
                $openResult = [System.Windows.MessageBox]::Show("HTML-Report wurde erstellt. Möchten Sie ihn jetzt öffnen?", "Report erstellt", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                if ($openResult -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $saveDialog.FileName
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Fehler beim HTML-Export: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    }
})

$btnBrowseUser.Add_Click({
    try {
        $txtStatus.Text = "Lade Benutzer und Gruppen..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        $users = Get-SystemUsers
        
        if ($users.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Keine Benutzer oder Gruppen gefunden.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }
        
        # Einfacher Auswahldialog
        $selectedUser = $users | Out-GridView -Title "Benutzer oder Gruppe auswählen" -OutputMode Single
        
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
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# DataGrid Selection Event
$dgPermissions.Add_SelectionChanged({
    $btnRemoveSelected.IsEnabled = $dgPermissions.SelectedItems.Count -gt 0
})

# Ausgewählte Berechtigungen entfernen
$btnRemoveSelected.Add_Click({
    if ($dgPermissions.SelectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens eine Berechtigung aus.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $selectedCount = $dgPermissions.SelectedItems.Count
    $result = [System.Windows.MessageBox]::Show("Möchten Sie $selectedCount ausgewählte Berechtigung(en) wirklich entfernen?", "Bestätigung", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $txtStatus.Text = "Entferne $selectedCount Berechtigung(en)..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        try {
            $selectedPermissions = @()
            foreach ($item in $dgPermissions.SelectedItems) {
                $selectedPermissions += $item
            }
            
            $removeResult = Remove-SelectedPermissions -SelectedPermissions $selectedPermissions -FolderPath $Global:CurrentFolder
            
            if ($removeResult.Success -gt 0) {
                $txtStatus.Text = "$($removeResult.Success) Berechtigung(en) erfolgreich entfernt"
                if ($removeResult.Errors -gt 0) {
                    $txtStatus.Text += ", $($removeResult.Errors) Fehler"
                    $errorDetails = $removeResult.ErrorDetails -join "`n"
                    [System.Windows.MessageBox]::Show("Einige Berechtigungen konnten nicht entfernt werden:`n`n$errorDetails", "Teilweise erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                }
                
                # Berechtigungen neu laden
                $includeSubfolders = $chkIncludeSubfolders.IsChecked
                $Global:PermissionsData = Get-FolderPermissions -FolderPath $Global:CurrentFolder -IncludeSubfolders $includeSubfolders
                $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
                $dgPermissions.ItemsSource = $filteredData
                $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
            } else {
                $txtStatus.Text = "Fehler beim Entfernen der Berechtigungen"
            }
        }
        catch {
            Write-LogEntry "Fehler beim Entfernen der Berechtigungen: $($_.Exception.Message)" "ERROR"
            $txtStatus.Text = "Fehler beim Entfernen der Berechtigungen"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    }
})

# Filter bei Eingabe automatisch anwenden
$txtUserFilter.Add_TextChanged({
    if ($Global:PermissionsData.Count -gt 0) {
        $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
        $dgPermissions.ItemsSource = $filteredData
        $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
    }
})

$txtRightsFilter.Add_TextChanged({
    if ($Global:PermissionsData.Count -gt 0) {
        $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
        $dgPermissions.ItemsSource = $filteredData
        $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
    }
})

$chkShowInheritedOnly.Add_Checked({
    if ($Global:PermissionsData.Count -gt 0) {
        $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
        $dgPermissions.ItemsSource = $filteredData
        $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
    }
})

$chkShowInheritedOnly.Add_Unchecked({
    if ($Global:PermissionsData.Count -gt 0) {
        $filteredData = Get-FilteredPermissions -Permissions $Global:PermissionsData -UserFilter $txtUserFilter.Text -RightsFilter $txtRightsFilter.Text -ShowInheritedOnly $chkShowInheritedOnly.IsChecked
        $dgPermissions.ItemsSource = $filteredData
        $txtPermissionCount.Text = "Einträge: $($filteredData.Count) von $($Global:PermissionsData.Count)"
    }
})

# Footer-Funktionalität
$txtWebsiteLink.Add_MouseLeftButtonUp({
    try {
        Start-Process "https://www.phinit.de"
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Öffnen der Website: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
})

# System-Info und Zeit aktualisieren
function Update-FooterInfo {
    try {
        # System-Info
        $psVersion = $PSVersionTable.PSVersion.ToString()
        $osInfo = (Get-CimInstance Win32_OperatingSystem).Caption
        $txtSystemInfo.Text = "PowerShell $psVersion | $osInfo"
        
        # Aktuelle Zeit
        $txtCurrentTime.Text = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    }
    catch {
        $txtSystemInfo.Text = "Windows PowerShell"
        $txtCurrentTime.Text = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    }
}

# Timer für Zeit-Update
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({
    $txtCurrentTime.Text = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
})
$timer.Start()

# Standardwerte setzen
$cmbRights.SelectedIndex = 0
$txtStatus.Text = "Bereit. Wählen Sie einen Ordner aus und laden Sie die Berechtigungen."
$txtPermissionCount.Text = "Einträge: 0"

# Footer initialisieren
Update-FooterInfo

# Window anzeigen
$window.ShowDialog() | Out-Nul
# SIG # Begin signature block
# MIIRcAYJKoZIhvcNAQcCoIIRYTCCEV0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD4XEFKh/19yuOR
# N2JC0IIhp4aPfcPkVnEkbqg7Jmo/9aCCDaowgga5MIIEoaADAgECAhEAmaOACiZV
# O2Wr3G6EprPqOTANBgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNV
# BAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBD
# ZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQg
# TmV0d29yayBDQSAyMB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIxOFowVjEL
# MAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEk
# MCIGA1UEAxMbQ2VydHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5cTbq96y34
# vuTmflN4mSAfgLKTvggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7VS5+djSo
# McbvIKck6+hI1shsylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1PH9ud0IF+
# njvMk2xqbNTIPsnWtw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOouu9Tj1yHI
# ohzuC8KNqfcYf7Z4/iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv8aGUsRda
# CtVD2bSlbfsq7BiqljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtMLK+Wo837
# Q4QOZgYqVWQ4x6cM7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9lDV2nT8m
# FSkcSkAExzd4prHwYjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/JHuurfTI
# 5XDYO962WZayx7ACFf5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkKbWpQ5bou
# fUnq1UiYPIAHlezf4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADaCi2JSplK
# ShBSND36E/ENVv8urPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAWd18Jx5n8
# 58JSqPECAwEAAaOCAVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFN10
# XUwA23ufoHTKsW73PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbR
# Og79MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAwBgNVHR8E
# KTAnMCWgI6Ahhh9odHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsG
# AQUFBwEBBGAwXjAoBggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVt
# LmNvbTAyBggrBgEFBQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0
# bmNhMi5jZXIwOQYDVR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6
# Ly93d3cuY2VydHVtLnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhYD+WPUCia
# U58Q7EP89DttyZqGYn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUStJl490L9
# 4C9LGF3vjzzH8Jq3iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChDUyuQy6rG
# DxLUUAsO0eqeLNhLVsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiRsWrhWM2f
# 8pXdd3x2mbJCKKtl2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7bWRLDm0Cd
# Y9rNLqyA3ahe8WlxVWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mATwZWwSD+B
# 7eMcZNhpn8zJ+6MTyE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3/bFAEloM
# U+vUBfSouCReZwSLo8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESYkOh1/w1t
# VxTpV2Na3PR7nxYVlPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR+x+zPF/2
# DaGgK2W1eEJfo2qyrBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C+xN4YaNj
# t2ywzOr+tKyEVAotnyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qwt4HOUBCr
# W602NCmvO1nm+/80nLy5r0AZvCQxaQ4wggbpMIIE0aADAgECAhBiOsZKIV2oSfsf
# 25d4iu6HMA0GCSqGSIb3DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNp
# Z25pbmcgMjAyMSBDQTAeFw0yNTA3MzExMTM4MDhaFw0yNjA3MzExMTM4MDdaMIGO
# MQswCQYDVQQGEwJERTEbMBkGA1UECAwSQmFkZW4tV8O8cnR0ZW1iZXJnMRQwEgYD
# VQQHDAtCYWllcnNicm9ubjEeMBwGA1UECgwVT3BlbiBTb3VyY2UgRGV2ZWxvcGVy
# MSwwKgYDVQQDDCNPcGVuIFNvdXJjZSBEZXZlbG9wZXIsIEhlcHAgQW5kcmVhczCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOt2txKXx2UtfBNIw2kVihIA
# cgPkK3lp7np/qE0evLq2J/L5kx8m6dUY4WrrcXPSn1+W2/PVs/XBFV4fDfwczZnQ
# /hYzc8Ot5YxPKLx6hZxKC5v8LjNIZ3SRJvMbOpjzWoQH7MLIIj64n8mou+V0CMk8
# UElmU2d0nxBQyau1njQPCLvlfInu4tDndyp3P87V5bIdWw6MkZFhWDkILTYInYic
# YEkut5dN9hT02t/3rXu230DEZ6S1OQtm9loo8wzvwjRoVX3IxnfpCHGW8Z9ie9I9
# naMAOG2YpvpoUbLG3fL/B6JVNNR1mm/AYaqVMtAXJpRlqvbIZyepcG0YGB+kOQLd
# oQCWlIp3a14Z4kg6bU9CU1KNR4ueA+SqLNu0QGtgBAdTfqoWvyiaeyEogstBHglr
# Z39y/RW8OOa50pSleSRxSXiGW+yH+Ps5yrOopTQpKHy0kRincuJpYXgxGdGxxKHw
# uVJHKXL0nWScEku0C38pM9sYanIKncuF0Ed7RvyNqmPP5pt+p/0ZG+zLNu/Rce0L
# E5FjAIRtW2hFxmYMyohkafzyjCCCG0p2KFFT23CoUfXx59nCU+lyWx/iyDMV4sqr
# cvmZdPZF7lkaIb5B4PYPvFFE7enApz4Niycj1gPUFlx4qTcXHIbFLJDp0ry6MYel
# X+SiMHV7yDH/rnWXm5d3AgMBAAGjggF4MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1Ud
# HwQ2MDQwMqAwoC6GLGh0dHA6Ly9jY3NjYTIwMjEuY3JsLmNlcnR1bS5wbC9jY3Nj
# YTIwMjEuY3JsMHMGCCsGAQUFBwEBBGcwZTAsBggrBgEFBQcwAYYgaHR0cDovL2Nj
# c2NhMjAyMS5vY3NwLWNlcnR1bS5jb20wNQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBv
# c2l0b3J5LmNlcnR1bS5wbC9jY3NjYTIwMjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA
# 23ufoHTKsW73PMAywHDNMB0GA1UdDgQWBBQYl6R41hwxInb9JVvqbCTp9ILCcTBL
# BgNVHSAERDBCMAgGBmeBDAEEATA2BgsqhGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIB
# FhlodHRwczovL3d3dy5jZXJ0dW0ucGwvQ1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAQ4guyo7zysB7MHMB
# OVKKY72rdY5hrlxPci8u1RgBZ9ZDGFzhnUM7iIivieAeAYLVxP922V3ag9sDVNR+
# mzCmu1pWCgZyBbNXykueKJwOfE8VdpmC/F7637i8a7Pyq6qPbcfvLSqiXtVrT4NX
# 4NIvODW3kIqf4nGwd0h31tuJVHLkdpGmT0q4TW0gAxnNoQ+lO8uNzCrtOBk+4e1/
# 3CZXSDnjR8SUsHrHdhnmqkAnYb40vf69dfDR148tToUj872yYeBUEGUsQUDgJ6HS
# kMVpLQz/Nb3xy9qkY33M7CBWKuBVwEcbGig/yj7CABhIrY1XwRddYQhEyozUS4mX
# NqXydAD6Ylt143qrECD2s3MDQBgP2sbRHdhVgzr9+n1iztXkPHpIlnnXPkZrt89E
# 5iGL+1PtjETrhTkr7nxjyMFjrbmJ8W/XglwopUTCGfopDFPlzaoFf5rH/v3uzS24
# yb6+dwQrvCwFA9Y9ZHy2ITJx7/Ll6AxWt7Lz9JCJ5xRyYeRUHs6ycB8EuMPAKyGp
# zdGtjWv2rkTXbkIYUjklFTpquXJBc/kO5L+Quu0a0uKn4ea16SkABy052XHQqd87
# cSJg3rGxsagi0IAfxGM608oupufSS/q9mpQPgkDuMJ8/zdre0st8OduAoG131W+X
# J7mm0gIuh2zNmSIet5RDoa8THmwxggMcMIIDGAIBATBqMFYxCzAJBgNVBAYTAlBM
# MSEwHwYDVQQKExhBc3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0Nl
# cnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBDQQIQYjrGSiFdqEn7H9uXeIruhzANBglg
# hkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MC8GCSqGSIb3DQEJBDEiBCA+I4mISfmDW6sMfSQthbsJ5ta6niOXCYv9XwD3hY2o
# nzANBgkqhkiG9w0BAQEFAASCAgABSkqMyCuzBB9AXV/mNvjtvCGmx3iq8f8WXXz1
# 9g7G7MTh8qqQRy2lGAd+Xe70vXjwQc6Yw48ncWJw3q3anes5IAqQufDkx6fKRZZ4
# Q0x+hVicuPxxpmNz+MGTOP/uhaAIDsJzcH/MEdzkRXXjrxp69EVktBaDki1OyXJk
# s051EqXlyEdAHfYtdxGJadQ7yUJX2bBBdZo4rMp0OrnQh1UYs4CJgFsep+4AKVm8
# r/C0BfYT5Mh5dXqKrh/5b27PfT1Jh2ZI7NOu7kOPLwW1PxKa4vMs3krwyHdXUjAg
# QTLAN5LiKJ8Z06rE8XsvvDxUEvdO9qt9pPy0n2sBxFlzmtNjk2gJ2TyZkLM6YzFq
# GwbEnV9Oa3BM6qm5o4HMOHB7Vyk5BuTOVJQJn1phosZdYo4baOdFAS2zIbySUob0
# ZC5QRxBbSEucAUKVNMAj02vpYX/yXMhzRGY/Oq3xT2oJuTNpBBrCb7YasnMpWgI7
# ANIlp2VSlWzKRr9EhEKeJ+3nTdBar3hwF3obESfytxsbT3Ae0Y+GiW6BRdYlVxUy
# r6GbkO56RlmdjE4JvMYlxa2pMcAlLAsE2Q3cWxvKZiff5uGFZabzwZzl+6FpT32f
# DyaZ9oAITNPRHlxOKm6wiVYY/cgxgZvkbalx10LiKRDJZQYaZ25qNUMO16AEejKM
# ir6xqA==
# SIG # End signature block
