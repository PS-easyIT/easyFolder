#Requires -Version 5.1
#Requires -RunAsAdministrator

# PS2EXE OPTIMIERUNGEN - KRITISCH FÜR GUI-EXE
$ProgressPreference = 'SilentlyContinue'  # Verhindert blinkende Progress-Fenster bei -noConsole

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
    WICHTIG FÜR EXE-KONVERTIERUNG MIT PS2EXE:
    - Alle Debug-Meldungen werden AUSSCHLIESSLICH in die Log-Datei geschrieben
    - Keine Write-Host, Write-Error oder Write-Warning Ausgaben in der Shell!
    - Debug-Logfile: %USERPROFILE%\Documents\FolderPermissions_Reports\Debug.log
    - $ProgressPreference = 'SilentlyContinue' am Anfang (verhindert Progress-Fenster)
    - Visual Styles VOR GUI-Erstellung aktiviert
    - ShowDialog() mit [VOID] versehen (verhindert "False" MessageBox)
    - UTF8-Encoding für deutsche Umlaute
    
    KOMPILIERUNG:
    ps2exe .\easyFPManager_V0.0.1.ps1 .\easyFPManager.exe -noConsole -STA -x64 `
           -iconFile .\icon.ico -title "Easy Folder Permissions Manager" `
           -version "0.0.1.0" -company "PhinIT" -copyright "© 2025 Andreas Hepp"
#>

# PS2EXE: Visual Styles MÜSSEN VOR allen GUI-Objekten aktiviert werden!
[System.Windows.Forms.Application]::EnableVisualStyles()

# Assemblies laden (Out-Null verhindert MessageBoxen bei -noConsole)
Add-Type -AssemblyName PresentationFramework | Out-Null
Add-Type -AssemblyName PresentationCore | Out-Null
Add-Type -AssemblyName WindowsBase | Out-Null
Add-Type -AssemblyName System.Windows.Forms | Out-Null

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
    
    # PS2EXE: ShowDialog() Rückgabewert in Variable speichern (verhindert MessageBox)
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
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
        
        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
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
        
        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
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

# Window anzeigen (VOID verhindert "False" MessageBox bei -noConsole)
[VOID]$window.ShowDialog()
# SIG # Begin signature block
# MIIoiQYJKoZIhvcNAQcCoIIoejCCKHYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9muTnCjJ6REHe
# 2Ael6vx7jip18ML/wexA7VJda7CJvqCCILswggXJMIIEsaADAgECAhAbtY8lKt8j
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
# BCAbugY4j18Y0HJoR/If9P7BjnD+lObqr/pR2QXX+VAGuTANBgkqhkiG9w0BAQEF
# AASCAgA8KT2C9kkzVYHZSUBy2ZLB50C+kyxao+EC+7CrFWjffo59Rr2dmhgIdY9O
# bFsEMRprACJCkGR3GVDzNQUgy/sHE/3ON69d5ZzsNCTqVWdbuwxqGO6Mnk+7zdJy
# 2HkDp4l0oxyk+ngciTiBN74a+TguyurCvlgcHYgtdM3woCz12f6dlD9l+uy/KWQF
# LWBl7zUDrao73z0NtphT1McRZcvv6Ib3xtd/v6BWT4jsQ7qJQwLEFnQ/iZGNckGt
# +O1B0tCifT5dl1wvD6mN133IB0pP2E4P04mgSve2BkXFkeleYeucznviqSzVg+32
# HmNqV/OIeZWiOBU/HSlpr7GUDVfk0T44AMH3hSx5dZzKVmO8aLQ+M75xvWBZM5T2
# P40gkZdyntvGVKNCGlYqIHPIlM30aXMeQtAd9lIOihUwdqsgfNUMmcR3+3d0VqjP
# hjpR8WbxJNxtMnZZHiXWprtKnKKuoNJ09d5igAJIP61w1ukCrK0S0fvZSeiKC6Ru
# uHOCjN7+1ZmQsA3QB9YqpakIxCvBFYaj41yTj2sZN0dqhCG5CnZ2DIh6rhMlG53p
# aK1HeLxFo3uHYjEYEBjHhvZYmSxMsagQ4r4TeF2ox7c6ff2SDtE4SGnDAbUplpLP
# +PZDcc7A8bFvzC1rZ7pbvidDopE481u4hBqoHISVB2Dk4/EU4qGCBAQwggQABgkq
# hkiG9w0BCQYxggPxMIID7QIBATBrMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3Rh
# bXBpbmcgMjAyMSBDQQIRAJ6cBPZVqLSnAm1JjGx4jaowDQYJYIZIAWUDBAICBQCg
# ggFXMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcN
# MjUxMDI1MTc1ODM5WjA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCDPodw1ne0rw8uJ
# D6Iw5dr3e1QPGm4rI93PF1ThjPqg1TA/BgkqhkiG9w0BCQQxMgQwF5va3g4W9Smo
# MBqsYyJXwDOC4WRKAi7J8Is0Z9ROdwpoyzG/49jzQdqGnjSlJSeNMIGgBgsqhkiG
# 9w0BCRACDDGBkDCBjTCBijCBhwQUwyW4mxf8xQJgYc4rcXtFB92camowbzBapFgw
# VjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5B
# LjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBAhEAnpwE9lWo
# tKcCbUmMbHiNqjANBgkqhkiG9w0BAQEFAASCAgAgl407BKcg19DW1LSmZngR4d13
# Z6SAZoT2QwlU+GIfgIbAcshxWQN/C0f0kkuJqw2zeHoXdjN+9wsrfq3Wbo80QVFP
# eS6x8Tb0dpIpJwoz2VOTNXf2TtY/ub6RW6ptyv3T5Xb6TqNz+BlsOsuLps02yTab
# RhKUu3ou0p26bB+mHNe+l7ffOQ4K28yA5HOkXsdqkjRkfqWZdrOR9Vsv5CY0hxls
# QayLOQZ4hwF3lzJ+wg6Xn64WjGF6rLr5RuE9Z7vd0TXIQl/5PT72juIyLvOtKNur
# J/dSUN3XPcPk7WI6djVU9bz9vBMg5LBPoK5nh5s2uaV9FdX3FURriJpSDrxiS+4G
# n1glMAkLC8/jgWqv1MXXbXJvHFOfP8+TFm5MDs8VoOMneQjzn1JaNnbSorsCH0HF
# ICgQVLtlb+sb8YN2Lo5LrLfJj++opLQl/JAcAUiBFEopXe5SHg6RZloL90i+6uSV
# onWPLY3lb9uVMI+ifKzhes5uJWzQHZ5GYsw4WNN03zryPhkm+8hK5azjkYTqAo/i
# 2vRiTgudyEXzP6mQR+sva8ai2A0dicuOULc+cA2Ks4o2LepXJrwU+tIT1KAvtkzW
# M0+28YDv79a8nr+pmNvJs57dWj1er5NjVJNzM/3JFnLb+0R0mm0xnFSFUaDmBt6m
# ZwkaoRDHPfNSaPxL/A==
# SIG # End signature block
