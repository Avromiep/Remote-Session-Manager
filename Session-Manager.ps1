<#
================================================================================
 Session-Manager.ps1  --  Remote Session Manager (WPF edition, v2.1)

 A modern WPF rewrite of Win-Session-Manager. Queries multiple RDS / terminal
 servers for logged-on sessions and lets a helpdesk operator shadow, connect,
 message, disconnect, log off, reset, and inspect/kill a user's processes.

 TARGET RUNTIME
   Windows PowerShell 5.1 (powershell.exe). WPF needs an STA thread (5.1 is STA
   by default) and 5.1 is present on every Server/desktop 2012 R2+. The script
   self-relaunches under powershell.exe -STA (elevated) if needed.

 FEATURES
   * Server list from the textbox or servers.txt; quser /server: per server
   * C: free space + AD display-name cache; parallel querying (runspace pool)
   * Shadow / Connect (RDP) / Send Message / Disconnect / Log Off / Reset
   * Per-user PROCESS list + End Task (remote DCOM CIM, filtered by SessionId)
   * "Shadow computer" box: look up ANY machine by name, drop its session into
     the grid, then right-click -> Shadow
   * Light / Dark theme with a toggle (persisted in settings.ini)
   * Auto-refresh toggle + interval; search filter; sortable columns

 TRANSPORT NOTE
   Remote WMI/CIM uses DCOM CIM sessions (New-CimSessionOption -Protocol Dcom),
   so no WinRM is required -- same RPC/DCOM transport the old Get-WmiObject used.
================================================================================
#>

# ----------------------------------------------------------------------------
# SECTION 1: STA + elevation. Relaunch under powershell.exe as admin if needed.
# ----------------------------------------------------------------------------
$ScriptPath = $PSCommandPath
if (-not $ScriptPath) { $ScriptPath = $MyInvocation.MyCommand.Definition }
$ScriptDir  = Split-Path -Parent $ScriptPath

$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$IsSta     = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA'
$IsDesktop = $PSVersionTable.PSEdition -ne 'Core'

if (-not ($IsAdmin -and $IsSta -and $IsDesktop)) {
    $psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $argList = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$ScriptPath`"")
    try { Start-Process -FilePath $psExe -ArgumentList $argList -Verb RunAs } catch { }
    return
}

# ----------------------------------------------------------------------------
# SECTION 2: Assemblies + paths.
# ----------------------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # only for MessageBox convenience

Set-Location $ScriptDir
$ServersFile  = Join-Path $ScriptDir 'servers.txt'
$SettingsFile = Join-Path $ScriptDir 'settings.ini'
$CacheFile    = Join-Path $ScriptDir 'scriptcache'
$LogoIco      = Join-Path $ScriptDir 'Images\logo.ico'

# ----------------------------------------------------------------------------
# SECTION 3: Settings (ini) -- backward compatible with the original keys.
# ----------------------------------------------------------------------------
function Read-Settings {
    $s = @{ RefreshOnStartup = 0; DefaultShadowOptions = 0; AutoRefresh = 0; AutoRefreshSeconds = 30; DarkMode = 0 }
    if (Test-Path $SettingsFile) {
        foreach ($line in (Get-Content $SettingsFile)) {
            if ($line -match 'RefreshOnStartup=(\d)')        { $s.RefreshOnStartup     = [int]$matches[1] }
            elseif ($line -match 'DefaultShadowOptions=(\d)') { $s.DefaultShadowOptions = [int]$matches[1] }
            elseif ($line -match 'AutoRefresh=(\d)')          { $s.AutoRefresh          = [int]$matches[1] }
            elseif ($line -match 'AutoRefreshSeconds=(\d+)')  { $s.AutoRefreshSeconds   = [int]$matches[1] }
            elseif ($line -match 'DarkMode=(\d)')             { $s.DarkMode             = [int]$matches[1] }
        }
    }
    return $s
}
function Write-Settings {
    param([hashtable]$s)
    $content = @(
        '[Settings]'
        "RefreshOnStartup=$($s.RefreshOnStartup)"
        "DefaultShadowOptions=$($s.DefaultShadowOptions)"
        "AutoRefresh=$($s.AutoRefresh)"
        "AutoRefreshSeconds=$($s.AutoRefreshSeconds)"
        "DarkMode=$($s.DarkMode)"
    ) -join "`r`n"
    try { Set-Content -Path $SettingsFile -Value $content -Encoding ASCII } catch { }
}
$script:Settings = Read-Settings
$script:DarkMode = [bool]$script:Settings.DarkMode

# ----------------------------------------------------------------------------
# SECTION 4: Display-name cache (compatible with the original CSV 'scriptcache').
# ----------------------------------------------------------------------------
function Read-DisplayNameCache {
    $cache = @{}
    if (Test-Path $CacheFile) {
        try {
            Import-Csv -Path $CacheFile -Delimiter ',' -ErrorAction Stop | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.Username) -and -not [string]::IsNullOrWhiteSpace($_.DisplayName)) {
                    $cache[$_.Username] = $_.DisplayName
                }
            }
        } catch { }
    }
    return $cache
}
function Write-DisplayNameCache {
    param([hashtable]$Cache)
    try {
        $Cache.GetEnumerator() |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) } |
            ForEach-Object { [PSCustomObject]@{ Username = $_.Key; DisplayName = [string]$_.Value } } |
            Export-Csv -Path $CacheFile -NoTypeInformation -Delimiter ',' -Force
    } catch { }
}
$script:NameCache = Read-DisplayNameCache

function Resolve-DisplayName {
    param([string]$Username)
    if ([string]::IsNullOrWhiteSpace($Username)) { return '' }
    if ($script:NameCache.ContainsKey($Username)) { return $script:NameCache[$Username] }
    $name = $Username
    try {
        $searcher = [ADSISearcher]"(samaccountname=$Username)"
        $searcher.PropertiesToLoad.Add('displayName') | Out-Null
        $searcher.ClientTimeout   = [TimeSpan]::FromSeconds(3)
        $searcher.ServerTimeLimit = [TimeSpan]::FromSeconds(3)
        $r = $searcher.FindOne()
        if ($r -and $r.Properties['displayname'].Count) {
            $dn = [string]$r.Properties['displayname'][0]
            if (-not [string]::IsNullOrWhiteSpace($dn)) { $name = $dn }
        }
    } catch { }
    $script:NameCache[$Username] = $name
    return $name
}

# ----------------------------------------------------------------------------
# SECTION 5: quser parser (as a string so it can be injected into worker runspaces).
# ----------------------------------------------------------------------------
$QuserParserText = @'
function ConvertFrom-Quser {
    param([string[]]$Raw)
    $rows = @()
    if (-not $Raw -or $Raw.Count -lt 2) { return $rows }
    $header = $Raw[0]
    $iUser  = $header.IndexOf('USERNAME')
    $iSess  = $header.IndexOf('SESSIONNAME')
    $iId    = $header.IndexOf('ID')
    $iState = $header.IndexOf('STATE')
    $iIdle  = $header.IndexOf('IDLE TIME')
    $iLogon = $header.IndexOf('LOGON TIME')
    if ($iId -lt 0 -or $iState -lt 0) { return $rows }
    for ($i = 1; $i -lt $Raw.Count; $i++) {
        $line = $Raw[$i]
        if ([string]::IsNullOrWhiteSpace($line) -or $line.Length -lt $iState) { continue }
        # Username ends where SESSIONNAME begins (so "console"/"rdp-tcp#N" never
        # bleeds into the username).
        $userEnd = if ($iSess -gt $iUser) { $iSess } else { $iId }
        $user  = $line.Substring([Math]::Max(0,$iUser), [Math]::Min($userEnd,$line.Length) - [Math]::Max(0,$iUser)).Trim().TrimStart('>').Trim()
        $id    = $line.Substring($iId,    [Math]::Min($iState,$line.Length) - $iId).Trim()
        $endState = if ($iIdle -gt $iState) { $iIdle } else { $line.Length }
        $state = $line.Substring($iState, [Math]::Min($endState,$line.Length) - $iState).Trim()
        $logon = if ($iLogon -ge 0 -and $line.Length -gt $iLogon) { $line.Substring($iLogon).Trim() } else { '' }
        if ($id -match '^\d+$') {
            $rows += [pscustomobject]@{ Username = $user; SessionID = $id; State = $state; LogonTime = $logon }
        }
    }
    return $rows
}
'@
Invoke-Expression $QuserParserText   # define it in this (UI) runspace too

# ----------------------------------------------------------------------------
# SECTION 6: Per-server worker scriptblock (runs in a background runspace).
# ----------------------------------------------------------------------------
$ServerWorker = [scriptblock]::Create(@"
param(`$Server)
$QuserParserText
`$out = [pscustomobject]@{ Server = `$Server; Sessions = @(); DiskFree = 'N/A'; Error = `$null }

# --- C: free space via a DCOM CIM session (no WinRM needed) ---
try {
    `$opt = New-CimSessionOption -Protocol Dcom
    `$cs  = New-CimSession -ComputerName `$Server -SessionOption `$opt -OperationTimeoutSec 15 -ErrorAction Stop
    `$disk = Get-CimInstance -CimSession `$cs -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
    if (`$disk) { `$out.DiskFree = ('{0:N1} GB' -f (`$disk.FreeSpace / 1GB)) }
    Remove-CimSession `$cs -ErrorAction SilentlyContinue
} catch { `$out.DiskFree = 'N/A' }

# --- sessions via quser ---
try {
    `$raw = quser /server:`$Server 2>&1
    `$textq = (`$raw | Out-String)
    if (`$textq -match 'No User exists') {
        `$out.Sessions = @()
    } elseif (`$textq -match 'Error|RPC|cannot|denied') {
        `$out.Error = (`$textq.Trim() -split "`r?`n")[0]
    } else {
        `$out.Sessions = @(ConvertFrom-Quser -Raw (`$raw | ForEach-Object { [string]`$_ }))
    }
} catch { `$out.Error = `$_.Exception.Message }

`$out
"@)

# ----------------------------------------------------------------------------
# SECTION 7: Theme -- palettes + Apply-Theme (swaps DynamicResource brushes).
# ----------------------------------------------------------------------------
$script:Themes = @{
    Light = @{ WinBg='#F4F6F8'; CardBg='#FFFFFF'; Border='#E1E4E8'; Text='#1F2328'; Subtle='#6A737D'; RowAlt='#F6F8FA'; Accent='#0A66C2'; Hover='#EAF3FB' }
    Dark  = @{ WinBg='#1B1D21'; CardBg='#26282D'; Border='#3A3D44'; Text='#E6E8EA'; Subtle='#9AA0A6'; RowAlt='#2E3136'; Accent='#4C9AFF'; Hover='#33383F' }
}
function Apply-Theme {
    param($win, [bool]$Dark)
    $pal = if ($Dark) { $script:Themes.Dark } else { $script:Themes.Light }
    foreach ($k in $pal.Keys) {
        $col = [System.Windows.Media.ColorConverter]::ConvertFromString($pal[$k])
        $brush = New-Object System.Windows.Media.SolidColorBrush $col
        $brush.Freeze()
        # Cast to [Brush] so the value stored is the raw object, not a PSObject
        # wrapper (which DynamicResource would reject as a brush).
        $win.Resources[$k] = [System.Windows.Media.Brush]$brush
    }
}

# Shared styles injected into both windows (reference DynamicResource brushes).
$ThemeStyles = @'
        <Style x:Key="Tool" TargetType="Button">
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Foreground" Value="{DynamicResource Text}"/>
            <Setter Property="Background" Value="{DynamicResource CardBg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="b" CornerRadius="5" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="b" Property="Background" Value="{DynamicResource Hover}"/>
                                <Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource Accent}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource Tool}">
            <Setter Property="Background" Value="{DynamicResource Accent}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="{DynamicResource Accent}"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource CardBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource Text}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="8,7"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource Accent}"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{DynamicResource Text}"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding State}" Value="Disc">
                    <Setter Property="Foreground" Value="{DynamicResource Subtle}"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
'@

# ----------------------------------------------------------------------------
# SECTION 8: Main window XAML (brushes are DynamicResource, set by Apply-Theme).
# ----------------------------------------------------------------------------
$mainTemplate = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Remote Session Manager" Height="640" Width="1120"
        WindowStartupLocation="CenterScreen" Background="{DynamicResource WinBg}"
        FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
%STYLES%
    </Window.Resources>

    <DockPanel>
        <!-- Toolbar row 1: servers + theme + settings -->
        <Border DockPanel.Dock="Top" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}"
                BorderThickness="0,0,0,1" Padding="12,10,12,6">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Servers:" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="{DynamicResource Subtle}"/>
                <Border Grid.Column="1" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="5" Background="{DynamicResource CardBg}">
                    <TextBox x:Name="TxtServers" BorderThickness="0" Height="26" VerticalContentAlignment="Center" Padding="6,0"
                             Background="Transparent" Foreground="{DynamicResource Text}" CaretBrush="{DynamicResource Text}"
                             ToolTip="Comma-separated. Leave blank to use servers.txt. Enter = Refresh"/>
                </Border>
                <Button Grid.Column="2" x:Name="BtnDark" Style="{StaticResource Tool}" Margin="8,0,0,0" Width="36"
                        FontFamily="Segoe UI Symbol" FontSize="15" ToolTip="Toggle light / dark theme"/>
                <Button Grid.Column="3" x:Name="BtnSettings" Style="{StaticResource Tool}" Content="Settings"/>
            </Grid>
        </Border>

        <!-- Toolbar row 2: search/actions (left) + shadow-computer lookup (right) -->
        <Border DockPanel.Dock="Top" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}"
                BorderThickness="0,0,0,1" Padding="12,6,12,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,6,0" Foreground="{DynamicResource Subtle}"/>
                    <Border BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="5" Background="{DynamicResource CardBg}" Margin="0,0,14,0">
                        <TextBox x:Name="TxtSearch" Width="200" Height="26" BorderThickness="0" VerticalContentAlignment="Center"
                                 Padding="6,0" Background="Transparent" Foreground="{DynamicResource Text}" CaretBrush="{DynamicResource Text}"/>
                    </Border>
                    <Button x:Name="BtnRefresh" Style="{StaticResource Primary}" Content="Refresh"/>
                    <CheckBox x:Name="ChkAuto" Content="Auto-refresh" VerticalAlignment="Center" Margin="6,0,4,0" Foreground="{DynamicResource Text}"/>
                    <TextBox x:Name="TxtInterval" Width="40" Height="24" Text="30" VerticalContentAlignment="Center" TextAlignment="Center"
                             Background="{DynamicResource CardBg}" Foreground="{DynamicResource Text}" BorderBrush="{DynamicResource Border}"
                             ToolTip="Auto-refresh interval (seconds)"/>
                    <TextBlock Text="s" VerticalAlignment="Center" Margin="3,0,0,0" Foreground="{DynamicResource Subtle}"/>
                </StackPanel>
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <TextBlock Text="Shadow computer:" VerticalAlignment="Center" Margin="0,0,6,0" Foreground="{DynamicResource Subtle}"/>
                    <Border BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="5" Background="{DynamicResource CardBg}" Margin="0,0,8,0">
                        <TextBox x:Name="TxtComputer" Width="160" Height="26" BorderThickness="0" VerticalContentAlignment="Center"
                                 Padding="6,0" Background="Transparent" Foreground="{DynamicResource Text}" CaretBrush="{DynamicResource Text}"
                                 ToolTip="Type a computer name, then Look Up (or Enter). Its session appears below; right-click it -> Shadow."/>
                    </Border>
                    <Button x:Name="BtnLookup" Style="{StaticResource Tool}" Content="Look Up" Margin="0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Status bar -->
        <Border DockPanel.Dock="Bottom" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}"
                BorderThickness="0,1,0,0" Padding="12,6">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" x:Name="StatusText" Foreground="{DynamicResource Subtle}" Text="Ready." VerticalAlignment="Center"/>
                <ProgressBar Grid.Column="1" x:Name="Progress" Width="180" Height="12" Minimum="0" Maximum="100" Visibility="Collapsed"/>
            </Grid>
        </Border>

        <!-- Sessions grid -->
        <Border Margin="12" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="6">
            <DataGrid x:Name="Grid" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                      SelectionMode="Extended" CanUserResizeRows="False" GridLinesVisibility="Horizontal"
                      RowHeaderWidth="0" BorderThickness="0" Background="{DynamicResource CardBg}"
                      Foreground="{DynamicResource Text}" RowBackground="{DynamicResource CardBg}"
                      AlternatingRowBackground="{DynamicResource RowAlt}" HorizontalGridLinesBrush="{DynamicResource Border}"
                      FontSize="13" CanUserAddRows="False">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Username"     Binding="{Binding Username}"     Width="140"/>
                    <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}"  Width="180"/>
                    <DataGridTextColumn Header="Server"       Binding="{Binding Server}"       Width="140"/>
                    <DataGridTextColumn Header="Session ID"   Binding="{Binding SessionID}"    Width="90"/>
                    <DataGridTextColumn Header="State"        Binding="{Binding State}"        Width="90"/>
                    <DataGridTextColumn Header="Logon Time"   Binding="{Binding LogonTime}"    Width="150"/>
                    <DataGridTextColumn Header="C: Free"      Binding="{Binding DiskSpace}"    Width="90"/>
                    <DataGridTextColumn Header="Sessions"     Binding="{Binding SessionsOpen}" Width="80"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>
    </DockPanel>
</Window>
'@
[xml]$xaml = $mainTemplate.Replace('%STYLES%', $ThemeStyles)

# ----------------------------------------------------------------------------
# SECTION 9: Load main window + controls.
# ----------------------------------------------------------------------------
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$TxtServers  = $window.FindName('TxtServers')
$TxtSearch   = $window.FindName('TxtSearch')
$TxtComputer = $window.FindName('TxtComputer')
$BtnLookup   = $window.FindName('BtnLookup')
$BtnRefresh  = $window.FindName('BtnRefresh')
$BtnSettings = $window.FindName('BtnSettings')
$BtnDark     = $window.FindName('BtnDark')
$ChkAuto     = $window.FindName('ChkAuto')
$TxtInterval = $window.FindName('TxtInterval')
$StatusText  = $window.FindName('StatusText')
$Progress    = $window.FindName('Progress')
$Grid        = $window.FindName('Grid')

# Window icon (OnLoad cache so the .ico file isn't locked).
try {
    if (Test-Path $LogoIco) {
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit(); $bmp.CacheOption = 'OnLoad'; $bmp.UriSource = [Uri]$LogoIco; $bmp.EndInit()
        $window.Icon = $bmp
    }
} catch { }

# Bound, filterable collection of sessions.
$script:Sessions = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$sessionsView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:Sessions)
$Grid.ItemsSource = $sessionsView

$script:SearchText = ''
$sessionsView.Filter = [Predicate[object]] {
    param($it)
    if ([string]::IsNullOrWhiteSpace($script:SearchText)) { return $true }
    $t = $script:SearchText
    return (
        ($it.Username    -and $it.Username.ToLower().Contains($t)) -or
        ($it.DisplayName -and $it.DisplayName.ToLower().Contains($t)) -or
        ($it.Server      -and $it.Server.ToLower().Contains($t)) -or
        ($it.State       -and $it.State.ToLower().Contains($t)) -or
        ($it.SessionID   -and ([string]$it.SessionID).Contains($t))
    )
}
$TxtSearch.Add_TextChanged({ $script:SearchText = $TxtSearch.Text.Trim().ToLower(); $sessionsView.Refresh() })

# ----------------------------------------------------------------------------
# SECTION 10: Refresh -- parallel per-server query via a runspace pool, polled by
# a DispatcherTimer. Supports full refresh and single-computer "append" lookups.
# ----------------------------------------------------------------------------
$script:Pool         = $null
$script:Jobs         = @()
$script:PollTimer    = $null
$script:RefreshStart = $null
$script:AppendMode   = $false
$script:BatchServers = @()
$script:SelectServer = $null

function Get-ServerList {
    if ($TxtServers.Text.Trim()) {
        return @($TxtServers.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if (Test-Path $ServersFile) {
        return @(Get-Content $ServersFile | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    return @()
}

function Start-Refresh {
    param([string[]]$Servers, [switch]$Append, [string]$SelectServer)
    if ($script:Pool) { return }   # a refresh is already running

    if (-not $Servers -or $Servers.Count -eq 0) {
        if ($Append) { return }
        $Servers = Get-ServerList
        if ($Servers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No servers specified and servers.txt was not found.",
                "Remote Session Manager", 'OK', 'Warning') | Out-Null
            return
        }
    }

    $script:AppendMode   = [bool]$Append
    $script:BatchServers = @($Servers)
    $script:SelectServer = $SelectServer

    $BtnRefresh.IsEnabled = $false; $BtnLookup.IsEnabled = $false
    $Progress.Visibility = 'Visible'; $Progress.Value = 0
    $StatusText.Text = if ($Append) { "Looking up $($Servers[0])..." } else { "Querying $($Servers.Count) server(s)..." }
    $script:RefreshStart = Get-Date

    $script:Pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(16, [Math]::Max(2, $Servers.Count)))
    $script:Pool.ApartmentState = 'MTA'
    $script:Pool.Open()
    $script:Jobs = foreach ($s in $Servers) {
        $ps = [powershell]::Create(); $ps.RunspacePool = $script:Pool
        [void]$ps.AddScript($ServerWorker).AddArgument($s)
        [pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Server = $s; Done = $false; Result = $null }
    }

    $script:PollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:PollTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:PollTimer.Add_Tick({ Poll-Refresh })
    $script:PollTimer.Start()
}

function Poll-Refresh {
    $total = $script:Jobs.Count; $done = 0
    foreach ($j in $script:Jobs) {
        if (-not $j.Done -and $j.Handle.IsCompleted) {
            try { $j.Result = $j.PS.EndInvoke($j.Handle) } catch { $j.Result = $null }
            try { $j.PS.Dispose() } catch { }
            $j.Done = $true
        }
        if ($j.Done) { $done++ }
    }
    $Progress.Value = [math]::Round(($done / $total) * 100)
    if ($done -eq $total -or ((Get-Date) - $script:RefreshStart).TotalSeconds -gt 60) {
        foreach ($j in $script:Jobs) { if (-not $j.Done) { try { $j.PS.Stop(); $j.PS.Dispose() } catch { }; $j.Done = $true } }
        $script:PollTimer.Stop(); $script:PollTimer = $null
        Complete-Refresh
    }
}

function Complete-Refresh {
    $all = @(); $counts = @{}; $errors = @()
    foreach ($j in $script:Jobs) {
        $obj = if ($j.Result) { $j.Result | Select-Object -Last 1 } else { $null }
        if ($null -eq $obj) { $errors += "$($j.Server): no response"; $counts[$j.Server] = 0; continue }
        if ($obj.Error) { $errors += "$($obj.Server): $($obj.Error)" }
        $counts[$obj.Server] = @($obj.Sessions).Count
        foreach ($sess in @($obj.Sessions)) {
            $all += [pscustomobject]@{
                Username = $sess.Username; DisplayName = ''; Server = $obj.Server
                SessionID = $sess.SessionID; State = $sess.State; LogonTime = $sess.LogonTime
                DiskSpace = $obj.DiskFree; SessionsOpen = 0
            }
        }
    }
    foreach ($row in $all) {
        $row.DisplayName  = Resolve-DisplayName -Username $row.Username
        $row.SessionsOpen = $counts[$row.Server]
    }
    Write-DisplayNameCache -Cache $script:NameCache

    if ($script:AppendMode) {
        # Replace any existing rows for the queried server(s), then add fresh ones.
        for ($i = $script:Sessions.Count - 1; $i -ge 0; $i--) {
            if ($script:BatchServers -contains $script:Sessions[$i].Server) { $script:Sessions.RemoveAt($i) }
        }
        foreach ($row in ($all | Sort-Object Server, Username)) { $script:Sessions.Add($row) }
    } else {
        $script:Sessions.Clear()
        foreach ($row in ($all | Sort-Object Server, Username)) { $script:Sessions.Add($row) }
    }
    $sessionsView.Refresh()

    try { $script:Pool.Close(); $script:Pool.Dispose() } catch { }
    $script:Pool = $null; $script:Jobs = @()
    $Progress.Visibility = 'Collapsed'
    $BtnRefresh.IsEnabled = $true; $BtnLookup.IsEnabled = $true

    if ($script:AppendMode -and $script:SelectServer) {
        # Select + reveal the looked-up computer's session so it's ready to shadow.
        $target = $script:Sessions | Where-Object { $_.Server -eq $script:SelectServer } | Select-Object -First 1
        if ($target) {
            $Grid.SelectedItem = $target
            try { $Grid.ScrollIntoView($target) } catch { }
            $Grid.Focus() | Out-Null
            $StatusText.Text = "Found $($target.Username) on $($script:SelectServer) (session $($target.SessionID)) -- right-click it -> Shadow Session."
        } elseif ($errors.Count) {
            $StatusText.Text = "Look up failed: $($errors[0])"
        } else {
            $StatusText.Text = "No active sessions found on $($script:SelectServer)."
        }
    } else {
        $secs = [math]::Round(((Get-Date) - $script:RefreshStart).TotalSeconds, 1)
        $msg = "Updated {0:HH:mm:ss}  -  {1} session(s) across {2} server(s)  ({3}s)" -f `
            (Get-Date), $script:Sessions.Count, $counts.Keys.Count, $secs
        if ($errors.Count) { $msg += "  -  {0} server(s) had errors" -f $errors.Count }
        $StatusText.Text = $msg
    }
    $script:AppendMode = $false; $script:SelectServer = $null; $script:BatchServers = @()
}

$BtnRefresh.Add_Click({ Start-Refresh })

# --- Shadow-computer lookup ---
function Invoke-Lookup {
    $name = $TxtComputer.Text.Trim()
    if (-not $name) { return }
    if ($TxtSearch.Text) { $TxtSearch.Text = '' }   # clear filter so the new row is visible
    Start-Refresh -Servers @($name) -Append -SelectServer $name
}
$BtnLookup.Add_Click({ Invoke-Lookup })
$TxtComputer.Add_KeyDown({ param($s,$e) if ($e.Key -eq 'Return') { Invoke-Lookup } })
$TxtServers.Add_KeyDown({ param($s,$e) if ($e.Key -eq 'Return') { Start-Refresh } })

# ----------------------------------------------------------------------------
# SECTION 11: Selection helpers.
# ----------------------------------------------------------------------------
function Get-SelectedRows { @($Grid.SelectedItems) }
function Get-SelectedRow  { if ($Grid.SelectedItems.Count -ge 1) { $Grid.SelectedItems[0] } else { $null } }

# ----------------------------------------------------------------------------
# SECTION 12: Session actions.
# ----------------------------------------------------------------------------
function Invoke-Shadow {
    $row = Get-SelectedRow
    if (-not $row) { return }
    $settings = Read-Settings

    $dlg = New-Object System.Windows.Window
    $dlg.Title = "Shadow $($row.Username) on $($row.Server)"
    $dlg.Width = 320; $dlg.Height = 170; $dlg.WindowStartupLocation = 'CenterOwner'
    $dlg.Owner = $window; $dlg.ResizeMode = 'NoResize'
    Apply-Theme -win $dlg -Dark $script:DarkMode
    $dlg.Background = $dlg.Resources['WinBg']
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $fg = $dlg.Resources['Text']
    $cbConsent = New-Object System.Windows.Controls.CheckBox
    $cbConsent.Content = 'No consent prompt'; $cbConsent.Margin = '0,0,0,10'; $cbConsent.Foreground = $fg
    $cbConsent.IsChecked = [bool]$settings.DefaultShadowOptions
    $cbControl = New-Object System.Windows.Controls.CheckBox
    $cbControl.Content = 'Enable control (unchecked = view only)'; $cbControl.Margin = '0,0,0,16'; $cbControl.Foreground = $fg
    $cbControl.IsChecked = [bool]$settings.DefaultShadowOptions
    $btns = New-Object System.Windows.Controls.StackPanel; $btns.Orientation = 'Horizontal'; $btns.HorizontalAlignment = 'Right'
    $ok = New-Object System.Windows.Controls.Button; $ok.Content = 'Connect'; $ok.Width = 80; $ok.Margin = '0,0,8,0'; $ok.IsDefault = $true
    $cancel = New-Object System.Windows.Controls.Button; $cancel.Content = 'Cancel'; $cancel.Width = 80; $cancel.IsCancel = $true
    $ok.Add_Click({ $dlg.DialogResult = $true })
    $btns.Children.Add($ok) | Out-Null; $btns.Children.Add($cancel) | Out-Null
    $sp.Children.Add($cbConsent) | Out-Null; $sp.Children.Add($cbControl) | Out-Null; $sp.Children.Add($btns) | Out-Null
    $dlg.Content = $sp
    if ($dlg.ShowDialog()) {
        $args = "/v:$($row.Server) /shadow:$($row.SessionID)"
        if ($cbControl.IsChecked) { $args += " /control" }
        if ($cbConsent.IsChecked) { $args += " /noconsentprompt" }
        try { Start-Process "mstsc.exe" -ArgumentList $args } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, "Shadow") | Out-Null
        }
    }
}

function Invoke-Connect {
    $row = Get-SelectedRow
    if (-not $row) { return }
    try { Start-Process "mstsc.exe" -ArgumentList "/v:$($row.Server)" } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Connect") | Out-Null
    }
}

function Invoke-Message {
    $rows = Get-SelectedRows
    if ($rows.Count -eq 0) { return }
    $dlg = New-Object System.Windows.Window
    $dlg.Title = "Send Message"; $dlg.Width = 420; $dlg.Height = 240
    $dlg.WindowStartupLocation = 'CenterOwner'; $dlg.Owner = $window; $dlg.ResizeMode = 'NoResize'
    Apply-Theme -win $dlg -Dark $script:DarkMode
    $dlg.Background = $dlg.Resources['WinBg']
    $fg = $dlg.Resources['Text']
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = "Message to $($rows.Count) session(s):"; $lbl.Margin = '0,0,0,8'; $lbl.Foreground = $fg
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.AcceptsReturn = $true; $tb.TextWrapping = 'Wrap'; $tb.Height = 110; $tb.Margin = '0,0,0,12'
    $tb.VerticalScrollBarVisibility = 'Auto'
    $tb.Background = $dlg.Resources['CardBg']; $tb.Foreground = $fg; $tb.CaretBrush = $fg
    $btns = New-Object System.Windows.Controls.StackPanel; $btns.Orientation = 'Horizontal'; $btns.HorizontalAlignment = 'Right'
    $ok = New-Object System.Windows.Controls.Button; $ok.Content = 'Send'; $ok.Width = 80; $ok.Margin = '0,0,8,0'; $ok.IsDefault = $true
    $cancel = New-Object System.Windows.Controls.Button; $cancel.Content = 'Cancel'; $cancel.Width = 80; $cancel.IsCancel = $true
    $ok.Add_Click({ $dlg.DialogResult = $true })
    $btns.Children.Add($ok) | Out-Null; $btns.Children.Add($cancel) | Out-Null
    $sp.Children.Add($lbl) | Out-Null; $sp.Children.Add($tb) | Out-Null; $sp.Children.Add($btns) | Out-Null
    $dlg.Content = $sp
    if ($dlg.ShowDialog() -and $tb.Text.Trim()) {
        $full = "Message from IT Support:`n`n$($tb.Text.Trim())"
        foreach ($row in $rows) {
            try { msg $row.Username /server:$($row.Server) $full } catch {
                [System.Windows.MessageBox]::Show("Error messaging $($row.Username) on $($row.Server): $_", "Send Message") | Out-Null
            }
        }
        $StatusText.Text = "Sent message to $($rows.Count) session(s)."
    }
}

function Invoke-Disconnect {
    $rows = Get-SelectedRows
    if ($rows.Count -eq 0) { return }
    $list = ($rows | ForEach-Object { "  $($_.Username)  (session $($_.SessionID) on $($_.Server))" }) -join "`n"
    if ([System.Windows.MessageBox]::Show("Disconnect these session(s)? They keep running.`n`n$list",
            "Confirm Disconnect", 'YesNo', 'Warning') -ne 'Yes') { return }
    foreach ($row in $rows) {
        try { tsdiscon $row.SessionID /server:$($row.Server) } catch {
            [System.Windows.MessageBox]::Show("Error disconnecting $($row.Username): $_", "Disconnect") | Out-Null
        }
    }
    Start-Refresh
}

function Invoke-Logoff {
    $rows = Get-SelectedRows
    if ($rows.Count -eq 0) { return }
    $list = ($rows | ForEach-Object { "  $($_.Username)  (session $($_.SessionID) on $($_.Server))" }) -join "`n"
    if ([System.Windows.MessageBox]::Show("Log off these session(s)? Unsaved work will be lost.`n`n$list",
            "Confirm Log Off", 'YesNo', 'Warning') -ne 'Yes') { return }
    foreach ($row in $rows) {
        try { logoff $row.SessionID /server:$($row.Server) } catch {
            [System.Windows.MessageBox]::Show("Error logging off $($row.Username): $_", "Log Off") | Out-Null
        }
    }
    Start-Refresh
}

function Invoke-Reset {
    $rows = Get-SelectedRows
    if ($rows.Count -eq 0) { return }
    $list = ($rows | ForEach-Object { "  $($_.Username)  (session $($_.SessionID) on $($_.Server))" }) -join "`n"
    if ([System.Windows.MessageBox]::Show("RESET these session(s)? This forcibly ends them immediately.`n`n$list",
            "Confirm Reset", 'YesNo', 'Warning') -ne 'Yes') { return }
    foreach ($row in $rows) {
        try { reset session $row.SessionID /server:$($row.Server) } catch {
            [System.Windows.MessageBox]::Show("Error resetting $($row.Username): $_", "Reset") | Out-Null
        }
    }
    Start-Refresh
}

# ----------------------------------------------------------------------------
# SECTION 13: Process window -- list + End Task for a user's session (remote).
# ----------------------------------------------------------------------------
$procTemplate = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Processes" Height="560" Width="620" WindowStartupLocation="CenterOwner"
        Background="{DynamicResource WinBg}" FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
%STYLES%
    </Window.Resources>
    <DockPanel>
        <Border DockPanel.Dock="Top" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}" BorderThickness="0,0,0,1" Padding="12,10">
            <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="PHeader" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="{DynamicResource Text}"/>
                <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="18,0,6,0" Foreground="{DynamicResource Subtle}"/>
                <Border BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="5" Background="{DynamicResource CardBg}" Margin="0,0,12,0">
                    <TextBox x:Name="PFilter" Width="160" Height="26" BorderThickness="0" VerticalContentAlignment="Center" Padding="6,0"
                             Background="Transparent" Foreground="{DynamicResource Text}" CaretBrush="{DynamicResource Text}"/>
                </Border>
                <Button x:Name="PRefresh" Style="{StaticResource Tool}" Content="Refresh"/>
                <Button x:Name="PEnd" Style="{StaticResource Primary}" Content="End Task"/>
                <Button x:Name="PEndAll" Style="{StaticResource Tool}" Content="End All" Margin="0"
                        ToolTip="End every process in this session"/>
            </StackPanel>
        </Border>
        <Border DockPanel.Dock="Bottom" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}" BorderThickness="0,1,0,0" Padding="12,6">
            <TextBlock x:Name="PStatus" Foreground="{DynamicResource Subtle}" Text="Loading..."/>
        </Border>
        <Border Margin="12" Background="{DynamicResource CardBg}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="6">
            <DataGrid x:Name="PGrid" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column"
                      SelectionMode="Extended" CanUserResizeRows="False" GridLinesVisibility="Horizontal"
                      RowHeaderWidth="0" BorderThickness="0" Background="{DynamicResource CardBg}" Foreground="{DynamicResource Text}"
                      RowBackground="{DynamicResource CardBg}" AlternatingRowBackground="{DynamicResource RowAlt}"
                      HorizontalGridLinesBrush="{DynamicResource Border}" CanUserAddRows="False">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Process"  Binding="{Binding Name}"       Width="*"/>
                    <DataGridTextColumn Header="PID"      Binding="{Binding Pid}"        Width="90"/>
                    <DataGridTextColumn Header="Memory"   Binding="{Binding MemDisplay}" Width="120"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>
    </DockPanel>
</Window>
'@

function Show-ProcessWindow {
    param([string]$Server, [string]$SessionId, [string]$Username)

    [xml]$pxaml = $procTemplate.Replace('%STYLES%', $ThemeStyles)
    $pw = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $pxaml))
    Apply-Theme -win $pw -Dark $script:DarkMode
    $pw.Owner = $window
    $pw.Title = "Processes - $Username (session $SessionId on $Server)"
    $PHeader = $pw.FindName('PHeader'); $PFilter = $pw.FindName('PFilter')
    $PRefresh = $pw.FindName('PRefresh'); $PEnd = $pw.FindName('PEnd'); $PEndAll = $pw.FindName('PEndAll')
    $PStatus = $pw.FindName('PStatus'); $PGrid = $pw.FindName('PGrid')
    $PHeader.Text = "$Username on $Server (session $SessionId)"

    $cim = $null
    try {
        $opt = New-CimSessionOption -Protocol Dcom
        $cim = New-CimSession -ComputerName $Server -SessionOption $opt -OperationTimeoutSec 20 -ErrorAction Stop
    } catch { $PStatus.Text = "Could not connect to ${Server}: $($_.Exception.Message)" }

    $procs = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    $pview = [System.Windows.Data.CollectionViewSource]::GetDefaultView($procs)
    $PGrid.ItemsSource = $pview
    $pfilterText = [ref]''
    $pview.Filter = [Predicate[object]] {
        param($it)
        if ([string]::IsNullOrWhiteSpace($pfilterText.Value)) { return $true }
        return ($it.Name -and $it.Name.ToLower().Contains($pfilterText.Value))
    }
    $PFilter.Add_TextChanged({ $pfilterText.Value = $PFilter.Text.Trim().ToLower(); $pview.Refresh() }.GetNewClosure())

    $loadProcs = {
        if (-not $cim) { return }
        $PStatus.Text = "Loading processes..."; $procs.Clear()
        try {
            $list = Get-CimInstance -CimSession $cim -ClassName Win32_Process -Filter "SessionId=$SessionId" -ErrorAction Stop
            foreach ($p in ($list | Sort-Object -Property WorkingSetSize -Descending)) {
                $memMB = [math]::Round(($p.WorkingSetSize / 1MB), 1)
                $procs.Add([pscustomobject]@{ Name = $p.Name; Pid = [int]$p.ProcessId; MemMB = $memMB; MemDisplay = ('{0:N1} MB' -f $memMB) })
            }
            $pview.Refresh(); $PStatus.Text = "$($procs.Count) process(es)."
        } catch { $PStatus.Text = "Error: $($_.Exception.Message)" }
    }.GetNewClosure()

    $endTask = {
        $sel = @($PGrid.SelectedItems)
        if ($sel.Count -eq 0 -or -not $cim) { return }
        $names = ($sel | ForEach-Object { "  $($_.Name)  (PID $($_.Pid))" }) -join "`n"
        if ([System.Windows.MessageBox]::Show("End these process(es) for $Username on ${Server}?`n`n$names",
                "Confirm End Task", 'YesNo', 'Warning') -ne 'Yes') { return }
        $fail = 0
        foreach ($proc in $sel) {
            try {
                $inst = Get-CimInstance -CimSession $cim -ClassName Win32_Process -Filter "ProcessId=$($proc.Pid)" -ErrorAction Stop
                if ($inst) { Invoke-CimMethod -CimSession $cim -InputObject $inst -MethodName Terminate -ErrorAction Stop | Out-Null }
            } catch { $fail++ }
        }
        & $loadProcs
        if ($fail) { $PStatus.Text = "$($procs.Count) process(es). $fail could not be ended." }
    }.GetNewClosure()

    $endAll = {
        if (-not $cim -or $procs.Count -eq 0) { return }
        if ([System.Windows.MessageBox]::Show("End ALL $($procs.Count) processes for $Username on ${Server} (session $SessionId)?",
                "Confirm End All", 'YesNo', 'Warning') -ne 'Yes') { return }
        $fail = 0
        foreach ($proc in @($procs)) {
            try {
                $inst = Get-CimInstance -CimSession $cim -ClassName Win32_Process -Filter "ProcessId=$($proc.Pid)" -ErrorAction Stop
                if ($inst) { Invoke-CimMethod -CimSession $cim -InputObject $inst -MethodName Terminate -ErrorAction Stop | Out-Null }
            } catch { $fail++ }
        }
        & $loadProcs
        if ($fail) { $PStatus.Text = "$($procs.Count) process(es). $fail could not be ended." }
    }.GetNewClosure()

    $PRefresh.Add_Click($loadProcs)
    $PEnd.Add_Click($endTask)
    $PEndAll.Add_Click($endAll)
    $PGrid.Add_MouseDoubleClick($endTask)

    $pmenu = New-Object System.Windows.Controls.ContextMenu
    $pmi = New-Object System.Windows.Controls.MenuItem; $pmi.Header = 'End Task'
    $pmi.Add_Click($endTask); $pmenu.Items.Add($pmi) | Out-Null
    $PGrid.ContextMenu = $pmenu

    $pw.Add_Closed({ if ($cim) { try { Remove-CimSession $cim } catch { } } }.GetNewClosure())
    $pw.Add_Loaded($loadProcs)
    $pw.ShowDialog() | Out-Null
}

function Open-Processes {
    $row = Get-SelectedRow
    if (-not $row) { return }
    Show-ProcessWindow -Server $row.Server -SessionId $row.SessionID -Username $row.Username
}

# ----------------------------------------------------------------------------
# SECTION 14: Context menu + double-click on the sessions grid.
# ----------------------------------------------------------------------------
$menu = New-Object System.Windows.Controls.ContextMenu
function New-MenuItem($text, $action) {
    $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = $text; $mi.Add_Click($action); return $mi
}
$menu.Items.Add((New-MenuItem 'View Processes...'  { Open-Processes }))   | Out-Null
$menu.Items.Add((New-Object System.Windows.Controls.Separator))           | Out-Null
$menu.Items.Add((New-MenuItem 'Shadow Session'     { Invoke-Shadow }))    | Out-Null
$menu.Items.Add((New-MenuItem 'Connect (RDP)'      { Invoke-Connect }))   | Out-Null
$menu.Items.Add((New-MenuItem 'Send Message'       { Invoke-Message }))   | Out-Null
$menu.Items.Add((New-Object System.Windows.Controls.Separator))           | Out-Null
$menu.Items.Add((New-MenuItem 'Disconnect'         { Invoke-Disconnect })) | Out-Null
$menu.Items.Add((New-MenuItem 'Log Off Session(s)' { Invoke-Logoff }))    | Out-Null
$menu.Items.Add((New-MenuItem 'Reset Session(s)'   { Invoke-Reset }))     | Out-Null
$Grid.ContextMenu = $menu
$menu.Add_Opened({ if ($Grid.SelectedItems.Count -eq 0) { $menu.IsOpen = $false } })
$Grid.Add_MouseDoubleClick({ Open-Processes })

# ----------------------------------------------------------------------------
# SECTION 15: Theme toggle + auto-refresh timer.
# ----------------------------------------------------------------------------
function Set-DarkMode {
    param([bool]$Dark)
    $script:DarkMode = $Dark
    Apply-Theme -win $window -Dark $Dark
    # Show a sun while dark (click -> light) and a moon while light (click -> dark).
    $BtnDark.Content = if ($Dark) { [char]0x2600 } else { [char]0x263D }
    $s = Read-Settings; $s.DarkMode = [int]$Dark; Write-Settings -s $s; $script:Settings.DarkMode = [int]$Dark
}
$BtnDark.Add_Click({ Set-DarkMode (-not $script:DarkMode) })

$autoTimer = New-Object System.Windows.Threading.DispatcherTimer
$autoTimer.Add_Tick({ if (-not $script:Pool) { Start-Refresh } })
$ChkAuto.Add_Checked({
    $secs = 30; [int]::TryParse($TxtInterval.Text, [ref]$secs) | Out-Null
    if ($secs -lt 5) { $secs = 5 }; if ($secs -gt 3600) { $secs = 3600 }
    $autoTimer.Interval = [TimeSpan]::FromSeconds($secs); $autoTimer.Start()
})
$ChkAuto.Add_Unchecked({ $autoTimer.Stop() })

# ----------------------------------------------------------------------------
# SECTION 16: Settings dialog.
# ----------------------------------------------------------------------------
function Show-Settings {
    $s = Read-Settings
    $dlg = New-Object System.Windows.Window
    $dlg.Title = "Settings"; $dlg.Width = 320; $dlg.Height = 190
    $dlg.WindowStartupLocation = 'CenterOwner'; $dlg.Owner = $window; $dlg.ResizeMode = 'NoResize'
    Apply-Theme -win $dlg -Dark $script:DarkMode
    $dlg.Background = $dlg.Resources['WinBg']
    $fg = $dlg.Resources['Text']
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $cbStart = New-Object System.Windows.Controls.CheckBox; $cbStart.Content = 'Refresh on startup'; $cbStart.Margin = '0,0,0,10'; $cbStart.Foreground = $fg
    $cbStart.IsChecked = [bool]$s.RefreshOnStartup
    $cbShadow = New-Object System.Windows.Controls.CheckBox; $cbShadow.Content = 'Default shadow options (control + no prompt)'; $cbShadow.Margin = '0,0,0,16'; $cbShadow.Foreground = $fg
    $cbShadow.IsChecked = [bool]$s.DefaultShadowOptions
    $btns = New-Object System.Windows.Controls.StackPanel; $btns.Orientation = 'Horizontal'; $btns.HorizontalAlignment = 'Right'
    $ok = New-Object System.Windows.Controls.Button; $ok.Content = 'OK'; $ok.Width = 80; $ok.Margin = '0,0,8,0'; $ok.IsDefault = $true
    $cancel = New-Object System.Windows.Controls.Button; $cancel.Content = 'Cancel'; $cancel.Width = 80; $cancel.IsCancel = $true
    $ok.Add_Click({ $dlg.DialogResult = $true })
    $btns.Children.Add($ok) | Out-Null; $btns.Children.Add($cancel) | Out-Null
    $sp.Children.Add($cbStart) | Out-Null; $sp.Children.Add($cbShadow) | Out-Null; $sp.Children.Add($btns) | Out-Null
    $dlg.Content = $sp
    if ($dlg.ShowDialog()) {
        $s.RefreshOnStartup     = [int][bool]$cbStart.IsChecked
        $s.DefaultShadowOptions = [int][bool]$cbShadow.IsChecked
        $s.AutoRefresh          = [int][bool]$ChkAuto.IsChecked
        $s.DarkMode             = [int]$script:DarkMode
        $iv = 30; [int]::TryParse($TxtInterval.Text, [ref]$iv) | Out-Null; $s.AutoRefreshSeconds = $iv
        Write-Settings -s $s; $script:Settings = $s
    }
}
$BtnSettings.Add_Click({ Show-Settings })

# ----------------------------------------------------------------------------
# SECTION 17: Startup.
# ----------------------------------------------------------------------------
Set-DarkMode $script:DarkMode   # apply saved theme + set toggle label
if ($script:Settings.AutoRefreshSeconds) { $TxtInterval.Text = [string]$script:Settings.AutoRefreshSeconds }
if ($script:Settings.AutoRefresh -eq 1)  { $ChkAuto.IsChecked = $true }

$window.Add_Loaded({ if ($script:Settings.RefreshOnStartup -eq 1) { Start-Refresh } })
$window.Add_Closed({
    if ($autoTimer.IsEnabled) { $autoTimer.Stop() }
    if ($script:PollTimer) { try { $script:PollTimer.Stop() } catch { } }
    if ($script:Pool) { try { $script:Pool.Close(); $script:Pool.Dispose() } catch { } }
})

$window.ShowDialog() | Out-Null
