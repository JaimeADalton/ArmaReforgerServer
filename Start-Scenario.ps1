<#
.SYNOPSIS
  Start-Scenario.ps1 - vBERCON-UX-20260313-FULL
  Gestor directo y cómodo para Arma Reforger usando bercon-cli

.REQUISITOS
  - bercon-cli-windows-amd64.exe en C:\BEAR\Servidores\Arma Reforger\
  - BEServer_x64.cfg en C:\BEAR\Servidores\Arma Reforger\arma_reforger\battleye\
  - BattlEye habilitado en el escenario (battlEye: true)
#>

$ErrorActionPreference = 'Stop'

#region ==================== CONFIGURACIÓN ====================

$script:Config = @{
    Root               = 'C:\BEAR\Servidores\Arma Reforger'
    ProfilePath        = 'C:\BEAR\Servidores\Arma Reforger'
    GamePort           = 2001
    MaxFPS             = 60
    RConTimeoutSec     = 5
    ShutdownTimeoutSec = 45
    RefreshMs          = 200
    StatusRefreshSec   = 5
    ConsoleWidth       = 124
    ConsoleHeight      = 48
    VersionTag         = 'vBERCON-UX-20260313-FULL'
}

$script:Paths = @{
    ServerDir    = Join-Path $Config.Root 'arma_reforger'
    AdminDir     = Join-Path $Config.Root 'administracion'
    ScenariosDir = Join-Path $Config.Root 'escenarios'
    ServerExe    = Join-Path (Join-Path $Config.Root 'arma_reforger') 'ArmaReforgerServer.exe'
    ActiveConfig = Join-Path (Join-Path $Config.Root 'administracion') 'active-config.txt'
    LauncherBat  = Join-Path (Join-Path $Config.Root 'administracion') 'manual_launch.bat'
    BerconCli    = Join-Path $Config.Root 'bercon-cli-windows-amd64.exe'
    BEConfig     = Join-Path (Join-Path (Join-Path $Config.Root 'arma_reforger') 'battleye') 'BEServer_x64.cfg'
    LogsRoot     = Join-Path $Config.Root 'logs'
}

$script:UiState = @{
    Notice       = ''
    NoticeColor  = 'DarkGray'
    NumberBuffer = ''
    Dirty        = $true
}

$script:SwitchState = @{
    TargetConfig = $null
    InProgress   = $false
}

$script:LastFrame         = ''
$script:LastStatusRefresh = [datetime]::MinValue

# Caché de jugadores: se actualiza solo bajo demanda (tecla P) o manualmente,
# NUNCA en el refresco automático del dashboard para no spamear RCon.
$script:PlayerCache = [PSCustomObject]@{
    State      = 'unknown'
    Count      = '-'
    AvgPing    = '-'
    InLobby    = '-'
    ValidCount = '-'
    UpdatedAt  = [datetime]::MinValue
}

#endregion

#region ==================== PALETA DE COLORES ====================

$script:Theme = @{
    HeaderBorder  = 'DarkYellow'
    SectionBorder = 'DarkGray'
    SectionLabel  = 'White'
    Label         = 'DarkGray'
    Value         = 'Gray'
    Accent        = 'Cyan'
    Muted         = 'DarkGray'
    Good          = 'Green'
    Warn          = 'Yellow'
    Bad           = 'Red'
    Info          = 'DarkCyan'
    ScenarioActive   = 'Green'
    ScenarioLast     = 'DarkYellow'
    ScenarioDefault  = 'Gray'
    FolderHeader     = 'DarkCyan'
    KeyHighlight  = 'White'
    KeyLabel      = 'DarkGray'
    StatusBar     = 'DarkGray'
    Clock         = 'DarkGray'
}

#endregion

#region ==================== RENDER BUFFER ====================

$script:FrameBuffer = [System.Collections.Generic.List[object]]::new()

function Buf-Clear  { $script:FrameBuffer.Clear() }

function Buf-Line {
    param([string]$Text = '', [string]$FgColor = 'Gray')
    $script:FrameBuffer.Add(@{ Text = $Text; FgColor = $FgColor })
}

function Buf-LineParts {
    param([array]$Parts)
    $script:FrameBuffer.Add(@{ Parts = $Parts })
}

function Render-Frame {
    <#
      Renderizado atómico: construye TODO el frame como un único string
      con secuencias ANSI/VT100 para color y lo vuelca de golpe con
      [Console]::Write(). Esto evita los cientos de Write-Host que
      causan parpadeo y la sensación de "cortina bajando".
    #>

    # ── Comparar con el frame anterior (texto plano, sin ANSI) ──
    $plainBuilder = [System.Text.StringBuilder]::new()
    foreach ($entry in $script:FrameBuffer) {
        if ($entry.ContainsKey('Parts')) {
            foreach ($p in $entry.Parts) { $plainBuilder.Append($p.Text) | Out-Null }
        } else { $plainBuilder.Append($entry.Text) | Out-Null }
        $plainBuilder.AppendLine() | Out-Null
    }
    $plain = $plainBuilder.ToString()
    if ($plain -eq $script:LastFrame) { return }
    $script:LastFrame = $plain

    # ── Construir frame con colores ANSI en un solo StringBuilder ──
    $width = [Math]::Max([Console]::WindowWidth, 80)
    $esc = [char]0x1B
    $reset = "$esc[0m"

    $sb = [System.Text.StringBuilder]::new(8192)

    # Ocultar cursor + mover a (0,0)
    $sb.Append("$esc[?25l$esc[H") | Out-Null

    foreach ($entry in $script:FrameBuffer) {
        if ($entry.ContainsKey('Parts')) {
            $lineLen = 0
            foreach ($p in $entry.Parts) {
                $ansi = Get-AnsiColor ($p.FgColor)
                $sb.Append($ansi).Append($p.Text) | Out-Null
                $lineLen += $p.Text.Length
            }
            $sb.Append($reset) | Out-Null
            $pad = $width - $lineLen
            if ($pad -gt 0) { $sb.Append(' ' * $pad) | Out-Null }
        }
        else {
            $ansi = Get-AnsiColor ($entry.FgColor)
            $text = $entry.Text
            $pad = $width - $text.Length
            $sb.Append($ansi).Append($text) | Out-Null
            if ($pad -gt 0) { $sb.Append(' ' * $pad) | Out-Null }
            $sb.Append($reset) | Out-Null
        }
        $sb.Append("`n") | Out-Null
    }

    # Limpiar líneas sobrantes del frame anterior
    $remaining = [Console]::WindowHeight - $script:FrameBuffer.Count - 1
    if ($remaining -gt 0) {
        $blank = ' ' * $width
        for ($i = 0; $i -lt $remaining; $i++) {
            $sb.Append($blank).Append("`n") | Out-Null
        }
    }

    # Mostrar cursor
    $sb.Append("$esc[?25h") | Out-Null

    # ── VOLCADO ATÓMICO: una sola escritura ──
    [Console]::Write($sb.ToString())
}

function Get-AnsiColor {
    param([string]$Name)
    $esc = [char]0x1B
    switch ($Name) {
        'Black'       { return "$esc[30m" }
        'DarkRed'     { return "$esc[31m" }
        'DarkGreen'   { return "$esc[32m" }
        'DarkYellow'  { return "$esc[33m" }
        'DarkBlue'    { return "$esc[34m" }
        'DarkMagenta' { return "$esc[35m" }
        'DarkCyan'    { return "$esc[36m" }
        'Gray'        { return "$esc[37m" }
        'DarkGray'    { return "$esc[90m" }
        'Red'         { return "$esc[91m" }
        'Green'       { return "$esc[92m" }
        'Yellow'      { return "$esc[93m" }
        'Blue'        { return "$esc[94m" }
        'Magenta'     { return "$esc[95m" }
        'Cyan'        { return "$esc[96m" }
        'White'       { return "$esc[97m" }
        default       { return "$esc[37m" }
    }
}

#endregion

#region ==================== UI LAYOUT HELPERS ====================

$script:BoxW = 78

function Pad {
    param([string]$Text, [int]$Width)
    if ($Text.Length -ge $Width) { return $Text.Substring(0, $Width) }
    return $Text + (' ' * ($Width - $Text.Length))
}

function UI-HeaderTop {
    $w = $script:BoxW - 2
    Buf-LineParts @(
        @{ Text = '  '; FgColor = 'Black' },
        @{ Text = "╔$('═' * $w)╗"; FgColor = $Theme.HeaderBorder }
    )
}

function UI-HeaderBot {
    $w = $script:BoxW - 2
    Buf-LineParts @(
        @{ Text = '  '; FgColor = 'Black' },
        @{ Text = "╚$('═' * $w)╝"; FgColor = $Theme.HeaderBorder }
    )
}

function UI-HeaderRow {
    param([array]$Parts)
    $textLen = 0; foreach ($p in $Parts) { $textLen += $p.Text.Length }
    $padLen = $script:BoxW - 4 - $textLen
    $all = [System.Collections.Generic.List[object]]::new()
    $all.Add(@{ Text = '  '; FgColor = 'Black' })
    $all.Add(@{ Text = '║ '; FgColor = $Theme.HeaderBorder })
    foreach ($p in $Parts) { $all.Add($p) }
    $all.Add(@{ Text = "$(' ' * [Math]::Max($padLen, 0)) ║"; FgColor = $Theme.HeaderBorder })
    Buf-LineParts $all.ToArray()
}

function UI-SectionTop {
    param([string]$Label = '')
    $bc = $Theme.SectionBorder; $w = $script:BoxW - 2
    if ($Label) {
        $tag = " $Label "
        $lineLen = $w - 2 - $tag.Length
        Buf-LineParts @(
            @{ Text = '  '; FgColor = 'Black' },
            @{ Text = '┌─'; FgColor = $bc },
            @{ Text = $tag; FgColor = $Theme.SectionLabel },
            @{ Text = "$('─' * [Math]::Max($lineLen, 0))┐"; FgColor = $bc }
        )
    } else {
        Buf-LineParts @(
            @{ Text = '  '; FgColor = 'Black' },
            @{ Text = "┌$('─' * $w)┐"; FgColor = $bc }
        )
    }
}

function UI-SectionBot {
    $w = $script:BoxW - 2
    Buf-LineParts @(
        @{ Text = '  '; FgColor = 'Black' },
        @{ Text = "└$('─' * $w)┘"; FgColor = $Theme.SectionBorder }
    )
}

function UI-Row {
    param([array]$Parts)
    $bc = $Theme.SectionBorder
    $textLen = 0; foreach ($p in $Parts) { $textLen += $p.Text.Length }
    $padLen = $script:BoxW - 5 - $textLen
    $all = [System.Collections.Generic.List[object]]::new()
    $all.Add(@{ Text = '  '; FgColor = 'Black' })
    $all.Add(@{ Text = '│  '; FgColor = $bc })
    foreach ($p in $Parts) { $all.Add($p) }
    $all.Add(@{ Text = "$(' ' * [Math]::Max($padLen, 0)) │"; FgColor = $bc })
    Buf-LineParts $all.ToArray()
}

function UI-EmptyRow {
    $s = $script:BoxW - 2
    Buf-LineParts @(
        @{ Text = '  '; FgColor = 'Black' },
        @{ Text = "│$(' ' * $s)│"; FgColor = $Theme.SectionBorder }
    )
}

#endregion

#region ==================== UTILIDADES BASE ====================

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Force -Path $Path | Out-Null }
}

function Write-FileNoBOM {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
}

function Set-Notice {
    param([Parameter(Mandatory)][string]$Message, [string]$Color = 'Yellow')
    $script:UiState.Notice = $Message
    $script:UiState.NoticeColor = $Color
    $script:UiState.Dirty = $true
}

function Clear-Notice {
    $script:UiState.Notice = ''
    $script:UiState.NoticeColor = 'DarkGray'
    $script:UiState.Dirty = $true
}

function Mark-Dirty { $script:UiState.Dirty = $true }

function Get-ActiveScenario {
    if (-not (Test-Path $Paths.ActiveConfig)) { return $null }
    $raw = Get-Content $Paths.ActiveConfig -Raw
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw.Trim()
}

function Set-ActiveScenario {
    param([Parameter(Mandatory)][string]$ConfigPath)
    Write-FileNoBOM -Path $Paths.ActiveConfig -Content $ConfigPath
}

function Get-ServerProcess {
    Get-Process -Name 'ArmaReforgerServer' -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Get-LatestLogDir {
    if (-not (Test-Path $Paths.LogsRoot)) { return $null }
    Get-ChildItem $Paths.LogsRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Format-Uptime {
    param($Uptime)
    if ($null -eq $Uptime -or $Uptime -isnot [System.TimeSpan]) { return '-' }
    return "{0:00}d {1:00}h {2:00}m {3:00}s" -f $Uptime.Days, $Uptime.Hours, $Uptime.Minutes, $Uptime.Seconds
}

function Get-ServerStatus {
    $proc = Get-ServerProcess
    $activeConfig = Get-ActiveScenario
    $latestLog = Get-LatestLogDir
    $status = [ordered]@{
        Running = $false; Pid = $null; PortOpen = $false; ActiveConfig = $activeConfig
        Uptime = $null
        LatestLogName      = if ($latestLog) { $latestLog.Name } else { $null }
        ActiveScenarioName = if ($activeConfig) { [IO.Path]::GetFileName($activeConfig) } else { $null }
    }
    if ($proc) {
        $status.Running = $true; $status.Pid = $proc.Id

        # Check de puerto cacheado: Get-NetUDPEndpoint es lento, solo cada 30s
        $now = Get-Date
        if ($null -eq $script:_PortCache -or ($now - $script:_PortCacheTime).TotalSeconds -ge 30 -or $script:_PortCachePid -ne $proc.Id) {
            try {
                $udp = Get-NetUDPEndpoint -OwningProcess $proc.Id -ErrorAction SilentlyContinue |
                       Where-Object { $_.LocalPort -eq $Config.GamePort }
                $script:_PortCache = [bool]$udp
            } catch { $script:_PortCache = $false }
            $script:_PortCacheTime = $now
            $script:_PortCachePid = $proc.Id
        }
        $status.PortOpen = $script:_PortCache

        try { $status.Uptime = (Get-Date) - $proc.StartTime } catch { }
    } else {
        $script:_PortCache = $false
    }
    return [PSCustomObject]$status
}

function Get-AllScenarios {
    $items = @(); $index = 1
    if (-not (Test-Path $Paths.ScenariosDir)) { return $items }
    $files = Get-ChildItem -Path $Paths.ScenariosDir -Filter '*.json' -File -Recurse | Sort-Object FullName
    foreach ($file in $files) {
        $folder = $file.DirectoryName.Substring($Paths.ScenariosDir.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($folder)) { $folder = '(raíz)' }
        $items += [PSCustomObject]@{ Index = $index; Name = $file.BaseName; FullPath = $file.FullName; Folder = $folder }
        $index++
    }
    return $items
}

function Get-ScenarioByIndex {
    param([int]$Index)
    $list = Get-AllScenarios
    return ($list | Where-Object { $_.Index -eq $Index } | Select-Object -First 1)
}

function Assert-Prerequisites {
    $missing = @()
    if (-not (Test-Path $Paths.ServerExe)) { $missing += $Paths.ServerExe }
    if (-not (Test-Path $Paths.BerconCli)) { $missing += $Paths.BerconCli }
    if (-not (Test-Path $Paths.BEConfig))  { $missing += $Paths.BEConfig }
    if ($missing.Count -gt 0) { throw "Faltan rutas obligatorias:`n - " + ($missing -join "`n - ") }
    Ensure-Directory -Path $Paths.AdminDir
}

function Prompt-Input {
    param([Parameter(Mandatory)][string]$Label)
    Write-Host ""
    $value = Read-Host $Label
    $script:LastFrame = ''; Mark-Dirty
    return $value
}

function Wait-AnyKey {
    Write-Host ""
    Write-Host "  Pulsa cualquier tecla para volver..." -ForegroundColor DarkGray
    [Console]::ReadKey($true) | Out-Null
    $script:LastFrame = ''; Mark-Dirty
}

#endregion

#region ==================== BERCON-CLI ====================

function Invoke-Bercon {
    param(
        [Parameter(Mandatory)][string[]]$Commands,
        [ValidateSet('table','json','raw','md','html')][string]$Format = 'table',
        [int]$TimeoutSec = 5
    )
    $berconArgs = @('-r',$Paths.BEConfig,'-t',$TimeoutSec.ToString(),'-f',$Format,'--') + $Commands
    $oldPref = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = & $Paths.BerconCli @berconArgs 2>&1
        $exitCode = $LASTEXITCODE
        $text = (($output | ForEach-Object { "$_" }) -join "`n").Trim()
        return [PSCustomObject]@{ Success = ($exitCode -eq 0); ExitCode = $exitCode; Output = $text }
    } catch {
        return [PSCustomObject]@{ Success = $false; ExitCode = -1; Output = $_.Exception.Message }
    } finally { $ErrorActionPreference = $oldPref }
}

function Get-PlayersData {
    <# UNA sola llamada RCon. Cachea el resultado en $script:_PlayersDataCache. #>
    $server = Get-ServerProcess
    if (-not $server) {
        $script:_PlayersDataCache = [PSCustomObject]@{ State = 'stopped'; Players = @(); Error = $null }
        return $script:_PlayersDataCache
    }
    $result = Invoke-Bercon -Commands @('players') -Format 'json' -TimeoutSec $Config.RConTimeoutSec
    if (-not $result.Success) {
        $script:_PlayersDataCache = [PSCustomObject]@{ State = 'warming'; Players = @(); Error = $result.Output }
        return $script:_PlayersDataCache
    }
    try {
        $players = $result.Output | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $players) { $players = @() }
        if ($players -isnot [System.Array]) { $players = @($players) }
        $script:_PlayersDataCache = [PSCustomObject]@{ State = 'ok'; Players = @($players); Error = $null }
    } catch {
        $script:_PlayersDataCache = [PSCustomObject]@{ State = 'warming'; Players = @(); Error = 'No se pudo interpretar la respuesta JSON.' }
    }
    return $script:_PlayersDataCache
}

function Update-PlayerCacheFromData {
    <# Actualiza la caché del dashboard a partir de datos ya obtenidos. SIN llamadas RCon. #>
    param([PSCustomObject]$Data)

    $summary = switch ($Data.State) {
        'stopped' { [PSCustomObject]@{ State='stopped'; Count='-'; AvgPing='-'; InLobby='-'; ValidCount='-' } }
        'warming' { [PSCustomObject]@{ State='warming'; Count='-'; AvgPing='-'; InLobby='-'; ValidCount='-' } }
        'ok' {
            $players = @($Data.Players); $count = $players.Count
            $avgPing = '-'; $inLobby = 0; $validCount = 0
            if ($count -gt 0) {
                $pings = @($players | Where-Object { $_.ping -ne $null } | ForEach-Object { [double]$_.ping })
                if ($pings.Count -gt 0) { $avgPing = [math]::Round(($pings | Measure-Object -Average).Average, 0) }
                $inLobby = @($players | Where-Object { $_.lobby -eq $true }).Count
                $validCount = @($players | Where-Object { $_.valid -eq $true }).Count
            }
            [PSCustomObject]@{ State='ok'; Count=$count; AvgPing=$avgPing; InLobby=$inLobby; ValidCount=$validCount }
        }
    }

    $script:PlayerCache = [PSCustomObject]@{
        State      = $summary.State
        Count      = $summary.Count
        AvgPing    = $summary.AvgPing
        InLobby    = $summary.InLobby
        ValidCount = $summary.ValidCount
        UpdatedAt  = Get-Date
    }
}

function Get-PlayerSummary {
    <# UNA llamada RCon + actualiza caché. Devuelve el resumen. #>
    $data = Get-PlayersData              # 1 llamada RCon
    Update-PlayerCacheFromData -Data $data   # solo math local
    return $script:PlayerCache
}

function Get-CachedPlayerSummary {
    <# Devuelve la caché SIN ninguna llamada RCon. Usado por el refresco automático. #>
    $proc = Get-ServerProcess
    if (-not $proc) {
        return [PSCustomObject]@{ State='stopped'; Count='-'; AvgPing='-'; InLobby='-'; ValidCount='-' }
    }
    return $script:PlayerCache
}

#endregion

#region ==================== PANTALLAS MODALES ====================

function Show-PlayersScreen {
    # UNA sola llamada RCon — Get-PlayersData cachea internamente
    $data = Get-PlayersData
    Update-PlayerCacheFromData -Data $data   # actualiza caché del dashboard sin RCon extra
    Clear-Host; Write-Host ""
    $w = 72; $inner = $w - 2; $tag = ' JUGADORES CONECTADOS '
    $lineLen = $inner - 2 - $tag.Length
    Write-Host "  ┌─" -NoNewline -ForegroundColor $Theme.SectionBorder
    Write-Host $tag -NoNewline -ForegroundColor $Theme.SectionLabel
    Write-Host "$('─' * [Math]::Max($lineLen,0))┐" -ForegroundColor $Theme.SectionBorder

    switch ($data.State) {
        'stopped' { Write-Host "`n    El servidor está detenido." -ForegroundColor $Theme.Muted }
        'warming' {
            Write-Host "`n    RCon todavía no responde — puede tardar unos segundos." -ForegroundColor $Theme.Warn
            if ($data.Error) { Write-Host "    $($data.Error)" -ForegroundColor $Theme.Muted }
        }
        'ok' {
            $players = @($data.Players)
            if ($players.Count -eq 0) { Write-Host "`n    0 jugadores conectados." -ForegroundColor $Theme.Good }
            else {
                $display = $players | Select-Object `
                    @{Name='ID';Expression={$_.id}}, @{Name='Nombre';Expression={$_.name}},
                    @{Name='Ping';Expression={$_.ping}}, @{Name='Válido';Expression={$_.valid}},
                    @{Name='Lobby';Expression={$_.lobby}}, @{Name='IP';Expression={$_.ip}},
                    @{Name='Puerto';Expression={$_.port}}
                Write-Host ""; Write-Host ($display | Format-Table -AutoSize | Out-String)
            }
        }
    }
    Wait-AnyKey
}

function Kick-PlayerInteractive {
    $data = Get-PlayersData
    Update-PlayerCacheFromData -Data $data
    if ($data.State -eq 'stopped') { Set-Notice "No hay servidor en ejecución." 'DarkGray'; return }
    if ($data.State -eq 'warming') { Set-Notice "RCon todavía no responde." 'Yellow'; return }
    $players = @($data.Players)
    if ($players.Count -eq 0) { Set-Notice "No hay jugadores conectados." 'DarkGray'; return }

    Clear-Host; Write-Host ""
    $w = 72; $inner = $w - 2; $tag = ' EXPULSAR JUGADOR '
    $lineLen = $inner - 2 - $tag.Length
    Write-Host "  ┌─" -NoNewline -ForegroundColor $Theme.SectionBorder
    Write-Host $tag -NoNewline -ForegroundColor $Theme.SectionLabel
    Write-Host "$('─' * [Math]::Max($lineLen,0))┐" -ForegroundColor $Theme.SectionBorder

    $display = $players | Select-Object `
        @{Name='ID';Expression={$_.id}}, @{Name='Nombre';Expression={$_.name}},
        @{Name='Ping';Expression={$_.ping}}, @{Name='Lobby';Expression={$_.lobby}}
    Write-Host ""; Write-Host ($display | Format-Table -AutoSize | Out-String)

    $id = Prompt-Input "  ID del jugador a expulsar"
    if ([string]::IsNullOrWhiteSpace($id)) { Set-Notice "Kick cancelado." 'DarkGray'; return }
    if ($id -notmatch '^\d+$') { Set-Notice "El ID debe ser numérico." 'Red'; return }
    $confirm = Prompt-Input "  Escribe SI para confirmar el kick del playerId $id"
    if ($confirm -ne 'SI') { Set-Notice "Kick cancelado." 'DarkGray'; return }

    $result = Invoke-Bercon -Commands @("kick $id") -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) { Set-Notice "Kick enviado para playerId $id." 'Green' }
    else { Set-Notice "Falló el kick: $($result.Output)" 'Red' }
}

function Send-ServerMessage {
    $msg = Prompt-Input "  Mensaje global"
    if ([string]::IsNullOrWhiteSpace($msg)) { Set-Notice "Mensaje cancelado." 'DarkGray'; return }
    $result = Invoke-Bercon -Commands @("say -1 $msg") -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) { Set-Notice "Mensaje global enviado." 'Green' }
    else { Set-Notice "No se pudo enviar: $($result.Output)" 'Red' }
}

function Lock-Server {
    $result = Invoke-Bercon -Commands @('#lock') -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) { Set-Notice "Servidor bloqueado." 'Green' }
    else { Set-Notice "No se pudo bloquear: $($result.Output)" 'Red' }
}

function Unlock-Server {
    $result = Invoke-Bercon -Commands @('#unlock') -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) { Set-Notice "Servidor desbloqueado." 'Green' }
    else { Set-Notice "No se pudo desbloquear: $($result.Output)" 'Red' }
}

function Send-CustomRConCommand {
    $cmd = Prompt-Input "  Comando RCon"
    if ([string]::IsNullOrWhiteSpace($cmd)) { Set-Notice "Comando cancelado." 'DarkGray'; return }
    $result = Invoke-Bercon -Commands @($cmd) -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        if ($result.Output) { Set-Notice "OK: $($result.Output)" 'Green' } else { Set-Notice "Comando enviado." 'Green' }
    } else { Set-Notice "Falló: $($result.Output)" 'Red' }
}

#endregion

#region ==================== CAMBIO DE ESCENARIO / PARADA ====================

function Clear-SwitchState {
    $script:SwitchState.TargetConfig = $null; $script:SwitchState.InProgress = $false
}

function Set-SwitchTarget {
    param([Parameter(Mandatory)][string]$ConfigPath)
    $script:SwitchState.TargetConfig = $ConfigPath; $script:SwitchState.InProgress = $true
}

function Get-SwitchTarget { return $script:SwitchState.TargetConfig }

function Write-LauncherBat {
    param([Parameter(Mandatory)][string]$ConfigPath)
    $scenarioFile = [IO.Path]::GetFileName($ConfigPath)
    $batContent = @"
@echo off
title ARMA REFORGER - SERVIDOR ACTIVO
cd /d "$($Paths.ServerDir)"
echo.
echo  ==================================================================
echo   INICIANDO SERVIDOR: $scenarioFile
echo  ==================================================================
echo.
"ArmaReforgerServer.exe" -config "$ConfigPath" -loadSessionSave -profile "$($Config.ProfilePath)" -maxFPS $($Config.MaxFPS) -download -nothrow
echo.
echo  ==================================================================
echo   EL SERVIDOR SE HA CERRADO.
echo  ==================================================================
pause
"@
    Write-FileNoBOM -Path $Paths.LauncherBat -Content $batContent
}

function Start-ServerDirect {
    param([string]$ConfigPath = $null)
    $targetConfig = $ConfigPath
    if ([string]::IsNullOrWhiteSpace($targetConfig)) { $targetConfig = Get-ActiveScenario }
    if (-not $targetConfig) { Set-Notice "No hay escenario seleccionado." 'Yellow'; return }
    if (-not (Test-Path $targetConfig)) { Set-Notice "JSON no encontrado: $targetConfig" 'Red'; return }
    $serverProc = Get-ServerProcess
    if ($serverProc) {
        $current = Get-ActiveScenario
        if ($current -and ($current -eq $targetConfig)) { Set-Notice "Ese escenario ya está en ejecución." 'Yellow' }
        else {
            Set-SwitchTarget -ConfigPath $targetConfig
            Set-Notice "Cambiando escenario — apagando servidor..." 'Cyan'
            Stop-ServerGraceful
        }
        return
    }
    Ensure-Directory -Path $Paths.AdminDir
    Write-LauncherBat -ConfigPath $targetConfig
    Start-Process -FilePath $Paths.LauncherBat | Out-Null
    Set-ActiveScenario -ConfigPath $targetConfig
    Clear-SwitchState
    Set-Notice "Servidor lanzado: $([IO.Path]::GetFileName($targetConfig))" 'Green'
    Start-Sleep -Seconds 2
}

function Stop-ServerGraceful {
    $serverProc = Get-ServerProcess
    if (-not $serverProc) {
        $switchTarget = Get-SwitchTarget
        if ($script:SwitchState.InProgress -and $switchTarget -and (Test-Path $switchTarget)) {
            Start-ServerDirect -ConfigPath $switchTarget; return
        }
        Set-Notice "No hay servidor en ejecución." 'DarkGray'; return
    }
    Set-Notice "Enviando #shutdown..." 'Cyan'
    $result = Invoke-Bercon -Commands @('#shutdown') -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if (-not $result.Success) {
        Set-Notice "No se pudo enviar #shutdown: $($result.Output)" 'Red'; Clear-SwitchState; return
    }
    $deadline = (Get-Date).AddSeconds($Config.ShutdownTimeoutSec)
    do { Start-Sleep -Milliseconds 700; $serverProc = Get-ServerProcess
    } while ($serverProc -and ((Get-Date) -lt $deadline))
    if (-not $serverProc) {
        $switchTarget = Get-SwitchTarget
        if ($script:SwitchState.InProgress -and $switchTarget -and (Test-Path $switchTarget)) {
            Set-Notice "Servidor detenido — lanzando nuevo escenario..." 'Green'
            Start-Sleep -Seconds 1; Start-ServerDirect -ConfigPath $switchTarget
        } else { Set-Notice "Servidor detenido correctamente." 'Green' }
    } else { Set-Notice "Comando enviado, pero el proceso sigue vivo." 'Yellow'; Clear-SwitchState }
}

#endregion

#region ==================== BUILD-FRAME ====================

function Build-Frame {
    param([PSCustomObject]$Status, [PSCustomObject]$Stats)

    Buf-Clear

    # ── Ancho dinámico ──
    $consoleW = [Math]::Max([Console]::WindowWidth, 80)
    $script:BoxW = [Math]::Min($consoleW - 2, 120)
    $contentW = $script:BoxW - 6
    $labelW   = 14
    $leftCol  = [Math]::Floor($contentW * 0.55)

    # ══════════════════════════════════════════════════════════════
    #  CABECERA
    # ══════════════════════════════════════════════════════════════

    Buf-Line ''
    UI-HeaderTop

    $tL = 'ARMA REFORGER  ───  GESTOR BERCON'
    $tR = $Config.VersionTag
    $gap = $contentW - $tL.Length - $tR.Length
    if ($gap -lt 1) { $gap = 1 }

    UI-HeaderRow @(
        @{ Text = $tL;             FgColor = 'White' },
        @{ Text = (' ' * $gap);    FgColor = 'Black' },
        @{ Text = $tR;             FgColor = $Theme.Muted }
    )

    UI-HeaderBot
    Buf-Line ''

    # ══════════════════════════════════════════════════════════════
    #  ESTADO DEL SERVIDOR
    # ══════════════════════════════════════════════════════════════

    UI-SectionTop -Label 'ESTADO'
    UI-EmptyRow

    # Fila 1: Servidor + Uptime
    $srvDot   = if ($Status.Running) { '●' } else { '○' }
    $srvText  = if ($Status.Running) { "EN EJECUCIÓN  ·  PID $($Status.Pid)" } else { 'DETENIDO' }
    $srvColor = if ($Status.Running) { $Theme.Good } else { $Theme.Bad }

    UI-Row @(
        @{ Text = (Pad 'Servidor' $labelW);  FgColor = $Theme.Label },
        @{ Text = (Pad "$srvDot $srvText" ($leftCol - $labelW)); FgColor = $srvColor },
        @{ Text = (Pad 'Uptime' 10);         FgColor = $Theme.Label },
        @{ Text = (Format-Uptime $Status.Uptime); FgColor = $Theme.Accent }
    )

    # Fila 2: Puerto + RCon
    $portDot   = if ($Status.PortOpen) { '●' } else { '○' }
    $portText  = if ($Status.PortOpen) { "UDP $($Config.GamePort) ABIERTO" } else { "UDP $($Config.GamePort) CERRADO" }
    $portColor = if ($Status.PortOpen) { $Theme.Good } else { $Theme.Muted }

    $rconDot = '○'; $rconText = '-'; $rconColor = $Theme.Muted
    switch ($Stats.State) {
        'ok'      { $rconDot = '●'; $rconText = 'OK';         $rconColor = $Theme.Good }
        'warming' { $rconDot = '◌'; $rconText = 'ARRANCANDO'; $rconColor = $Theme.Warn }
        'unknown' { $rconDot = '○'; $rconText = 'Pulsa P';    $rconColor = $Theme.Muted }
    }

    UI-Row @(
        @{ Text = (Pad 'Puerto' $labelW);    FgColor = $Theme.Label },
        @{ Text = (Pad "$portDot $portText" ($leftCol - $labelW)); FgColor = $portColor },
        @{ Text = (Pad 'RCon' 10);           FgColor = $Theme.Label },
        @{ Text = "$rconDot $rconText";      FgColor = $rconColor }
    )

    # Fila 3: Jugadores + Ping
    $plCount = "$($Stats.Count)"
    $plExtra = ''
    $plColor = $Theme.Muted
    if ($Stats.State -eq 'ok') {
        $plColor = $Theme.Good
        if ($Stats.Count -ne '-' -and $Stats.Count -gt 0) {
            $plExtra = "  ·  $($Stats.InLobby) en lobby"
        }
    } elseif ($Stats.State -eq 'unknown') {
        $plCount = '-'
    }
    $plText = if ($Stats.State -eq 'unknown') { 'Pulsa P para consultar' } else { "$plCount conectados$plExtra" }
    $pingText = if ($Stats.AvgPing -eq '-') { '-' } else { "$($Stats.AvgPing) ms avg" }

    UI-Row @(
        @{ Text = (Pad 'Jugadores' $labelW); FgColor = $Theme.Label },
        @{ Text = (Pad $plText ($leftCol - $labelW)); FgColor = $plColor },
        @{ Text = (Pad 'Ping' 10);           FgColor = $Theme.Label },
        @{ Text = $pingText;                 FgColor = $Theme.Accent }
    )

    UI-EmptyRow

    # Fila 4: Misión
    $mLabel = 'Misión'; $mText = '-'; $mColor = $Theme.Muted
    if ($Status.Running -and $Status.ActiveScenarioName) {
        $mText = $Status.ActiveScenarioName; $mColor = $Theme.Warn
    } elseif ($Status.ActiveScenarioName) {
        $mLabel = 'Últ. misión'; $mText = $Status.ActiveScenarioName; $mColor = $Theme.ScenarioLast
    }
    UI-Row @(
        @{ Text = (Pad $mLabel $labelW); FgColor = $Theme.Label },
        @{ Text = $mText;                FgColor = $mColor }
    )

    # Fila 5: Profile
    UI-Row @(
        @{ Text = (Pad 'Profile' $labelW); FgColor = $Theme.Label },
        @{ Text = $Config.ProfilePath;     FgColor = $Theme.Info }
    )

    # Fila 6: Log
    if ($Status.LatestLogName) {
        UI-Row @(
            @{ Text = (Pad 'Log' $labelW);     FgColor = $Theme.Label },
            @{ Text = $Status.LatestLogName;    FgColor = $Theme.Info }
        )
    }

    UI-EmptyRow
    UI-SectionBot

    # ══════════════════════════════════════════════════════════════
    #  ESCENARIOS
    # ══════════════════════════════════════════════════════════════

    Buf-Line ''
    UI-SectionTop -Label 'ESCENARIOS'

    $list = Get-AllScenarios

    if ($list.Count -eq 0) {
        UI-EmptyRow
        UI-Row @( @{ Text = "No se han encontrado JSON en $($Paths.ScenariosDir)"; FgColor = $Theme.Muted } )
        UI-EmptyRow
    } else {
        $currentFolder = ''
        foreach ($s in $list) {
            if ($s.Folder -ne $currentFolder) {
                $currentFolder = $s.Folder
                UI-EmptyRow
                $folderTag = " $currentFolder "
                $folderLine = '─' * [Math]::Min(30, [Math]::Max(1, 36 - $folderTag.Length))
                UI-Row @( @{ Text = "──$folderTag$folderLine"; FgColor = $Theme.FolderHeader } )
            }

            $isSelected = ($Status.ActiveConfig -and ($s.FullPath -eq $Status.ActiveConfig))
            $marker = '  '; $color = $Theme.ScenarioDefault

            if ($Status.Running -and $isSelected) {
                $marker = '▶ '; $color = $Theme.ScenarioActive
            } elseif ((-not $Status.Running) -and $isSelected) {
                $marker = '· '; $color = $Theme.ScenarioLast
            }

            UI-Row @(
                @{ Text = "  $marker";          FgColor = $color },
                @{ Text = "$($s.Index.ToString('00'))  "; FgColor = $Theme.Muted },
                @{ Text = $s.Name;              FgColor = $color }
            )
        }
    }

    UI-EmptyRow
    UI-SectionBot

    # ══════════════════════════════════════════════════════════════
    #  ACCIONES
    # ══════════════════════════════════════════════════════════════

    Buf-Line ''
    UI-SectionTop -Label 'ACCIONES'
    UI-EmptyRow

    # Cuadrícula de acciones: 4 columnas por fila, ancho uniforme
    $colW = [Math]::Floor(($contentW - 4) / 4)  # 4 columnas con margen

    $actions = @(
        @( @('nº+↵','Lanzar'),    @('S','Apagar'),      @('M','Mensaje'),    @('L','Bloquear')    ),
        @( @('U','Desbloquear'),   @('K','Kick'),        @('P','Jugadores'),  @('C','RCon')        ),
        @( @('R','Limpiar aviso'), @('Q','Salir'),       $null,               $null                )
    )

    foreach ($row in $actions) {
        $parts = [System.Collections.Generic.List[object]]::new()
        foreach ($action in $row) {
            if ($null -eq $action) {
                $parts.Add(@{ Text = (Pad '' $colW); FgColor = 'Black' })
            } else {
                $key   = $action[0]
                $label = $action[1]
                $cell  = "$key  $label"
                $parts.Add(@{ Text = $key;  FgColor = $Theme.KeyHighlight })
                $parts.Add(@{ Text = (Pad "  $label" ($colW - $key.Length)); FgColor = $Theme.KeyLabel })
            }
        }
        UI-Row $parts.ToArray()
    }

    UI-EmptyRow

    # --- Selección numérica ---
    $selText  = if ([string]::IsNullOrWhiteSpace($script:UiState.NumberBuffer)) { '▌' } else { "$($script:UiState.NumberBuffer)▌" }
    $selColor = if ([string]::IsNullOrWhiteSpace($script:UiState.NumberBuffer)) { $Theme.Muted } else { $Theme.Accent }

    UI-Row @(
        @{ Text = (Pad 'Selección:' 14); FgColor = $Theme.Label },
        @{ Text = $selText;              FgColor = $selColor }
    )

    UI-EmptyRow
    UI-SectionBot

    # ══════════════════════════════════════════════════════════════
    #  BARRA DE ESTADO
    # ══════════════════════════════════════════════════════════════

    Buf-Line ''

    $clock = (Get-Date).ToString('HH:mm:ss')
    $noticeText  = $script:UiState.Notice
    $noticeColor = if ([string]::IsNullOrWhiteSpace($noticeText)) { $Theme.Muted } else { $script:UiState.NoticeColor }

    $barW = $script:BoxW - 6
    $noticeMaxW = $barW - $clock.Length - 3
    if ($noticeText.Length -gt $noticeMaxW) {
        $noticeText = $noticeText.Substring(0, [Math]::Max($noticeMaxW - 1, 0)) + '…'
    }

    $midGap = $barW - $noticeText.Length - $clock.Length
    if ($midGap -lt 1) { $midGap = 1 }

    if ([string]::IsNullOrWhiteSpace($noticeText)) {
        Buf-LineParts @(
            @{ Text = "  ▐ "; FgColor = $Theme.StatusBar },
            @{ Text = (' ' * ($barW - $clock.Length)); FgColor = 'Black' },
            @{ Text = $clock; FgColor = $Theme.Clock },
            @{ Text = ' ▌'; FgColor = $Theme.StatusBar }
        )
    } else {
        Buf-LineParts @(
            @{ Text = "  ▐ "; FgColor = $Theme.StatusBar },
            @{ Text = $noticeText; FgColor = $noticeColor },
            @{ Text = (' ' * $midGap); FgColor = 'Black' },
            @{ Text = $clock; FgColor = $Theme.Clock },
            @{ Text = ' ▌'; FgColor = $Theme.StatusBar }
        )
    }

    Buf-Line ''
}

#endregion

#region ==================== LOOP DE ENTRADA ====================

function Handle-Key {
    param([System.ConsoleKeyInfo]$KeyInfo)
    $key = $KeyInfo.Key; $char = $KeyInfo.KeyChar

    switch ($key) {
        'Enter' {
            if (-not [string]::IsNullOrWhiteSpace($script:UiState.NumberBuffer)) {
                $num = 0
                if ([int]::TryParse($script:UiState.NumberBuffer, [ref]$num)) {
                    $item = Get-ScenarioByIndex -Index $num
                    if ($item) { Start-ServerDirect -ConfigPath $item.FullPath }
                    else { Set-Notice "No existe ese índice." 'Red' }
                } else { Set-Notice "Selección inválida." 'Red' }
                $script:UiState.NumberBuffer = ''
            }
            Mark-Dirty; return
        }
        'Backspace' {
            if ($script:UiState.NumberBuffer.Length -gt 0) {
                $script:UiState.NumberBuffer = $script:UiState.NumberBuffer.Substring(0, $script:UiState.NumberBuffer.Length - 1)
                Mark-Dirty
            }; return
        }
        'Escape' { $script:UiState.NumberBuffer = ''; Clear-Notice; Mark-Dirty; return }
        'S' { Stop-ServerGraceful; return }
        'M' { Send-ServerMessage; return }
        'L' { Lock-Server; return }
        'U' { Unlock-Server; return }
        'K' { Kick-PlayerInteractive; return }
        'P' { Show-PlayersScreen; return }
        'C' { Send-CustomRConCommand; return }
        'R' { Clear-Notice; return }
        'Q' { throw [System.OperationCanceledException]::new("Salir") }
    }

    if ($char -match '^\d$') { $script:UiState.NumberBuffer += [string]$char; Mark-Dirty }
}

#endregion

#region ==================== MAIN ====================

Assert-Prerequisites
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

# ── Habilitar secuencias VT100 en la consola de Windows ──
# Windows PowerShell 5.1 no las activa por defecto.
# Se necesita SetConsoleMode con ENABLE_VIRTUAL_TERMINAL_PROCESSING (0x0004).
try {
    $VtHelper = Add-Type -PassThru -Name VtHelper -Namespace Win32 -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern IntPtr GetStdHandle(int nStdHandle);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
'@
    $hOut = $VtHelper::GetStdHandle(-11)  # STD_OUTPUT_HANDLE
    $mode = 0
    $null = $VtHelper::GetConsoleMode($hOut, [ref]$mode)
    $null = $VtHelper::SetConsoleMode($hOut, $mode -bor 0x0004)
} catch {
    # Si falla (entorno raro), continuamos — los colores no se verán pero no crashea
}

# ── Forzar tamaño de ventana 124×48 y título ──
try {
    $host.UI.RawUI.WindowTitle = "ARMA REFORGER — Gestor BerCon $($Config.VersionTag)"

    $maxSize = $host.UI.RawUI.MaxPhysicalWindowSize
    $targetW = [Math]::Min($Config.ConsoleWidth, $maxSize.Width)
    $targetH = [Math]::Min($Config.ConsoleHeight, $maxSize.Height)

    # El buffer debe ser >= que la ventana
    $bufSize = $host.UI.RawUI.BufferSize
    if ($bufSize.Width -lt $targetW)  { $bufSize.Width = $targetW }
    if ($bufSize.Height -lt $targetH) { $bufSize.Height = $targetH }
    $host.UI.RawUI.BufferSize = $bufSize

    $winSize = New-Object System.Management.Automation.Host.Size($targetW, $targetH)
    $host.UI.RawUI.WindowSize = $winSize

    # Ajustar buffer al mismo ancho para evitar scroll horizontal
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($targetW, $bufSize.Height)
} catch {
    # Si no se puede redimensionar (terminal que no lo soporte), seguimos
}

Clear-Host

try {
    while ($true) {
        $now = Get-Date
        if (($now - $script:LastStatusRefresh).TotalSeconds -ge $Config.StatusRefreshSec) {
            $script:LastStatusRefresh = $now; Mark-Dirty
        }
        if ($script:UiState.Dirty) {
            $status = Get-ServerStatus; $stats = Get-CachedPlayerSummary
            Build-Frame -Status $status -Stats $stats
            Render-Frame
            $script:UiState.Dirty = $false
        }
        $pollUntil = (Get-Date).AddMilliseconds($Config.RefreshMs)
        while ((Get-Date) -lt $pollUntil) {
            if ([Console]::KeyAvailable) {
                $keyInfo = [Console]::ReadKey($true); Handle-Key -KeyInfo $keyInfo; break
            }
            Start-Sleep -Milliseconds 50
        }
    }
} catch [System.OperationCanceledException] {
    # Salida limpia: reset ANSI, mostrar cursor, limpiar pantalla
    [Console]::Write("$([char]0x1B)[0m$([char]0x1B)[?25h")
    Clear-Host
}

#endregion
