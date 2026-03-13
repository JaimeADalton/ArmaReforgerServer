<#
.SYNOPSIS
  Start-Scenario.ps1 - vBERCON-UX-20260313
  Gestor directo y cómodo para Arma Reforger usando bercon-cli

.REQUISITOS
  - bercon-cli-windows-amd64.exe en C:\BEAR\Servidores\Arma Reforger\
  - BEServer_x64.cfg en C:\BEAR\Servidores\Arma Reforger\arma_reforger\battleye\
  - BattlEye habilitado en el escenario (battlEye: true)
#>

$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

#region ==================== CONFIGURACIÓN ====================

$script:Config = @{
    Root               = 'C:\BEAR\Servidores\Arma Reforger'
    ProfilePath        = 'C:\BEAR\Servidores\Arma Reforger'
    GamePort           = 2001
    MaxFPS             = 60
    RConTimeoutSec     = 5
    ShutdownTimeoutSec = 45
    VersionTag         = 'vBERCON-UX-20260313'
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
    Notice      = ''
    NoticeColor = 'DarkGray'
}

#endregion

#region ==================== UTILIDADES BASE ====================

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Force -Path $Path | Out-Null
    }
}

function Write-FileNoBOM {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content
    )
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Set-Notice {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Color = 'Yellow'
    )
    $script:UiState.Notice = $Message
    $script:UiState.NoticeColor = $Color
}

function Clear-Notice {
    $script:UiState.Notice = ''
    $script:UiState.NoticeColor = 'DarkGray'
}

function Wait-AnyKey {
    Write-Host ""
    Write-Host "Pulsa cualquier tecla para volver..." -ForegroundColor DarkGray
    [void][Console]::ReadKey($true)
}

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
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Format-Uptime {
    param($Uptime)

    if ($null -eq $Uptime) { return '-' }
    if ($Uptime -isnot [System.TimeSpan]) { return '-' }

    return "{0:00}d {1:00}h {2:00}m {3:00}s" -f `
        $Uptime.Days, $Uptime.Hours, $Uptime.Minutes, $Uptime.Seconds
}

function Get-ServerStatus {
    $proc = Get-ServerProcess
    $activeConfig = Get-ActiveScenario
    $latestLog = Get-LatestLogDir

    $status = [ordered]@{
        Running          = $false
        Pid              = $null
        PortOpen         = $false
        ActiveConfig     = $activeConfig
        Uptime           = $null
        LatestLogName    = if ($latestLog) { $latestLog.Name } else { $null }
        ActiveScenarioName = if ($activeConfig) { [IO.Path]::GetFileName($activeConfig) } else { $null }
    }

    if ($proc) {
        $status.Running = $true
        $status.Pid = $proc.Id

        try {
            $udp = Get-NetUDPEndpoint -OwningProcess $proc.Id -ErrorAction SilentlyContinue |
                   Where-Object { $_.LocalPort -eq $Config.GamePort }
            if ($udp) { $status.PortOpen = $true }
        }
        catch {
            # No romper la UI si falla este cmdlet
        }

        try {
            $status.Uptime = (Get-Date) - $proc.StartTime
        }
        catch {
            $status.Uptime = $null
        }
    }

    return [PSCustomObject]$status
}

function Get-AllScenarios {
    $items = @()
    $index = 1

    if (-not (Test-Path $Paths.ScenariosDir)) {
        return $items
    }

    $files = Get-ChildItem -Path $Paths.ScenariosDir -Filter '*.json' -File -Recurse |
             Sort-Object FullName

    foreach ($file in $files) {
        $folder = $file.DirectoryName.Substring($Paths.ScenariosDir.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($folder)) { $folder = '(raíz)' }

        $items += [PSCustomObject]@{
            Index    = $index
            Name     = $file.BaseName
            FullPath = $file.FullName
            Folder   = $folder
        }

        $index++
    }

    return $items
}

function Assert-Prerequisites {
    $missing = @()

    if (-not (Test-Path $Paths.ServerExe)) { $missing += $Paths.ServerExe }
    if (-not (Test-Path $Paths.BerconCli)) { $missing += $Paths.BerconCli }
    if (-not (Test-Path $Paths.BEConfig))  { $missing += $Paths.BEConfig }

    if ($missing.Count -gt 0) {
        throw "Faltan rutas obligatorias:`n - " + ($missing -join "`n - ")
    }

    Ensure-Directory -Path $Paths.AdminDir
}

#endregion

#region ==================== BERCON-CLI ====================

function Invoke-Bercon {
    param(
        [Parameter(Mandatory)][string[]]$Commands,
        [ValidateSet('table', 'json', 'raw', 'md', 'html')]
        [string]$Format = 'table',
        [int]$TimeoutSec = 5
    )

    $args = @(
        '-r', $Paths.BEConfig,
        '-t', $TimeoutSec.ToString(),
        '-f', $Format,
        '--'
    ) + $Commands

    $oldPref = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = & $Paths.BerconCli @args 2>&1
        $exitCode = $LASTEXITCODE
        $text = (($output | ForEach-Object { "$_" }) -join "`n").Trim()

        return [PSCustomObject]@{
            Success  = ($exitCode -eq 0)
            ExitCode = $exitCode
            Output   = $text
        }
    }
    catch {
        return [PSCustomObject]@{
            Success  = $false
            ExitCode = -1
            Output   = $_.Exception.Message
        }
    }
    finally {
        $ErrorActionPreference = $oldPref
    }
}

function Get-PlayersRaw {
    $server = Get-ServerProcess
    if (-not $server) {
        return [PSCustomObject]@{
            State  = 'stopped'
            Result = $null
        }
    }

    $result = Invoke-Bercon -Commands @('players') -Format 'json' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        return [PSCustomObject]@{
            State  = 'ok'
            Result = $result
        }
    }

    return [PSCustomObject]@{
        State  = 'unavailable'
        Result = $result
    }
}

function Get-PlayerSummary {
    $raw = Get-PlayersRaw

    switch ($raw.State) {
        'stopped' {
            return [PSCustomObject]@{
                State      = 'stopped'
                Count      = '-'
                AvgPing    = '-'
                InLobby    = '-'
                ValidCount = '-'
            }
        }

        'unavailable' {
            return [PSCustomObject]@{
                State      = 'warming'
                Count      = '-'
                AvgPing    = '-'
                InLobby    = '-'
                ValidCount = '-'
            }
        }

        'ok' {
            try {
                $players = $raw.Result.Output | ConvertFrom-Json -ErrorAction Stop
            }
            catch {
                return [PSCustomObject]@{
                    State      = 'warming'
                    Count      = '-'
                    AvgPing    = '-'
                    InLobby    = '-'
                    ValidCount = '-'
                }
            }

            if ($null -eq $players) { $players = @() }
            if ($players -isnot [System.Array]) { $players = @($players) }

            $count = @($players).Count
            $avgPing = '-'
            $inLobby = 0
            $validCount = 0

            if ($count -gt 0) {
                $pings = @($players | Where-Object { $_.ping -ne $null } | ForEach-Object { [double]$_.ping })
                if ($pings.Count -gt 0) {
                    $avgPing = [math]::Round((($pings | Measure-Object -Average).Average), 0)
                }

                $inLobby = @($players | Where-Object { $_.lobby -eq $true }).Count
                $validCount = @($players | Where-Object { $_.valid -eq $true }).Count
            }

            return [PSCustomObject]@{
                State      = 'ok'
                Count      = $count
                AvgPing    = $avgPing
                InLobby    = $inLobby
                ValidCount = $validCount
            }
        }
    }
}

function Show-PlayersScreen {
    $raw = Get-PlayersRaw

    Clear-Host
    Write-Host ""
    Write-Host "  ╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                    JUGADORES CONECTADOS                      ║" -ForegroundColor Cyan
    Write-Host "  ╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    switch ($raw.State) {
        'stopped' {
            Write-Host "`n  El servidor está detenido." -ForegroundColor DarkGray
            Wait-AnyKey
            return
        }

        'unavailable' {
            Write-Host "`n  RCon todavía no responde o sigue arrancando." -ForegroundColor Yellow
            Write-Host "  No es necesariamente un error; puede tardar unos segundos." -ForegroundColor DarkGray
            if ($raw.Result -and $raw.Result.Output) {
                Write-Host ""
                Write-Host "  Detalle:" -ForegroundColor Yellow
                Write-Host "  $($raw.Result.Output)" -ForegroundColor DarkGray
            }
            Wait-AnyKey
            return
        }

        'ok' {
            try {
                $players = $raw.Result.Output | ConvertFrom-Json -ErrorAction Stop
            }
            catch {
                Write-Host "`n  No se pudo interpretar la respuesta JSON de jugadores." -ForegroundColor Red
                Wait-AnyKey
                return
            }

            if ($null -eq $players) { $players = @() }
            if ($players -isnot [System.Array]) { $players = @($players) }

            if (@($players).Count -eq 0) {
                Write-Host "`n  0 jugadores conectados." -ForegroundColor Green
                Wait-AnyKey
                return
            }

            $display = $players | Select-Object `
                @{Name='ID';Expression={$_.id}},
                @{Name='Nombre';Expression={$_.name}},
                @{Name='Ping';Expression={$_.ping}},
                @{Name='Valido';Expression={$_.valid}},
                @{Name='Lobby';Expression={$_.lobby}},
                @{Name='IP';Expression={$_.ip}},
                @{Name='Puerto';Expression={$_.port}}

            Write-Host ""
            Write-Host ($display | Format-Table -AutoSize | Out-String)
            Wait-AnyKey
            return
        }
    }
}

function Send-ServerMessage {
    $msg = Read-Host "`nEscribe el mensaje global"
    if ([string]::IsNullOrWhiteSpace($msg)) {
        Set-Notice "Mensaje cancelado." 'DarkGray'
        return
    }

    $result = Invoke-Bercon -Commands @("say -1 $msg") -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        Set-Notice "Mensaje global enviado." 'Green'
    }
    else {
        Set-Notice "No se pudo enviar el mensaje: $($result.Output)" 'Red'
    }
}

function Lock-Server {
    $result = Invoke-Bercon -Commands @('#lock') -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        Set-Notice "Servidor bloqueado para nuevas conexiones." 'Green'
    }
    else {
        Set-Notice "No se pudo bloquear el servidor: $($result.Output)" 'Red'
    }
}

function Unlock-Server {
    $result = Invoke-Bercon -Commands @('#unlock') -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        Set-Notice "Servidor desbloqueado." 'Green'
    }
    else {
        Set-Notice "No se pudo desbloquear el servidor: $($result.Output)" 'Red'
    }
}

function Kick-PlayerInteractive {
    $raw = Get-PlayersRaw

    if ($raw.State -eq 'stopped') {
        Set-Notice "No hay servidor en ejecución." 'DarkGray'
        return
    }

    if ($raw.State -eq 'unavailable') {
        Set-Notice "RCon todavía no responde. Prueba de nuevo en unos segundos." 'Yellow'
        return
    }

    if ($raw.Result.Output -match '\(0\s+in total\)') {
        Set-Notice "No hay jugadores conectados." 'DarkGray'
        return
    }

    Clear-Host
    Write-Host ""
    Write-Host "  ╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                    EXPULSAR JUGADOR                          ║" -ForegroundColor Cyan
    Write-Host "  ╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host $raw.Result.Output

    $id = Read-Host "`nID del jugador a expulsar"
    if ([string]::IsNullOrWhiteSpace($id)) {
        Set-Notice "Kick cancelado." 'DarkGray'
        return
    }

    if ($id -notmatch '^\d+$') {
        Set-Notice "El ID debe ser numérico." 'Red'
        return
    }

    $confirm = Read-Host "Escribe SI para confirmar el kick del playerId $id"
    if ($confirm -ne 'SI') {
        Set-Notice "Kick cancelado." 'DarkGray'
        return
    }

    $result = Invoke-Bercon -Commands @("kick $id") -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        Set-Notice "Kick enviado para playerId $id." 'Green'
    }
    else {
        Set-Notice "Falló el kick: $($result.Output)" 'Red'
    }
}

function Send-CustomRConCommand {
    $cmd = Read-Host "`nComando RCon personalizado"
    if ([string]::IsNullOrWhiteSpace($cmd)) {
        Set-Notice "Comando cancelado." 'DarkGray'
        return
    }

    $result = Invoke-Bercon -Commands @($cmd) -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if ($result.Success) {
        if ($result.Output) {
            Set-Notice "Comando enviado. Salida: $($result.Output)" 'Green'
        }
        else {
            Set-Notice "Comando enviado." 'Green'
        }
    }
    else {
        Set-Notice "Falló el comando: $($result.Output)" 'Red'
    }
}

function Stop-ServerGraceful {
    $serverProc = Get-ServerProcess
    if (-not $serverProc) {
        Set-Notice "No hay servidor en ejecución." 'DarkGray'
        return
    }

    Set-Notice "Enviando #shutdown..." 'Cyan'
    $result = Invoke-Bercon -Commands @('#shutdown') -Format 'raw' -TimeoutSec $Config.RConTimeoutSec
    if (-not $result.Success) {
        Set-Notice "No se pudo enviar #shutdown: $($result.Output)" 'Red'
        return
    }

    $deadline = (Get-Date).AddSeconds($Config.ShutdownTimeoutSec)
    do {
        Start-Sleep -Milliseconds 700
        $serverProc = Get-ServerProcess
    } while ($serverProc -and ((Get-Date) -lt $deadline))

    if (-not $serverProc) {
        Set-Notice "Servidor detenido correctamente." 'Green'
    }
    else {
        Set-Notice "El comando se envió, pero el proceso sigue vivo." 'Yellow'
    }
}

#endregion

#region ==================== ARRANQUE ====================

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
    $activeConfig = Get-ActiveScenario
    if (-not $activeConfig) {
        Set-Notice "No hay escenario seleccionado." 'Yellow'
        return
    }

    if (-not (Test-Path $activeConfig)) {
        Set-Notice "El JSON seleccionado no existe: $activeConfig" 'Red'
        return
    }

    $serverProc = Get-ServerProcess
    if ($serverProc) {
        Set-Notice "Ya hay una instancia del servidor en ejecución. Usa [S] para apagarla." 'Yellow'
        return
    }

    Ensure-Directory -Path $Paths.AdminDir
    Write-LauncherBat -ConfigPath $activeConfig

    Start-Process -FilePath $Paths.LauncherBat | Out-Null
    Set-Notice "Servidor lanzado: $([IO.Path]::GetFileName($activeConfig))" 'Green'
    Start-Sleep -Seconds 2
}

#endregion

#region ==================== INTERFAZ ====================

function Show-Header {
    $status = Get-ServerStatus
    $stats = Get-PlayerSummary

    $procText  = if ($status.Running) { "● CORRIENDO (PID $($status.Pid))" } else { "○ DETENIDO" }
    $procColor = if ($status.Running) { "Green" } else { "Red" }

    $portText  = if ($status.PortOpen) { "● ABIERTO (UDP $($Config.GamePort))" } else { "○ CERRADO / SIN CONFIRMAR" }
    $portColor = if ($status.PortOpen) { "Green" } else { "DarkGray" }

    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host ("  ║  ARMA REFORGER - GESTOR BERCON {0,-35}║" -f $Config.VersionTag) -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Host "`n    RESUMEN:" -ForegroundColor Yellow
    Write-Host "      Proceso:                  " -NoNewline; Write-Host $procText -ForegroundColor $procColor
    Write-Host "      Puerto juego:             " -NoNewline; Write-Host $portText -ForegroundColor $portColor
    Write-Host "      Uptime:                   " -NoNewline; Write-Host (Format-Uptime $status.Uptime) -ForegroundColor Cyan
    Write-Host "      Jugadores:                " -NoNewline; Write-Host $stats.Count -ForegroundColor Green
    Write-Host "      Estado RCon:              " -NoNewline
    switch ($stats.State) {
        'ok'      { Write-Host 'OK' -ForegroundColor Green }
        'warming' { Write-Host 'ARRANCANDO / NO DISPONIBLE' -ForegroundColor Yellow }
        'stopped' { Write-Host '-' -ForegroundColor DarkGray }
    }

    Write-Host "      Profile:                  " -NoNewline; Write-Host $Config.ProfilePath -ForegroundColor DarkCyan

    if ($status.Running -and $status.ActiveScenarioName) {
        Write-Host "      Misión activa:            " -NoNewline
        Write-Host $status.ActiveScenarioName -ForegroundColor Yellow
    }
    elseif ($status.ActiveScenarioName) {
        Write-Host "      Última misión seleccionada:" -NoNewline
        Write-Host " $($status.ActiveScenarioName)" -ForegroundColor DarkYellow
    }

    if ($status.LatestLogName) {
        Write-Host "      Último log:               " -NoNewline
        Write-Host $status.LatestLogName -ForegroundColor DarkCyan
    }

    Write-Host ""
    Write-Host "    AVISO: " -NoNewline -ForegroundColor Yellow
    if ([string]::IsNullOrWhiteSpace($script:UiState.Notice)) {
        Write-Host "-" -ForegroundColor DarkGray
    }
    else {
        Write-Host $script:UiState.Notice -ForegroundColor $script:UiState.NoticeColor
    }
}

function Show-ScenarioList {
    $status = Get-ServerStatus
    $list = Get-AllScenarios

    Write-Host "`n    ESCENARIOS DISPONIBLES:" -ForegroundColor Yellow

    if ($list.Count -eq 0) {
        Write-Host "      (No se han encontrado JSON en $($Paths.ScenariosDir))" -ForegroundColor DarkGray
        return
    }

    $currentFolder = ''
    foreach ($s in $list) {
        if ($s.Folder -ne $currentFolder) {
            $currentFolder = $s.Folder
            Write-Host "`n      --- $currentFolder ---" -ForegroundColor DarkCyan
        }

        $isSelected = ($status.ActiveConfig -and ($s.FullPath -eq $status.ActiveConfig))
        $marker = '  '
        $color  = 'White'

        if ($status.Running -and $isSelected) {
            $marker = '► '
            $color = 'Green'
        }
        elseif ((-not $status.Running) -and $isSelected) {
            $marker = '· '
            $color = 'DarkYellow'
        }

        Write-Host "      $marker[$($s.Index.ToString('00'))] " -NoNewline -ForegroundColor $color
        Write-Host ($s.Name.PadRight(40)) -ForegroundColor $color
    }
}

function Show-Actions {
    Write-Host "`n    ACCIONES RÁPIDAS:" -ForegroundColor Yellow
    Write-Host "      [1-99] Lanzar misión"
    Write-Host "      [S]    Apagar servidor"
    Write-Host "      [P]    Ver jugadores"
    Write-Host "      [M]    Enviar mensaje global"
    Write-Host "      [L]    Bloquear entradas"
    Write-Host "      [U]    Desbloquear entradas"
    Write-Host "      [K]    Expulsar jugador por ID"
    Write-Host "      [C]    Comando RCon personalizado"
    Write-Host "      [R]    Limpiar aviso / refrescar"
    Write-Host "      [Q]    Salir"
}

#endregion

#region ==================== MAIN ====================

Assert-Prerequisites

while ($true) {
    Show-Header
    Show-ScenarioList
    Show-Actions

    $sel = Read-Host "`n    Selección"

    if ([string]::IsNullOrWhiteSpace($sel)) {
        continue
    }

    switch -Regex ($sel.ToUpper()) {
        '^[0-9]+$' {
            $list = Get-AllScenarios
            $item = $list | Where-Object { $_.Index -eq [int]$sel } | Select-Object -First 1
            if ($item) {
                Set-ActiveScenario -ConfigPath $item.FullPath
                Start-ServerDirect
            }
            else {
                Set-Notice "No existe ese índice de escenario." 'Red'
            }
        }

        '^S$' { Stop-ServerGraceful }
        '^P$' { Show-PlayersScreen }
        '^M$' { Send-ServerMessage }
        '^L$' { Lock-Server }
        '^U$' { Unlock-Server }
        '^K$' { Kick-PlayerInteractive }
        '^C$' { Send-CustomRConCommand }
        '^R$' { Clear-Notice }
        '^Q$' { break }
        default { Set-Notice "Opción no válida." 'Red' }
    }
}

#endregion
