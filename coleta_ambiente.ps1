# ============================================================
#  CHECK DE AMBIENTE - MobyCRM / MobyPharma / FCerta
#  Valida servidor e estação conforme requisitos técnicos
# ============================================================

$ErrorActionPreference = "SilentlyContinue"

# ============================================================
#  SELEÇÃO: SERVIDOR OU ESTAÇÃO
# ============================================================
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║   CHECK DE AMBIENTE - MobyPharma/CRM    ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Tipo de máquina:" -ForegroundColor White
Write-Host "  [1] Servidor  (até 10 micros)" -ForegroundColor Yellow
Write-Host "  [2] Estação   (workstation)"   -ForegroundColor Yellow
Write-Host ""

do {
    $choice = Read-Host "  Digite 1 ou 2"
} while ($choice -notin @("1","2"))

$isServer = ($choice -eq "1")
$tipoMaq  = if ($isServer) { "SERVIDOR" } else { "ESTAÇÃO" }

# ============================================================
#  REQUISITOS MÍNIMOS (conforme documentação MobyPharma)
# ============================================================
if ($isServer) {
    $REQ_CPU_CORES  = 4
    $REQ_CPU_GHZ    = 3.0
    $REQ_RAM_GB     = 8
    $REQ_DISK_GB    = 50
    $REQ_NET_MBITS  = 100
    $REQ_OS_OK      = @("Windows 10","Windows 8.1","Windows 2012","Windows 2008")
} else {
    $REQ_CPU_CORES  = 2      # sem exigência de núcleos mínimos, só GHz
    $REQ_CPU_GHZ    = 2.0
    $REQ_RAM_GB     = 4
    $REQ_DISK_GB    = 4
    $REQ_NET_MBITS  = 10
    $REQ_OS_OK      = @("Windows 10","Windows 8.1")
}

$REQ_RES_W       = 1280
$REQ_RES_H       = 600
$REQ_DOTNET      = @("2.0","3.5","4.5","4.6")   # 4.6 cobre 4.6.x
$REQ_CHROME_MIN  = 83

# ============================================================
#  SAÍDA
# ============================================================
$OutDir = "C:\Temp\Moby_Check"
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$stamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportPath = Join-Path $OutDir "Check_${tipoMaq}_$stamp.txt"

# ============================================================
#  FUNÇÕES AUXILIARES
# ============================================================
function Head($t) {
    Write-Host ""
    Write-Host ("=" * 62) -ForegroundColor Cyan
    Write-Host "  $t" -ForegroundColor Cyan
    Write-Host ("=" * 62) -ForegroundColor Cyan
}
function OK($m)   { Write-Host "  [ OK ] $m" -ForegroundColor Green }
function WARN($m) { Write-Host "  [AVS ] $m" -ForegroundColor Yellow }
function FAIL($m) { Write-Host "  [ NOK] $m" -ForegroundColor Red }
function INF($m)  { Write-Host "         $m" -ForegroundColor Gray }

function Format-Size([long]$bytes) {
    if ($bytes -ge 1GB) { return "$([math]::Round($bytes/1GB,2)) GB" }
    if ($bytes -ge 1MB) { return "$([math]::Round($bytes/1MB,0)) MB" }
    return "$bytes bytes"
}

function Get-AdapterSpeedMbps($adapter) {
    foreach ($prop in @("ReceiveLinkSpeed","TransmitLinkSpeed")) {
        $bps = $adapter.$prop
        if ($bps -and $bps -gt 0) { return [math]::Round($bps / 1000000, 0) }
    }
    $ls = "$($adapter.LinkSpeed)".Trim()
    if ($ls -match '(\d+\.?\d*)\s*(G|M|K)?bps') {
        $num  = [double]$Matches[1]
        $mbps = switch ($Matches[2]) {
            'G' { $num * 1000 }
            'M' { $num }
            'K' { $num / 1000 }
            default { $num / 1000000 }
        }
        if ($mbps -gt 0) { return [math]::Round($mbps, 0) }
    }
    $wmi = Get-CimInstance Win32_NetworkAdapter -EA SilentlyContinue |
           Where-Object { $_.Name -eq $adapter.InterfaceDescription -and $_.Speed -gt 0 } |
           Select-Object -First 1
    if ($wmi) { return [math]::Round($wmi.Speed / 1000000, 0) }
    return 0
}

function Find-FcertaRoot {
    $namePattern = '^(fcerta|formula.?certa|formulacerta)$'
    $drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -EA SilentlyContinue |
              Select-Object -ExpandProperty DeviceID
    foreach ($d in $drives) {
        $hit = Get-ChildItem "$d\" -Directory -EA SilentlyContinue |
               Where-Object { $_.Name -match $namePattern } |
               Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    foreach ($d in $drives) {
        foreach ($name in @("Fcerta","FCerta","fcerta","FormulaCerta")) {
            $p = "$d\$name"
            if (Test-Path $p) { return $p }
        }
    }
    return $null
}

function Get-FirebirdVersion {
    $svcNames = @("FirebirdServerDefaultInstance","FirebirdServer",
                  "FirebirdGuardianDefaultInstance","FirebirdGuardian")
    foreach ($n in $svcNames) {
        $svc = Get-CimInstance Win32_Service -Filter "Name='$n'" -EA SilentlyContinue
        if ($svc -and $svc.PathName) {
            $exe = $svc.PathName.Trim()
            if ($exe.StartsWith('"')) { $exe = $exe.Split('"')[1] }
            else                      { $exe = $exe.Split(' ')[0] }
            if (Test-Path $exe) {
                $vi  = (Get-Item $exe).VersionInfo
                $ver = if ($vi.ProductVersion) { $vi.ProductVersion } else { $vi.FileVersion }
                if ($ver -match '^\d+\.\d+') { return $ver }
            }
        }
    }
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($rp in $regPaths) {
        $app = Get-ItemProperty $rp -EA SilentlyContinue |
               Where-Object { $_.DisplayName -match "Firebird" -and $_.DisplayVersion } |
               Sort-Object DisplayVersion -Descending | Select-Object -First 1
        if ($app) { return $app.DisplayVersion }
    }
    foreach ($p in @(
        "$env:ProgramFiles\Firebird\Firebird_*\bin\fbserver.exe",
        "$env:ProgramFiles\Firebird\Firebird*\bin\fbserver.exe",
        "C:\Firebird\bin\fbserver.exe"
    )) {
        $f = Get-Item $p -EA SilentlyContinue | Select-Object -First 1
        if ($f) {
            $ver = $f.VersionInfo.ProductVersion
            if ($ver -match '^\d+\.\d+') { return $ver }
        }
    }
    return ""
}

function Get-DotNetVersions {
    $versions = @()
    # .NET 1.x - 4.x via registro
    $ndpPath = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"
    if (Test-Path $ndpPath) {
        Get-ChildItem $ndpPath -EA SilentlyContinue | ForEach-Object {
            $name = $_.PSChildName
            if ($name -match '^v(\d)') {
                # Sub-chaves (Client, Full, etc)
                $subKeys = Get-ChildItem $_.PSPath -EA SilentlyContinue
                if ($subKeys) {
                    foreach ($sub in $subKeys) {
                        $inst    = (Get-ItemProperty $sub.PSPath -EA SilentlyContinue).Install
                        $release = (Get-ItemProperty $sub.PSPath -EA SilentlyContinue).Release
                        $ver     = (Get-ItemProperty $sub.PSPath -EA SilentlyContinue).Version
                        if ($inst -eq 1 -and $ver) { $versions += $ver }
                        # .NET 4.5+ usa Release key
                        if ($release) {
                            $mapped = switch ($true) {
                                ($release -ge 533320) { "4.8.1" }
                                ($release -ge 528040) { "4.8" }
                                ($release -ge 461808) { "4.7.2" }
                                ($release -ge 461308) { "4.7.1" }
                                ($release -ge 460798) { "4.7" }
                                ($release -ge 394802) { "4.6.2" }
                                ($release -ge 394254) { "4.6.1" }
                                ($release -ge 393295) { "4.6" }
                                ($release -ge 379893) { "4.5.2" }
                                ($release -ge 378675) { "4.5.1" }
                                ($release -ge 378389) { "4.5" }
                                default               { $null }
                            }
                            if ($mapped) { $versions += $mapped }
                        }
                    }
                } else {
                    $inst = (Get-ItemProperty $_.PSPath -EA SilentlyContinue).Install
                    $ver  = (Get-ItemProperty $_.PSPath -EA SilentlyContinue).Version
                    if ($inst -eq 1 -and $ver) { $versions += $ver }
                }
            }
        }
    }
    return ($versions | Sort-Object -Unique)
}

function Get-ChromeVersion {
    $paths = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            return (Get-Item $p).VersionInfo.ProductVersion
        }
    }
    # Registro
    $regPaths = @(
        "HKLM:\SOFTWARE\Google\Chrome",
        "HKLM:\SOFTWARE\WOW6432Node\Google\Chrome",
        "HKCU:\SOFTWARE\Google\Chrome"
    )
    foreach ($rp in $regPaths) {
        $ver = (Get-ItemProperty $rp -EA SilentlyContinue).Version
        if ($ver) { return $ver }
    }
    return ""
}

function Get-LastBackupDate([string]$backupPath) {
    if (-not (Test-Path $backupPath)) { return "" }
    $last = Get-ChildItem $backupPath -Directory -EA SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($last) { return $last.LastWriteTime.ToString("dd/MM/yyyy") }
    return ""
}

# ============================================================
#  COLETA DE DADOS
# ============================================================
Write-Host ""
Write-Host "  Coletando informações — $tipoMaq..." -ForegroundColor Cyan

$os          = Get-CimInstance Win32_OperatingSystem
$osName      = $os.Caption.Trim()
$osBuild     = $os.BuildNumber
$osArch      = $os.OSArchitecture

$cpu         = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpuName     = $cpu.Name.Trim()
$cpuCores    = $cpu.NumberOfCores
$cpuLogic    = $cpu.NumberOfLogicalProcessors
$cpuGHz      = [math]::Round($cpu.MaxClockSpeed / 1000, 2)

$ramTotalGB  = [math]::Round(($os.TotalVisibleMemorySize * 1KB) / 1GB, 1)
$ramFreeGB   = [math]::Round(($os.FreePhysicalMemory  * 1KB) / 1GB, 1)

$gpus        = Get-CimInstance Win32_VideoController
$gpuList     = ($gpus | ForEach-Object { $_.Name.Trim() }) -join " | "

$adapters    = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }

$fbVer       = Get-FirebirdVersion
$dotNets     = Get-DotNetVersions
$chromeVer   = Get-ChromeVersion

$fcertaRoot  = Find-FcertaRoot
$dbSize      = ""; $imSize = ""; $lastBackup = ""; $dbPath = ""; $imPath = ""

if ($fcertaRoot) {
    $dbFile = Join-Path $fcertaRoot "db\alterdb.ib"
    if (Test-Path $dbFile) { $dbPath = $dbFile; $dbSize = Format-Size (Get-Item $dbFile).Length }

    $imFile = Join-Path $fcertaRoot "db\alterim.ib"
    if (Test-Path $imFile) { $imPath = $imFile; $imSize = Format-Size (Get-Item $imFile).Length }

    $lastBackup = Get-LastBackupDate (Join-Path $fcertaRoot "FormulaCertaUpdate\Backup")
}

# Disco do sistema (C:)
$sysDrive   = Get-PSDrive -Name "C" -EA SilentlyContinue
$diskFreeGB = if ($sysDrive) { [math]::Round($sysDrive.Free / 1GB, 1) } else { 0 }

# Resolução
$resOk = $false
foreach ($g in $gpus) {
    if ($g.CurrentHorizontalResolution -ge $REQ_RES_W -and
        $g.CurrentVerticalResolution   -ge $REQ_RES_H) { $resOk = $true; break }
}

# Chrome versão
$chromeOk = $false
if ($chromeVer -match '^(\d+)') {
    $chromeMajor = [int]$Matches[1]
    $chromeOk = ($chromeMajor -ge $REQ_CHROME_MIN)
}

# .NET - verifica se cada versão requerida está coberta
$dotNetStatus = @{}
foreach ($req in $REQ_DOTNET) {
    $dotNetStatus[$req] = ($dotNets | Where-Object { $_ -like "$req*" }).Count -gt 0
}
$dotNetOk = -not ($dotNetStatus.Values -contains $false)

# ============================================================
#  EXIBIÇÃO - DIAGNÓSTICO
# ============================================================

Head "1. SISTEMA OPERACIONAL  [$tipoMaq]"
INF "Sistema : $osName"
INF "Build   : $osBuild  |  Arquit.: $osArch"
$osOk = ($REQ_OS_OK | Where-Object { $osName -like "*$_*" }).Count -gt 0
if ($osOk) { OK "$osName" }
else        { FAIL "SO não suportado: $osName" ; INF "Suportados: $($REQ_OS_OK -join ' | ')" }

Head "2. PROCESSADOR"
INF "Modelo  : $cpuName"
INF "Cores   : $cpuCores físicos / $cpuLogic lógicos  |  Clock: $cpuGHz GHz"
if ($isServer) {
    if ($cpuCores -ge $REQ_CPU_CORES) { OK "$cpuCores núcleos (mínimo $REQ_CPU_CORES)" }
    else                               { FAIL "Apenas $cpuCores núcleo(s) — mínimo $REQ_CPU_CORES para servidor" }
}
if ($cpuGHz -ge $REQ_CPU_GHZ) { OK "$cpuGHz GHz (mínimo $REQ_CPU_GHZ GHz)" }
else                            { FAIL "$cpuGHz GHz insuficiente — mínimo $REQ_CPU_GHZ GHz" }

Head "3. MEMÓRIA RAM"
INF "Total   : $ramTotalGB GB  |  Livre: $ramFreeGB GB"
if ($ramTotalGB -ge $REQ_RAM_GB) { OK "$ramTotalGB GB RAM (mínimo $REQ_RAM_GB GB)" }
else                              { FAIL "RAM insuficiente: $ramTotalGB GB (mínimo $REQ_RAM_GB GB)" }

Head "4. DISCO RÍGIDO"
INF "Livre em C: : $diskFreeGB GB"
if ($diskFreeGB -ge $REQ_DISK_GB) { OK "$diskFreeGB GB livres (mínimo $REQ_DISK_GB GB)" }
elseif ($diskFreeGB -ge ($REQ_DISK_GB * 0.5)) { WARN "Espaço livre baixo: $diskFreeGB GB (mínimo $REQ_DISK_GB GB)" }
else  { FAIL "Espaço insuficiente: $diskFreeGB GB (mínimo $REQ_DISK_GB GB)" }

Head "5. RESOLUÇÃO DO MONITOR"
foreach ($g in $gpus) {
    $w = $g.CurrentHorizontalResolution; $h = $g.CurrentVerticalResolution
    INF "Monitor : ${w}x${h}  —  $($g.Name.Trim())"
    if ($w -ge $REQ_RES_W -and $h -ge $REQ_RES_H) { OK "${w}x${h} (mínimo ${REQ_RES_W}x${REQ_RES_H})" }
    else { FAIL "${w}x${h} abaixo do mínimo ${REQ_RES_W}x${REQ_RES_H}" }
}

Head "6. .NET FRAMEWORK"
INF "Instalados: $($dotNets -join ', ')"
foreach ($req in $REQ_DOTNET | Sort-Object) {
    if ($dotNetStatus[$req]) { OK ".NET $req encontrado" }
    else                     { FAIL ".NET $req NÃO encontrado" }
}

Head "7. NAVEGADOR"
if ($chromeVer) {
    INF "Chrome  : $chromeVer"
    if ($chromeOk) { OK "Chrome $chromeVer (mínimo v$REQ_CHROME_MIN)" }
    else           { FAIL "Chrome $chromeVer desatualizado (mínimo v$REQ_CHROME_MIN)" }
} else {
    WARN "Google Chrome não encontrado — instale v$REQ_CHROME_MIN ou superior"
}

$ie = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Internet Explorer" -EA SilentlyContinue
if ($ie -and $ie.svcVersion -match '^11') { OK "Internet Explorer 11 detectado" }

Head "8. FIREBIRD"
if ($fbVer) {
    INF "Versão  : $fbVer"
    OK "Firebird $fbVer detectado"
} else {
    FAIL "Firebird não encontrado"
}

Head "9. BANCO DE DADOS FCerta"
if ($fcertaRoot) {
    INF "Raiz    : $fcertaRoot"
    if ($dbPath) { INF "DB      : $dbPath" ; OK "Banco de Dados: $dbSize" }
    else         { WARN "alterdb.ib não encontrado" }
    if ($imPath) { INF "Imagem  : $imPath" ; OK "Banco de Imagem: $imSize" }
    else         { WARN "alterim.ib não encontrado" }
    if ($lastBackup) { INF "Backup  : $lastBackup" ; OK "Último backup: $lastBackup" }
    else             { WARN "Pasta de backup não localizada" }

    $fDrive    = Split-Path -Qualifier $fcertaRoot
    $fDriveInf = Get-PSDrive -Name ($fDrive -replace ":","") -EA SilentlyContinue
    if ($fDriveInf) {
        $fFree  = [math]::Round($fDriveInf.Free / 1GB, 1)
        $fTotal = [math]::Round(($fDriveInf.Used + $fDriveInf.Free) / 1GB, 1)
        INF "Disco $fDrive : $fFree GB livres de $fTotal GB"
        if    ($fFree -lt 10) { FAIL "Espaço crítico: $fFree GB livres" }
        elseif($fFree -lt 20) { WARN "Espaço baixo: $fFree GB livres" }
        else                  { OK   "Espaço livre FCerta: $fFree GB" }
    }
} else {
    WARN "Pasta FCerta não encontrada em nenhum disco"
}

Head "10. REDE"
foreach ($a in $adapters) {
    $mbps   = Get-AdapterSpeedMbps $a
    $isWifi = $a.PhysicalMediaType -like "*802.11*" -or $a.Name -match "Wi-?Fi|Wireless"
    $tipo   = if ($isWifi) { "Wireless" } else { "Cabeada" }
    $minReq = if ($isWifi) { 10 } else { $REQ_NET_MBITS }
    INF "$tipo : $($a.Name) — $mbps Mbps"
    if ($mbps -ge $minReq) { OK  "$tipo: $mbps Mbps (mínimo $minReq Mbps)" }
    else                   { FAIL "$tipo: $mbps Mbps insuficiente (mínimo $minReq Mbps)" }
}

# ============================================================
#  CHECKLIST FINAL
# ============================================================
$checks = [ordered]@{
    "SO compatível"          = $osOk
    "Processador GHz"        = ($cpuGHz -ge $REQ_CPU_GHZ)
    "RAM suficiente"         = ($ramTotalGB -ge $REQ_RAM_GB)
    "Disco livre suficiente" = ($diskFreeGB -ge $REQ_DISK_GB)
    "Resolução OK"           = $resOk
    ".NET Framework"         = $dotNetOk
    "Navegador (Chrome/IE)"  = ($chromeOk -or ($ie -and $ie.svcVersion -match '^11'))
    "Firebird instalado"     = ($fbVer -ne "")
    "FCerta localizado"      = ($null -ne $fcertaRoot)
    "Banco de dados (.ib)"   = ($dbPath -ne "")
    "Rede suficiente"        = (($adapters | Where-Object {
                                    $m = Get-AdapterSpeedMbps $_
                                    $w = $_.PhysicalMediaType -like "*802.11*" -or $_.Name -match "Wi-?Fi|Wireless"
                                    $m -ge (if ($w) { 10 } else { $REQ_NET_MBITS })
                                }).Count -gt 0)
}

if ($isServer) {
    $checks["CPU Núcleos (servidor)"] = ($cpuCores -ge $REQ_CPU_CORES)
}

$apto = -not ($checks.Values -contains $false)

Write-Host ""
Write-Host ("=" * 62) -ForegroundColor $(if ($apto) { "Green" } else { "Red" })
Write-Host ""
foreach ($k in $checks.Keys) {
    $v   = $checks[$k]
    $cor = if ($v) { "Green" } else { "Red" }
    $ico = if ($v) { "[ OK ]" } else { "[ NOK]" }
    Write-Host "  $ico  $k" -ForegroundColor $cor
}
Write-Host ""
if ($apto) {
    Write-Host "  RESULTADO: AMBIENTE APTO para MobyPharma/CRM [$tipoMaq]" -ForegroundColor Green
} else {
    Write-Host "  RESULTADO: AMBIENTE NAO APTO — corrija os itens [NOK]"    -ForegroundColor Red
}
Write-Host ""
Write-Host ("=" * 62) -ForegroundColor $(if ($apto) { "Green" } else { "Red" })

# ============================================================
#  BRIEFING - layout para ticket
# ============================================================
$briefing = @"

════════════════════════════════════════════════════════════
 BRIEFING MobyCRM — $env:COMPUTERNAME  [$tipoMaq]
 Gerado em: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
════════════════════════════════════════════════════════════

Já preencheu o briefing MobyCRM:

Responsável:

Cargo:

PAF:

Última data do Sobre: $lastBackup

É TS:

Link de Pagamento:

IA:

BOT:

Informações Técnicas:

Sistema Operacional: $osName (Build $osBuild)
Processador: $cpuName ($cpuCores cores / $cpuLogic threads @ $cpuGHz GHz)
Memória RAM: $ramTotalGB GB
GPU: $gpuList
Firebird: $fbVer
.NET Framework: $($dotNets -join ', ')
Chrome: $chromeVer
Banco de Dados: $dbSize
Banco de Imagem: $imSize
Pasta FCerta: $fcertaRoot
Disco livre (C:): $diskFreeGB GB

RESULTADO: $(if ($apto) { "APTO" } else { "NAO APTO" }) [$tipoMaq]

════════════════════════════════════════════════════════════
"@

Write-Host $briefing -ForegroundColor White

# ============================================================
#  SALVA TXT
# ============================================================
$checklistTxt = ($checks.GetEnumerator() | ForEach-Object {
    $ico = if ($_.Value) { "[OK]" } else { "[NOK]" }
    "  $ico  $($_.Key)"
}) -join "`n"

@"
════════════════════════════════════════════════════════════
 CHECK DE AMBIENTE - MobyPharma/CRM  [$tipoMaq]
 Computador : $env:COMPUTERNAME
 Usuário    : $env:USERNAME
 Gerado em  : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
════════════════════════════════════════════════════════════

HARDWARE
  SO          : $osName (Build $osBuild) [$osArch]
  Processador : $cpuName ($cpuCores cores / $cpuLogic threads @ $cpuGHz GHz)
  RAM         : $ramTotalGB GB total / $ramFreeGB GB livre
  Disco C:    : $diskFreeGB GB livres
  GPU         : $gpuList

AMBIENTE FCerta
  Pasta raiz  : $fcertaRoot
  Banco Dados : $dbSize  ($dbPath)
  Banco Imagem: $imSize  ($imPath)
  Firebird    : $fbVer
  Último Backup: $lastBackup

SOFTWARE
  .NET        : $($dotNets -join ', ')
  Chrome      : $chromeVer

REDE
$(($adapters | ForEach-Object {
    $m = Get-AdapterSpeedMbps $_
    "  $($_.Name): $m Mbps"
}) -join "`n")

CHECKLIST
$checklistTxt

RESULTADO: $(if ($apto) { "APTO PARA MobyPharma/CRM" } else { "NAO APTO — corrija os itens [NOK]" })
$briefing
"@ | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host "  Relatorio salvo em: $ReportPath" -ForegroundColor Cyan
Write-Host ""
