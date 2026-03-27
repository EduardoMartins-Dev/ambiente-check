# ============================================================
#  CHECK DE AMBIENTE - MobyCRM / FCerta / Phusion
#  Busca automática em todos os discos locais
# ============================================================

$ErrorActionPreference = "SilentlyContinue"

# ---------- Requisitos mínimos (KA-01264 / KA-01485) --------
$REQ_RAM_GB_MIN   = 8
$REQ_NET_MBITS    = 100
$REQ_OS_OK        = @("Windows 10", "Windows 11")

# ---------- Saída -----------------------------------------
$OutDir = "C:\Temp\Moby_Check"
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$stamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportPath = Join-Path $OutDir "Check_Ambiente_$stamp.txt"

# ============================================================
#  FUNÇÕES AUXILIARES
# ============================================================

function Sep  { Write-Host (("-" * 60)) -ForegroundColor DarkGray }
function Head($t) {
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  $t" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
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
    # Método 1: ReceiveLinkSpeed / TransmitLinkSpeed (numérico em bps)
    foreach ($prop in @("ReceiveLinkSpeed","TransmitLinkSpeed")) {
        $bps = $adapter.$prop
        if ($bps -and $bps -gt 0) { return [math]::Round($bps / 1000000, 0) }
    }

    # Método 2: parse da string LinkSpeed ("100 Mbps", "1 Gbps", "10 Gbps")
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

    # Método 3: WMI Win32_NetworkAdapter (fallback)
    $wmi = Get-CimInstance Win32_NetworkAdapter -EA SilentlyContinue |
           Where-Object { $_.Name -eq $adapter.InterfaceDescription -and $_.Speed -gt 0 } |
           Select-Object -First 1
    if ($wmi) { return [math]::Round($wmi.Speed / 1000000, 0) }

    return 0
}

# Todos os discos fixos locais
function Get-LocalDrives {
    Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -EA SilentlyContinue |
        Select-Object -ExpandProperty DeviceID
}

# ============================================================
#  AUTO-DISCOVERY - encontra a pasta raiz do FCerta
# ============================================================
function Find-FcertaRoot {
    $namePattern = '^(fcerta|formula.?certa|formulacerta)$'
    $drives = Get-LocalDrives

    # 1) Busca pastas no 1º nível de cada disco cujo nome bate com o padrão
    foreach ($d in $drives) {
        $hit = Get-ChildItem "$d\" -Directory -EA SilentlyContinue |
               Where-Object { $_.Name -match $namePattern } |
               Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }

    # 2) Fallback: paths conhecidos
    foreach ($d in $drives) {
        foreach ($name in @("Fcerta","FCerta","fcerta","FormulaCerta")) {
            $p = "$d\$name"
            if (Test-Path $p) { return $p }
        }
    }

    return $null
}

# ============================================================
#  FIREBIRD - versão
# ============================================================
function Get-FirebirdVersion {
    # Via serviço Windows
    $svcNames = @(
        "FirebirdServerDefaultInstance",
        "FirebirdServer",
        "FirebirdGuardianDefaultInstance",
        "FirebirdGuardian"
    )
    foreach ($n in $svcNames) {
        $svc = Get-CimInstance Win32_Service -Filter "Name='$n'" -EA SilentlyContinue
        if ($svc -and $svc.PathName) {
            $exe = $svc.PathName.Trim()
            if ($exe.StartsWith('"')) { $exe = $exe.Split('"')[1] }
            else                      { $exe = $exe.Split(' ')[0] }
            if (Test-Path $exe) {
                $vi = (Get-Item $exe).VersionInfo
                $ver = if ($vi.ProductVersion) { $vi.ProductVersion } else { $vi.FileVersion }
                if ($ver -match '^\d+\.\d+') { return $ver }
            }
        }
    }

    # Via registro (Uninstall)
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($rp in $regPaths) {
        $app = Get-ItemProperty $rp -EA SilentlyContinue |
               Where-Object { $_.DisplayName -match "Firebird" -and $_.DisplayVersion } |
               Sort-Object DisplayVersion -Descending |
               Select-Object -First 1
        if ($app) { return $app.DisplayVersion }
    }

    # Via executável em pastas comuns
    $exePaths = @(
        "$env:ProgramFiles\Firebird\Firebird_*\bin\fbserver.exe",
        "$env:ProgramFiles\Firebird\Firebird*\bin\fbserver.exe",
        "C:\Firebird\bin\fbserver.exe"
    )
    foreach ($p in $exePaths) {
        $f = Get-Item $p -EA SilentlyContinue | Select-Object -First 1
        if ($f) {
            $ver = $f.VersionInfo.ProductVersion
            if ($ver -match '^\d+\.\d+') { return $ver }
        }
    }

    return ""
}

# Data do backup mais recente (última subpasta modificada)
function Get-LastBackupDate([string]$backupPath) {
    if (-not (Test-Path $backupPath)) { return "" }
    $last = Get-ChildItem $backupPath -Directory -EA SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
    if ($last) { return $last.LastWriteTime.ToString("dd/MM/yyyy") }
    return ""
}

# ============================================================
#  COLETA DE DADOS
# ============================================================

Write-Host ""
Write-Host "  Coletando informações do ambiente..." -ForegroundColor Cyan
Write-Host ""

# --- SO ---
$os       = Get-CimInstance Win32_OperatingSystem
$osName   = $os.Caption.Trim()
$osBuild  = $os.BuildNumber
$osArch   = $os.OSArchitecture

# --- CPU ---
$cpu      = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpuName  = $cpu.Name.Trim()
$cpuCores = $cpu.NumberOfCores
$cpuLogic = $cpu.NumberOfLogicalProcessors
$cpuGHz   = [math]::Round($cpu.MaxClockSpeed / 1000, 2)

# --- RAM ---
$ramTotalBytes = $os.TotalVisibleMemorySize * 1KB
$ramFreeBytes  = $os.FreePhysicalMemory * 1KB
$ramTotalGB    = [math]::Round($ramTotalBytes / 1GB, 1)
$ramFreeGB     = [math]::Round($ramFreeBytes  / 1GB, 1)

# --- GPU ---
$gpus    = Get-CimInstance Win32_VideoController
$gpuList = ($gpus | ForEach-Object { $_.Name.Trim() }) -join " | "

# --- Firebird ---
$fbVer = Get-FirebirdVersion

# --- FCerta root ---
$fcertaRoot = Find-FcertaRoot
$dbSize     = ""
$imSize     = ""
$lastBackup = ""
$dbPath     = ""
$imPath     = ""

if ($fcertaRoot) {
    # Banco de dados principal: <root>\db\alterdb.ib
    $dbFile = Join-Path $fcertaRoot "db\alterdb.ib"
    if (Test-Path $dbFile) {
        $dbPath = $dbFile
        $dbSize = Format-Size (Get-Item $dbFile).Length
    }

    # Banco de imagens: <root>\db\alterim.ib
    $imFile = Join-Path $fcertaRoot "db\alterim.ib"
    if (Test-Path $imFile) {
        $imPath = $imFile
        $imSize = Format-Size (Get-Item $imFile).Length
    }

    # Último backup
    $backupPath = Join-Path $fcertaRoot "FormulaCertaUpdate\Backup"
    $lastBackup = Get-LastBackupDate $backupPath
}

# --- Rede ---
$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }

# ============================================================
#  EXIBIÇÃO - DIAGNÓSTICO DETALHADO
# ============================================================

Head "1. SISTEMA OPERACIONAL"
INF "Sistema : $osName"
INF "Build   : $osBuild  |  Arquit.: $osArch"
$osOk = $REQ_OS_OK | Where-Object { $osName -like "*$_*" }
if ($osOk) { OK "$osName" }
else        { FAIL "SO não suportado: $osName (requer Windows 10 ou 11 Pro/Enterprise)" }

Head "2. PROCESSADOR"
INF "Modelo  : $cpuName"
INF "Cores   : $cpuCores físicos / $cpuLogic lógicos  |  Clock: $cpuGHz GHz"
if ($cpuCores -ge 4) { OK "$cpuCores núcleos" }
else                  { WARN "Apenas $cpuCores núcleo(s) — pode impactar desempenho" }

Head "3. MEMÓRIA RAM"
INF "Total   : $ramTotalGB GB  |  Livre: $ramFreeGB GB"
if ($ramTotalGB -ge $REQ_RAM_GB_MIN) { OK "$ramTotalGB GB RAM" }
else                                  { FAIL "RAM insuficiente: $ramTotalGB GB (mínimo $REQ_RAM_GB_MIN GB)" }

Head "4. GPU"
foreach ($g in $gpus) {
    $vram = [math]::Round($g.AdapterRAM / 1GB, 1)
    INF "Modelo  : $($g.Name.Trim())"
    INF "VRAM    : $vram GB  |  Res: $($g.CurrentHorizontalResolution)x$($g.CurrentVerticalResolution)"
    OK "$($g.Name.Trim())"
}

Head "5. FIREBIRD"
if ($fbVer) {
    INF "Versão  : $fbVer"
    OK "Firebird $fbVer detectado"
} else {
    FAIL "Firebird não encontrado ou não instalado"
}

Head "6. BANCO DE DADOS FCerta"
if ($fcertaRoot) {
    INF "Raiz    : $fcertaRoot"

    if ($dbPath) {
        INF "DB      : $dbPath"
        OK  "Banco de Dados: $dbSize"
    } else {
        WARN "alterdb.ib não encontrado em $fcertaRoot\db\"
    }

    if ($imPath) {
        INF "Imagem  : $imPath"
        OK  "Banco de Imagem: $imSize"
    } else {
        WARN "alterim.ib não encontrado em $fcertaRoot\db\"
    }

    if ($lastBackup) {
        INF "Backup  : $lastBackup (última data do Sobre)"
        OK  "Backup localizado"
    } else {
        WARN "Pasta de backup não localizada"
    }

    # Espaço livre no disco da FCerta
    $drive     = Split-Path -Qualifier $fcertaRoot
    $driveInfo = Get-PSDrive -Name ($drive -replace ":","") -EA SilentlyContinue
    if ($driveInfo) {
        $freeGB  = [math]::Round($driveInfo.Free / 1GB, 1)
        $totalGB = [math]::Round(($driveInfo.Used + $driveInfo.Free) / 1GB, 1)
        INF "Disco   : $drive — $freeGB GB livres de $totalGB GB"
        if    ($freeGB -lt 10)  { FAIL "Espaço livre crítico: $freeGB GB (mínimo 10 GB)" }
        elseif($freeGB -lt 20)  { WARN "Espaço livre baixo: $freeGB GB" }
        else                    { OK   "Espaço livre: $freeGB GB" }
    }
} else {
    FAIL "Pasta FCerta não encontrada em nenhum disco local"
}

Head "7. REDE"
foreach ($a in $adapters) {
    $mbps    = Get-AdapterSpeedMbps $a
    $isWifi  = $a.PhysicalMediaType -like "*802.11*" -or $a.Name -match "Wi-?Fi|Wireless"
    $tipo    = if ($isWifi) { "Wireless" } else { "Cabeada" }
    $minReq  = if ($isWifi) { 108 } else { $REQ_NET_MBITS }

    INF "$tipo : $($a.Name) — $mbps Mbps"
    if ($mbps -ge $minReq) { OK  "$tipo OK: $mbps Mbps (mínimo $minReq Mbps)" }
    else                    { FAIL "$tipo insuficiente: $mbps Mbps (mínimo $minReq Mbps)" }
}

# ============================================================
#  RESULTADO FINAL - APTO OU NÃO
# ============================================================

# Avalia critérios obrigatórios
$checks = @{
    "SO compatível"         = ($osName -match "Windows (10|11)")
    "RAM suficiente"        = ($ramTotalGB -ge $REQ_RAM_GB_MIN)
    "Firebird instalado"    = ($fbVer -ne "")
    "FCerta localizado"     = ($fcertaRoot -ne $null)
    "Banco de dados (.ib)"  = ($dbPath -ne "")
    "Rede suficiente"       = (($adapters | Where-Object {
                                    $mbps = Get-AdapterSpeedMbps $_
                                    $isW  = $_.PhysicalMediaType -like "*802.11*" -or $_.Name -match "Wi-?Fi|Wireless"
                                    $min  = if ($isW) { 108 } else { $REQ_NET_MBITS }
                                    $mbps -ge $min
                                }).Count -gt 0)
}

$apto = -not ($checks.Values -contains $false)

Write-Host ""
Write-Host ("=" * 60) -ForegroundColor $(if ($apto) { "Green" } else { "Red" })
Write-Host ""
foreach ($k in $checks.Keys) {
    $v    = $checks[$k]
    $cor  = if ($v) { "Green" } else { "Red" }
    $icon = if ($v) { "[OK]" }  else { "[NOK]" }
    Write-Host "  $icon  $k" -ForegroundColor $cor
}
Write-Host ""
if ($apto) {
    Write-Host "  RESULTADO: AMBIENTE APTO para MobyCRM" -ForegroundColor Green
} else {
    Write-Host "  RESULTADO: AMBIENTE NAO APTO — corrija os itens [NOK]" -ForegroundColor Red
}
Write-Host ""
Write-Host ("=" * 60) -ForegroundColor $(if ($apto) { "Green" } else { "Red" })

# ============================================================
#  BRIEFING - layout para copiar/colar no ticket
# ============================================================

$briefing = @"

============================================================
 BRIEFING MobyCRM — $env:COMPUTERNAME
 Gerado em: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
============================================================

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

Sistema Operacional: $osName
Processador: $cpuName ($cpuCores cores / $cpuLogic threads @ $cpuGHz GHz)
Memória RAM: $ramTotalGB GB
GPU: $gpuList
Firebird: $fbVer
Banco de Dados: $dbSize
Banco de Imagem: $imSize
Pasta FCerta: $fcertaRoot

RESULTADO: $(if ($apto) { "APTO" } else { "NAO APTO" })

============================================================
"@

Write-Host $briefing -ForegroundColor White

# ============================================================
#  SALVA TXT
# ============================================================

# Reconstrói checklist para o arquivo
$checklistTxt = ($checks.GetEnumerator() | ForEach-Object {
    $icon = if ($_.Value) { "[OK]" } else { "[NOK]" }
    "  $icon  $($_.Key)"
}) -join "`n"

$fileTxt = @"
============================================================
 CHECK DE AMBIENTE - MobyCRM / FCerta / Phusion
 Computador : $env:COMPUTERNAME
 Usuário    : $env:USERNAME
 Gerado em  : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
============================================================

--- HARDWARE ---
Sistema Operacional : $osName (Build $osBuild) [$osArch]
Processador         : $cpuName ($cpuCores cores / $cpuLogic threads @ $cpuGHz GHz)
Memória RAM         : $ramTotalGB GB total / $ramFreeGB GB livre
GPU                 : $gpuList

--- AMBIENTE FCerta ---
Pasta raiz   : $fcertaRoot
Banco de Dados   : $dbSize  ($dbPath)
Banco de Imagem  : $imSize  ($imPath)
Firebird     : $fbVer
Último Backup: $lastBackup

--- REDE ---
$(($adapters | ForEach-Object {
    $mbps = Get-AdapterSpeedMbps $_
    "  $($_.Name): $mbps Mbps"
}) -join "`n")

--- CHECKLIST ---
$checklistTxt

--- RESULTADO ---
$(if ($apto) { "APTO PARA MobyCRM" } else { "NAO APTO — corrija os itens [NOK] acima" })

--- BRIEFING ---
$briefing
"@

$fileTxt | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host "  Relatorio salvo em: $ReportPath" -ForegroundColor Cyan
Write-Host ""
