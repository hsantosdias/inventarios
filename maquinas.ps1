<#
IT Asset Inventory (PowerShell) - v3a (ASCII safe)
- Collects: Hostname, logged user, make/model, BIOS serial, CPU, RAM, disks, GPU, OS, build, uptime, IP/MAC, BIOS date, battery, BitLocker, domain/workgroup.
- Lists installed software (x64/x86).
- Output: TXT (summary), CSV (software), JSON (structured) at C:\Temp\Inventario\<HOSTNAME>_<yyyyMMdd-HHmmss>\
- Optional copy to \\SERVER\Inventario (adjust below).
- Prompts for Asset ID and valuation fields. Linear depreciation (configurable).
- Tested on PowerShell 5+ (Windows 10/11).
#>

#cmdkey /add:192.168.1.101 /user:suporte /pass:027500

# =================== CONFIG ===================
$SaidaRaiz                  = "C:\Temp\Inventario"
$CopiarParaCompartilhamento = $true       # set $false if you do not want to copy to share
$DestinoCompartilhamento    = "\\192.168.1.101\Inventario"  # SMB share path

# Valuation parameters
$DepreciacaoMesesPadrao     = 48   # typical: 36/48/60
$PisoResidualPercentual     = 0.15 # 15% residual floor
# ==============================================

# =================== FUNCTIONS ===================
function Try-Get { param([scriptblock]$Block) try { & $Block } catch { $null } }

function Convert-DmtfDate {
  param([string]$dmtf)
  if ([string]::IsNullOrWhiteSpace($dmtf)) { return $null }
  if ($dmtf -match '^\*+$' -or $dmtf -match '^0+$') { return $null }
  try { [Management.ManagementDateTimeConverter]::ToDateTime($dmtf) } catch { $null }
}

function Fmt-Date($d) {
  if ($d) { $d.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/D" }
}

function Parse-Data {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $dt = $null
  $formats = @("dd/MM/yyyy","d/M/yyyy","yyyy-MM-dd","dd-MM-yyyy")
  foreach ($f in $formats) { if ([datetime]::TryParseExact($s,$f,$null,[System.Globalization.DateTimeStyles]::AssumeLocal,[ref]$dt)) { return $dt } }
  if ([datetime]::TryParse($s,[ref]$dt)) { return $dt }
  return $null
}

function Parse-Preco {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $clean = ($s -replace '[^\d,.\-]','').Trim()
  if ($clean -match ',\d{1,2}$' -and $clean -match '\.') { $clean = $clean -replace '\.','' -replace ',','.' }
  elseif ($clean -match ',\d{1,2}$') { $clean = $clean -replace ',','.' }
  [decimal]$v = $null
  if ([decimal]::TryParse($clean,[ref]$v)) { return $v } else { return $null }
}

function Calc-Meses {
  param([datetime]$from,[datetime]$to)
  if (-not $from -or -not $to) { return $null }
  $months = (($to.Year - $from.Year) * 12) + ($to.Month - $from.Month)
  if ($to.Day -lt $from.Day) { $months-- }
  if ($months -lt 0) { $months = 0 }
  return $months
}
# ================================================

# =================== PROMPTS ===================
# =================== PROMPTS (with confirmation) ===================
do {
  Clear-Host
  Write-Host "== Asset Enrollment =="

  $Patrimonio     = Read-Host "Asset ID (e.g., 2025-00123)"
  $Local          = Read-Host "Location/Department (e.g., Lab Robotica / Room 12)"
  $Responsavel    = Read-Host "Responsible person (e.g., Maria Silva)"
  $EstadoStr      = Read-Host "General condition (New/Good/Fair/Poor) [optional]"
  $DataCompraStr  = Read-Host "Purchase date (dd/mm/yyyy) [optional]"
  $PrecoCompraStr = Read-Host "Purchase price (e.g., 3499,90) [optional]"
  $Notas          = Read-Host "Notes [optional]"

  # parse/normalize
  $Estado      = if ($EstadoStr) { $EstadoStr.Trim() } else { $null }
  $DataCompra  = Parse-Data $DataCompraStr
  $PrecoCompra = Parse-Preco $PrecoCompraStr

  # preview text (avoid passing $null into Fmt-Date here)
  $dcText = if ($DataCompra) { Fmt-Date $DataCompra } else { "N/D" }
  $pcText = if ($PrecoCompra) { "R$ {0:N2}" -f $PrecoCompra } else { "N/D" }
  $estadoText = if ($Estado) { $Estado } else { "N/D" }
  $notasText  = if ($Notas)  { $Notas }  else { "-" }

  Write-Host ""
  Write-Host "== Review your entries =="
  Write-Host ("Asset ID           : {0}" -f $Patrimonio)
  Write-Host ("Location           : {0}" -f $Local)
  Write-Host ("Responsible        : {0}" -f $Responsavel)
  Write-Host ("Condition          : {0}" -f $estadoText)
  Write-Host ("Purchase date      : {0}" -f $dcText)
  Write-Host ("Purchase price     : {0}" -f $pcText)
  Write-Host ("Notes              : {0}" -f $notasText)
  Write-Host ""
  $confirm = Read-Host "Type 1 to CONFIRM and continue, 2 to RESTART (or 3 to CANCEL)"

  if ($confirm -eq '3') {
    Write-Host "Process canceled by user."
    exit 0
  }
} while ($confirm -ne '1')
# ==================================================================

# ==============================================

# ============== COLLECTION (CIM/WMI) ===========
$stamp    = Get-Date -Format "yyyyMMdd-HHmmss"
$hostName = $env:COMPUTERNAME
$saidaDir = Join-Path $SaidaRaiz ("{0}_{1}" -f $hostName,$stamp)
New-Item -ItemType Directory -Path $saidaDir -Force | Out-Null

$os     = Try-Get { Get-CimInstance Win32_OperatingSystem }
$cs     = Try-Get { Get-CimInstance Win32_ComputerSystem }
$bios   = Try-Get { Get-CimInstance Win32_BIOS }
$cpu    = Try-Get { Get-CimInstance Win32_Processor | Select-Object -First 1 }
$ramGB  = if ($cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory/1GB,2) } else { $null }
$gpus   = Try-Get { Get-CimInstance Win32_VideoController }
$disks  = Try-Get { Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" }
$nets   = Try-Get { Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true" }

# Dates / uptime
$lastBoot      = if ($os.LastBootUpTime -is [datetime]) { $os.LastBootUpTime } else { Convert-DmtfDate $os.LastBootUpTime }
$uptime        = if ($lastBoot) { (Get-Date) - $lastBoot } else { $null }
$uptimeText    = if ($uptime) { "{0}d {1}h {2}m" -f $uptime.Days,$uptime.Hours,$uptime.Minutes } else { "N/D" }
$uptimeMin     = if ($uptime) { [math]::Round($uptime.TotalMinutes) } else { $null }
$osInstallDate = if ($os.InstallDate   -is [datetime]) { $os.InstallDate } else { Convert-DmtfDate $os.InstallDate }
$biosDate      = if ($bios.ReleaseDate -is [datetime]) { $bios.ReleaseDate } else { Convert-DmtfDate $bios.ReleaseDate }

# BitLocker
$bitlockerInfo = $null
if (Get-Command -Name "manage-bde.exe" -ErrorAction SilentlyContinue) {
  $bitlockerInfo = Try-Get { & manage-bde.exe -status C: 2>$null }
}

# Battery
$bateria = Try-Get { Get-CimInstance Win32_Battery }

# Domain/Workgroup
$dominio   = $cs.Domain
$workgroup = if ($cs.Workgroup) { $cs.Workgroup } else { $null }

# Logged user
$usuarioAtual = Try-Get { (Get-CimInstance Win32_ComputerSystem).UserName }

# Disks view
$disksView = $disks | ForEach-Object {
  [pscustomobject]@{
    Unidade   = $_.DeviceID
    TamanhoGB = [math]::Round( ($_.Size/1GB), 2 )
    LivreGB   = [math]::Round( ($_.FreeSpace/1GB), 2 )
    UsoPct    = if($_.Size){ [math]::Round( (1-($_.FreeSpace/$_.Size))*100, 1) } else { $null }
    FS        = $_.FileSystem
    Label     = $_.VolumeName
  }
}

# Network view
$redeView = $nets | ForEach-Object {
  [pscustomobject]@{
    Descricao = $_.Description
    IPv4      = ($_.IPAddress | Where-Object { $_ -match '^\d+\.' }) -join ', '
    IPv6      = ($_.IPAddress | Where-Object { $_ -match ':' }) -join ', '
    MAC       = $_.MACAddress
    Gateway   = ($_.DefaultIPGateway -join ', ')
    DNS       = ($_.DNSServerSearchOrder -join ', ')
  }
}

# GPU view
$gpuView = $gpus | ForEach-Object {
  [pscustomobject]@{
    Nome    = $_.Name
    Driver  = $_.DriverVersion
    VRAMMB  = if ($_.AdapterRAM) { [math]::Round($_.AdapterRAM/1MB) } else { $null }
  }
}

# Installed software
function Get-InstalledApps {
  $paths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
  )
  foreach ($p in $paths) {
    Try-Get {
      Get-ItemProperty $p | Where-Object { $_.DisplayName } | ForEach-Object {
        [pscustomobject]@{
          Name        = $_.DisplayName
          Version     = $_.DisplayVersion
          Publisher   = $_.Publisher
          InstallDate = $_.InstallDate
          Uninstall   = $_.UninstallString
          RegistryKey = $_.PSPath
        }
      }
    }
  }
}
$apps = Get-InstalledApps | Sort-Object Name, Version
# ==============================================

# =================== VALUATION =================
$hoje = Get-Date
$DataBaseIdade = $DataCompra
if (-not $DataBaseIdade) { $DataBaseIdade = $biosDate }
if (-not $DataBaseIdade) { $DataBaseIdade = $osInstallDate }
$MesesUso = Calc-Meses $DataBaseIdade $hoje

$ValorEstimado     = $null
$DepreciacaoMeses  = $DepreciacaoMesesPadrao
$PrecoCompraBase   = $PrecoCompra

if ($PrecoCompraBase -and $MesesUso -ne $null) {
  $deprMensal = [decimal]($PrecoCompraBase / $DepreciacaoMeses)
  $acumulada  = [decimal]($deprMensal * [Math]::Min($MesesUso,$DepreciacaoMeses))
  $valorResid = [decimal]([Math]::Max(($PrecoCompraBase * (1 - $PisoResidualPercentual)), $PrecoCompraBase - $acumulada))
  $pisoAbs    = [decimal]($PrecoCompraBase * $PisoResidualPercentual)
  if ($valorResid -lt $pisoAbs) { $valorResid = $pisoAbs }
  $ValorEstimado = [decimal]([Math]::Round($valorResid,2))
}
# ==============================================

# =================== OUTPUT FILES ===================
$txtPath  = Join-Path $saidaDir ("{0}_resumo.txt"    -f $hostName)
$csvPath  = Join-Path $saidaDir ("{0}_softwares.csv" -f $hostName)
$jsonPath = Join-Path $saidaDir ("{0}_sistema.json"  -f $hostName)

# TXT
$sb = New-Object System.Text.StringBuilder
$null = $sb.AppendLine("==== IT Inventory ====")
$null = $sb.AppendLine("Date/Time: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
$null = $sb.AppendLine("Asset ID: $Patrimonio")
$null = $sb.AppendLine("Location: $Local")
$null = $sb.AppendLine("Owner/Responsible: $Responsavel")
$null = $sb.AppendLine("Condition: " + ($(if ($Estado) { $Estado } else { "N/D" })))
$null = $sb.AppendLine("Purchase price: " + ($(if ($PrecoCompra) { ("R$ {0:N2}" -f $PrecoCompra) } else { "N/D" })))
$null = $sb.AppendLine("Purchase date: " + (Fmt-Date $DataCompra))
$null = $sb.AppendLine("Notes: " + ($(if ($Notas) { $Notas } else { "-" })))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("Hostname: $hostName")
$null = $sb.AppendLine("Logged user: $usuarioAtual")
$null = $sb.AppendLine("Make/Model: $($cs.Manufacturer) / $($cs.Model)")
$null = $sb.AppendLine("BIOS Serial: $($bios.SerialNumber)")
$null = $sb.AppendLine("CPU: $($cpu.Name)  | Cores: $($cpu.NumberOfCores)  | Threads: $($cpu.NumberOfLogicalProcessors)")
$null = $sb.AppendLine("RAM: ${ramGB} GB")
$null = $sb.AppendLine("OS: $($os.Caption) ($($os.OSArchitecture))  Build: $($os.BuildNumber)")
$null = $sb.AppendLine("OS Install: " + (Fmt-Date $osInstallDate))
$null = $sb.AppendLine("Last boot: " + (Fmt-Date $lastBoot) + "  | Uptime: $uptimeText")
$null = $sb.AppendLine("Domain: $dominio  | Workgroup: $workgroup")
$null = $sb.AppendLine("BIOS: $($bios.SMBIOSBIOSVersion)  | BIOS Date: " + (Fmt-Date $biosDate))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Disks ==")
$disksView | ForEach-Object { $null = $sb.AppendLine("  - $($_.Unidade): Size=$($_.TamanhoGB)GB, Free=$($_.LivreGB)GB, Use=$($_.UsoPct)%  FS=$($_.FS)  Label=$($_.Label)") }
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Network ==")
$redeView | ForEach-Object {
  $null = $sb.AppendLine("  - $($_.Descricao)")
  $null = $sb.AppendLine("      IPv4: $($_.IPv4)")
  $null = $sb.AppendLine("      IPv6: $($_.IPv6)")
  $null = $sb.AppendLine("      MAC : $($_.MAC)")
  $null = $sb.AppendLine("      GW  : $($_.Gateway)")
  $null = $sb.AppendLine("      DNS : $($_.DNS)")
}
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Video ==")
$gpuView | ForEach-Object { $null = $sb.AppendLine("  - $($_.Nome) (Driver: $($_.Driver), VRAM: $($_.VRAMMB) MB)") }
$null = $sb.AppendLine("")
if ($bitlockerInfo) {
  $null = $sb.AppendLine("== BitLocker (C:) ==")
  $null = $sb.AppendLine(($bitlockerInfo | Out-String))
}
if ($bateria) {
  $null = $sb.AppendLine("== Battery ==")
  $null = $sb.AppendLine("  Status: $($bateria.BatteryStatus) | Charge: $($bateria.EstimatedChargeRemaining)%  | Remaining: $($bateria.EstimatedRunTime) min")
}
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Installed software ==")
$null = $sb.AppendLine("  Total: " + ($apps | Measure-Object).Count)
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Valuation (estimate) ==")
$null = $sb.AppendLine("  Age base: " + ($(if ($DataBaseIdade) { (Fmt-Date $DataBaseIdade) } else { "N/D" })))
$null = $sb.AppendLine("  Months in use: " + ($(if ($MesesUso -ne $null) { $MesesUso } else { "N/D" })))
$null = $sb.AppendLine("  Depreciation (months): $DepreciacaoMeses")
$null = $sb.AppendLine("  Residual floor: $([int]($PisoResidualPercentual*100))%")
$null = $sb.AppendLine("  Estimated value: " + ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D (inform price/purchase date for full calc)" })))

$sb.ToString() | Out-File -FilePath $txtPath -Encoding UTF8

# CSV (software)
$apps | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

# JSON
$payload = [pscustomobject]@{
  Timestamp    = Get-Date
  Patrimonio   = $Patrimonio
  Local        = $Local
  Responsavel  = $Responsavel
  Estado       = $Estado
  DataCompra   = $DataCompra
  PrecoCompra  = $PrecoCompra
  Notas        = $Notas
  Hostname     = $hostName
  Usuario      = $usuarioAtual
  Fabricante   = $cs.Manufacturer
  Modelo       = $cs.Model
  Serial       = $bios.SerialNumber
  CPU          = $cpu.Name
  Cores        = $cpu.NumberOfCores
  Threads      = $cpu.NumberOfLogicalProcessors
  RAM_GB       = $ramGB
  SO           = $os.Caption
  Arquitetura  = $os.OSArchitecture
  Build        = $os.BuildNumber
  OSInstall    = $osInstallDate
  BIOSDate     = $biosDate
  LastBoot     = $lastBoot
  UptimeMin    = $uptimeMin
  Dominio      = $dominio
  Workgroup    = $workgroup
  GPUs         = $gpuView
  Discos       = $disksView
  Redes        = $redeView
  BaseIdade    = $DataBaseIdade
  MesesUso     = $MesesUso
  DepMesesCfg  = $DepreciacaoMeses
  PisoResidual = $PisoResidualPercentual
  ValorEstimado= $ValorEstimado
}
$payload | Add-Member -NotePropertyName OSInstallText -NotePropertyValue (Fmt-Date $osInstallDate)
$payload | Add-Member -NotePropertyName BIOSDateText   -NotePropertyValue (Fmt-Date $biosDate)
$payload | Add-Member -NotePropertyName LastBootText   -NotePropertyValue (Fmt-Date $lastBoot)
$payload | Add-Member -NotePropertyName DataCompraText -NotePropertyValue (Fmt-Date $DataCompra)

$payload | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

# Copy to share (optional)
if ($CopiarParaCompartilhamento -and (Test-Path $DestinoCompartilhamento)) {
  try {
    Copy-Item -Path $saidaDir -Destination $DestinoCompartilhamento -Recurse -Force
  } catch {
    Write-Warning ("Failed to copy to {0}: {1}" -f $DestinoCompartilhamento, $_.Exception.Message)
  }
}

Write-Host "Inventory completed. Files at: $saidaDir"
if ($CopiarParaCompartilhamento) { Write-Host "Copy attempted to: $DestinoCompartilhamento" }
