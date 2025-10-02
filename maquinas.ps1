<# 
Inventário de Patrimônio de TI (PowerShell)
- Coleta: Hostname, usuário logado, fabricante/modelo, serial (BIOS), CPU, RAM, disco, GPU, SO, build, uptime, IP/MAC, BIOS/UEFI, bateria (se houver), BitLocker (se disponível), domínios.
- Lista softwares instalados (x64 e x86) com Nome, Versão, Publisher, Instalação.
- Saída: TXT (resumo) e CSV (softwares) em C:\Temp\Inventario\<HOSTNAME>_<yyyyMMdd-HHmmss>\
- Opcional: copia para \\SERVIDOR\Inventario$ (ajuste abaixo).
Compatível com PowerShell 5+ (Windows 10/11).
#>

# ======= CONFIGURAÇÕES =======
$SaidaRaiz = "C:\Temp\Inventario"
$CopiarParaCompartilhamento = $true         # defina $false se não for copiar
$DestinoCompartilhamento = "\\192.168.1.101\Inventario"  # ajuste para seu SMB
# =============================

# Garantir pasta
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$hostName = $env:COMPUTERNAME
$saidaDir = Join-Path $SaidaRaiz ("{0}_{1}" -f $hostName,$stamp)
New-Item -ItemType Directory -Path $saidaDir -Force | Out-Null

# Funções utilitárias
function Try-Get { param([scriptblock]$Block) try { & $Block } catch { $null } }

function Convert-DmtfDate {
  param([string]$dmtf)
  if ([string]::IsNullOrWhiteSpace($dmtf)) { return $null }
  if ($dmtf -match '^\*+$' -or $dmtf -match '^0+$') { return $null }
  try { [Management.ManagementDateTimeConverter]::ToDateTime($dmtf) } catch { $null }
}

function Fmt-Date([datetime]$d) { if ($d) { $d.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/D" } }

# Coletas principais (CIM/WMI)
$os     = Try-Get { Get-CimInstance Win32_OperatingSystem }
$cs     = Try-Get { Get-CimInstance Win32_ComputerSystem }
$bios   = Try-Get { Get-CimInstance Win32_BIOS }
$cpu    = Try-Get { Get-CimInstance Win32_Processor | Select-Object -First 1 }
$ramGB  = if ($cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory/1GB,2) } else { $null }
$gpus   = Try-Get { Get-CimInstance Win32_VideoController }
$disks  = Try-Get { Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" }
$nets   = Try-Get { Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true" }

# Uptime + datas legíveis (tolerantes a DMTF inválido)
$lastBoot      = if ($os.LastBootUpTime -is [datetime]) { $os.LastBootUpTime } else { Convert-DmtfDate $os.LastBootUpTime }
$uptime        = if ($lastBoot) { (Get-Date) - $lastBoot } else { $null }
$osInstallDate = if ($os.InstallDate     -is [datetime]) { $os.InstallDate }     else { Convert-DmtfDate $os.InstallDate }
$biosDate      = if ($bios.ReleaseDate   -is [datetime]) { $bios.ReleaseDate }   else { Convert-DmtfDate $bios.ReleaseDate }

# BitLocker (se disponível)
$bitlockerInfo = $null
if (Get-Command -Name "manage-bde.exe" -ErrorAction SilentlyContinue) {
  $bitlockerInfo = Try-Get { & manage-bde.exe -status C: 2>$null }
}

# Bateria (se notebook)
$bateria = Try-Get { Get-CimInstance Win32_Battery }

# Domínio/Workgroup
$dominio = $cs.Domain
$workgroup = if ($cs.Workgroup) { $cs.Workgroup } else { $null }

# Usuário atual
$usuarioAtual = Try-Get { (Get-CimInstance Win32_ComputerSystem).UserName }

# Sistema de Arquivos/Discos
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

# Rede
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

# GPUs
$gpuView = $gpus | ForEach-Object {
  [pscustomobject]@{
    Nome    = $_.Name
    Driver  = $_.DriverVersion
    VRAMMB  = if ($_.AdapterRAM) { [math]::Round($_.AdapterRAM/1MB) } else { $null }
  }
}

# Softwares instalados (x64 e x86)
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

# Resumo legível (TXT)
$txtPath = Join-Path $saidaDir ("{0}_resumo.txt" -f $hostName)
$sb = New-Object System.Text.StringBuilder
$nl = [Environment]::NewLine
$null = $sb.AppendLine("==== Inventário de TI ====")
$null = $sb.AppendLine("Data/Hora: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$null = $sb.AppendLine("Hostname: $hostName")
$null = $sb.AppendLine("Usuário atual: $usuarioAtual")
$null = $sb.AppendLine("Fabricante/Modelo: $($cs.Manufacturer) / $($cs.Model)")
$null = $sb.AppendLine("Serial (BIOS): $($bios.SerialNumber)")
$null = $sb.AppendLine("CPU: $($cpu.Name)  | Núcleos: $($cpu.NumberOfCores)  | Threads: $($cpu.NumberOfLogicalProcessors)")
$null = $sb.AppendLine("Memória RAM: ${ramGB} GB")
$null = $sb.AppendLine("SO: $($os.Caption) ($($os.OSArchitecture))  Build: $($os.BuildNumber)")
$null = $sb.AppendLine("Instalação SO: $(Fmt-Date $osInstallDate)")
$null = $sb.AppendLine("Última inicialização: $(Fmt-Date $lastBoot)  | Uptime: " + (if ($uptime) { "{0}d {1}h {2}m" -f $uptime.Days,$uptime.Hours,$uptime.Minutes } else { "N/D" }))
$null = $sb.AppendLine("Domínio: $dominio  | Workgroup: $workgroup")
$null = $sb.AppendLine("BIOS: $($bios.SMBIOSBIOSVersion)  | Data BIOS: $(Fmt-Date $biosDate)")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Vídeo ==")
$gpuView | ForEach-Object { $null = $sb.AppendLine("  - $($_.Nome) (Driver: $($_.Driver), VRAM: $($_.VRAMMB) MB)") }
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Discos (unidades lógicas) ==")
$disksView | ForEach-Object {
  $null = $sb.AppendLine("  - $($_.Unidade): Tamanho=$($_.TamanhoGB)GB, Livre=$($_.LivreGB)GB, Uso=$($_.UsoPct)%  FS=$($_.FS)  Label=$($_.Label)")
}
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Rede ==")
$redeView | ForEach-Object {
  $null = $sb.AppendLine("  - $($_.Descricao)")
  $null = $sb.AppendLine("      IPv4: $($_.IPv4)")
  $null = $sb.AppendLine("      IPv6: $($_.IPv6)")
  $null = $sb.AppendLine("      MAC : $($_.MAC)")
  $null = $sb.AppendLine("      GW  : $($_.Gateway)")
  $null = $sb.AppendLine("      DNS : $($_.DNS)")
}
$null = $sb.AppendLine("")
if ($bitlockerInfo) {
  $null = $sb.AppendLine("== BitLocker (C:) ==")
  $null = $sb.AppendLine(($bitlockerInfo | Out-String))
}
if ($bateria) {
  $null = $sb.AppendLine("== Bateria ==")
  $null = $sb.AppendLine("  Status: $($bateria.BatteryStatus) | Estimativa: $($bateria.EstimatedChargeRemaining)%  | Tempo restante: $($bateria.EstimatedRunTime) min")
}
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Contagem de softwares instalados ==")
$null = $sb.AppendLine("  Total: " + ($apps | Measure-Object).Count)

$sb.ToString() | Out-File -FilePath $txtPath -Encoding UTF8

# Softwares para CSV
$csvPath = Join-Path $saidaDir ("{0}_softwares.csv" -f $hostName)
$apps | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

# Salvar também um JSON com o bloco principal (útil para integração futura)
$jsonPath = Join-Path $saidaDir ("{0}_sistema.json" -f $hostName)
$payload = [pscustomobject]@{
  Timestamp   = Get-Date
  Hostname    = $hostName
  Usuario     = $usuarioAtual
  Fabricante  = $cs.Manufacturer
  Modelo      = $cs.Model
  Serial      = $bios.SerialNumber
  CPU         = $cpu.Name
  Cores       = $cpu.NumberOfCores
  Threads     = $cpu.NumberOfLogicalProcessors
  RAM_GB      = $ramGB
  SO          = $os.Caption
  Arquitetura = $os.OSArchitecture
  Build       = $os.BuildNumber
  LastBoot    = $lastBoot
  UptimeMin   = if ($uptime) { [math]::Round($uptime.TotalMinutes) } else { $null }
  Dominio     = $dominio
  Workgroup   = $workgroup
  GPUs        = $gpuView
  Discos      = $disksView
  Redes       = $redeView
}
# Campos de datas legíveis também no JSON
$payload | Add-Member -NotePropertyName OSInstallDateText -NotePropertyValue (Fmt-Date $osInstallDate)
$payload | Add-Member -NotePropertyName LastBootText      -NotePropertyValue (Fmt-Date $lastBoot)
$payload | Add-Member -NotePropertyName BIOSDateText      -NotePropertyValue (Fmt-Date $biosDate)

$payload | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

# Copiar para compartilhamento (opcional)
if ($CopiarParaCompartilhamento -and (Test-Path $DestinoCompartilhamento)) {
  try {
    Copy-Item -Path $saidaDir -Destination $DestinoCompartilhamento -Recurse -Force
  } catch {
    Write-Warning ("Falha ao copiar para {0}: {1}" -f $DestinoCompartilhamento, $_.Exception.Message)
  }
}

Write-Host "Inventário concluído. Arquivos em: $saidaDir"
if ($CopiarParaCompartilhamento) { Write-Host "Cópia tentada para: $DestinoCompartilhamento" }
