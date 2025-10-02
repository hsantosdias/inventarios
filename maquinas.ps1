<#
Inventario de Patrimonio de TI (PowerShell) - v3b PT-BR
- Coleta: Hostname real, usuario logado, fabricante/modelo, serial BIOS, CPU, RAM, discos, GPU,
  SO, build, uptime, IP/MAC, data BIOS, bateria, BitLocker, dominio/workgroup.
- Lista softwares instalados (x64/x86).
- Saida: TXT (resumo detalhado), CSV (softwares), JSON (estrutura) em:
  C:\Temp\Inventario\<HOSTNAME>_<yyyyMMdd-HHmmss>\
- Opcional: copia para \\SERVIDOR\Inventario (ajuste abaixo).
- Prompts em PT-BR com confirmacao.
- Adicional: arquivo TXT com o NOME DO COMPUTADOR informado pelo operador + resumo do equipamento.
Compatibilidade: PowerShell 5+ (Windows 10/11).
#>

# =================== CONFIG ===================
$SaidaRaiz                  = "C:\Temp\Inventario"
$CopiarParaCompartilhamento = $true
$DestinoCompartilhamento    = "\\192.168.1.101\Inventario"

# Parametros de avaliacao
$DepreciacaoMesesPadrao     = 48    # 36/48/60
$PisoResidualPercentual     = 0.15  # 15%
# ==============================================

# =================== FUNCOES ===================
function Try-Get { param([scriptblock]$Block) try { & $Block } catch { $null } }

function Convert-DmtfDate {
  param([string]$dmtf)
  if ([string]::IsNullOrWhiteSpace($dmtf)) { return $null }
  if ($dmtf -match '^\*+$' -or $dmtf -match '^0+$') { return $null }
  try { [Management.ManagementDateTimeConverter]::ToDateTime($dmtf) } catch { $null }
}

function Fmt-Date($d) {
  if ($d) { $d.ToString("dd/MM/yyyy HH:mm:ss") } else { "N/D" }
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
# ==============================================

# =================== PROMPTS (PT-BR + confirmacao) ===================
do {
  Clear-Host
  Write-Host "== Cadastro de Patrimonio =="

  $NomeComputadorInformado = Read-Host "Nome do computador (etiqueta/operador)"
  $Patrimonio     = Read-Host "Numero do patrimonio (ex.: 2025-00123)"
  $Local          = Read-Host "Local/Setor (ex.: Robotica / Sala 12)"
  $Responsavel    = Read-Host "Responsavel (ex.: Maria Silva)"
  $EstadoStr      = Read-Host "Condicao geral (Novo/Bom/Regular/Ruim) [opcional]"
  $DataCompraStr  = Read-Host "Data de compra (dd/mm/aaaa) [opcional]"
  $PrecoCompraStr = Read-Host "Preco de compra (ex.: 3499,90) [opcional]"
  $Notas          = Read-Host "Observacoes [opcional]"

  # parse/normalize
  $Estado      = if ($EstadoStr) { $EstadoStr.Trim() } else { $null }
  $DataCompra  = Parse-Data $DataCompraStr
  $PrecoCompra = Parse-Preco $PrecoCompraStr

  # preview
  $dcText = if ($DataCompra) { Fmt-Date $DataCompra } else { "N/D" }
  $pcText = if ($PrecoCompra) { "R$ {0:N2}" -f $PrecoCompra } else { "N/D" }
  $estadoText = if ($Estado) { $Estado } else { "N/D" }
  $notasText  = if ($Notas)  { $Notas }  else { "-" }

  Write-Host ""
  Write-Host "== Confirme os dados =="
  Write-Host ("Nome informado     : {0}" -f $NomeComputadorInformado)
  Write-Host ("Patrimonio         : {0}" -f $Patrimonio)
  Write-Host ("Local/Setor        : {0}" -f $Local)
  Write-Host ("Responsavel        : {0}" -f $Responsavel)
  Write-Host ("Condicao           : {0}" -f $estadoText)
  Write-Host ("Data de compra     : {0}" -f $dcText)
  Write-Host ("Preco de compra    : {0}" -f $pcText)
  Write-Host ("Observacoes        : {0}" -f $notasText)
  Write-Host ""
  $confirm = Read-Host "Digite 1 para CONFIRMAR e continuar, ou 2 para REINICIAR"

} while ($confirm -ne '1')
# ====================================================================

# ============== COLETA (CIM/WMI) ==============
$stamp    = Get-Date -Format "yyyyMMdd-HHmmss"
$hostName = $env:COMPUTERNAME
$saidaDir = Join-Path $SaidaRaiz ("{0}_{1}" -f $hostName,$stamp)
New-Item -ItemType Directory -Path $saidaDir -Force | Out-Null

$os     = Try-Get { Get-CimInstance Win32_OperatingSystem }
$cs     = Try-Get { Get-CimInstance Win32_ComputerSystem }
$bios   = Try-Get { Get-CimInstance Win32_BIOS }
$cpu    = Try-Get { Get-CimInstance Win32_Processor | Select-Object -First 1 }
$ramGB  = if ($cs -and $cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory/1GB,2) } else { $null }
$gpus   = Try-Get { Get-CimInstance Win32_VideoController }
$disks  = Try-Get { Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" }
$nets   = Try-Get { Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true" }

# Datas / uptime
$lastBoot      = if ($os -and ($os.LastBootUpTime -is [datetime])) { $os.LastBootUpTime } else { if ($os) { Convert-DmtfDate $os.LastBootUpTime } else { $null } }
$uptime        = if ($lastBoot) { (Get-Date) - $lastBoot } else { $null }
$uptimeText    = if ($uptime) { "{0}d {1}h {2}m" -f $uptime.Days,$uptime.Hours,$uptime.Minutes } else { "N/D" }
$uptimeMin     = if ($uptime) { [math]::Round($uptime.TotalMinutes) } else { $null }
$osInstallDate = if ($os -and ($os.InstallDate -is [datetime])) { $os.InstallDate } else { if ($os) { Convert-DmtfDate $os.InstallDate } else { $null } }
$biosDate      = if ($bios -and ($bios.ReleaseDate -is [datetime])) { $bios.ReleaseDate } else { if ($bios) { Convert-DmtfDate $bios.ReleaseDate } else { $null } }

# BitLocker
$bitlockerInfo = $null
if (Get-Command -Name "manage-bde.exe" -ErrorAction SilentlyContinue) {
  $bitlockerInfo = Try-Get { & manage-bde.exe -status C: 2>$null }
}

# Bateria
$bateria = Try-Get { Get-CimInstance Win32_Battery }

# Dominio/Workgroup
$dominio   = if ($cs) { $cs.Domain } else { $null }
$workgroup = if ($cs -and $cs.Workgroup) { $cs.Workgroup } else { $null }

# Usuario logado
$usuarioAtual = Try-Get { (Get-CimInstance Win32_ComputerSystem).UserName }

# Discos
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

# Softwares
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
          Nome        = $_.DisplayName
          Versao      = $_.DisplayVersion
          Publicador  = $_.Publisher
          DataInstal  = $_.InstallDate
          Desinstalar = $_.UninstallString
          ChaveReg    = $_.PSPath
        }
      }
    }
  }
}
$apps = Get-InstalledApps | Sort-Object Nome, Versao
# ==============================================

# =================== AVALIACAO =================
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

# =================== ARQUIVOS DE SAIDA ===================
$txtPath   = Join-Path $saidaDir ("{0}_resumo.txt"     -f $hostName)
$csvPath   = Join-Path $saidaDir ("{0}_softwares.csv"  -f $hostName)
$jsonPath  = Join-Path $saidaDir ("{0}_sistema.json"   -f $hostName)
# Novo: TXT com nome informado pelo operador
$rotuloPath = Join-Path $saidaDir ("{0}info.txt" -f ($NomeComputadorInformado -replace '[\\/:*?""<>|]','_'))

# TXT detalhado (PT-BR)
$sb = New-Object System.Text.StringBuilder
$null = $sb.AppendLine("==== Inventario de TI ====")
$null = $sb.AppendLine("Data/Hora: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))
$null = $sb.AppendLine("Nome informado: $NomeComputadorInformado")
$null = $sb.AppendLine("Patrimonio: $Patrimonio")
$null = $sb.AppendLine("Local/Setor: $Local")
$null = $sb.AppendLine("Responsavel: $Responsavel")
$null = $sb.AppendLine("Condicao: " + ($(if ($Estado) { $Estado } else { "N/D" })))
$null = $sb.AppendLine("Preco de compra: " + ($(if ($PrecoCompra) { ("R$ {0:N2}" -f $PrecoCompra) } else { "N/D" })))
$null = $sb.AppendLine("Data de compra: " + (Fmt-Date $DataCompra))
$null = $sb.AppendLine("Observacoes: " + ($(if ($Notas) { $Notas } else { "-" })))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("Hostname (real): $hostName")
$null = $sb.AppendLine("Usuario logado: $usuarioAtual")
$null = $sb.AppendLine("Fabricante/Modelo: $($cs.Manufacturer) / $($cs.Model)")
$null = $sb.AppendLine("Serial da BIOS: $($bios.SerialNumber)")
$null = $sb.AppendLine("CPU: $($cpu.Name)  | Nucleos: $($cpu.NumberOfCores)  | Threads: $($cpu.NumberOfLogicalProcessors)")
$null = $sb.AppendLine("Memoria RAM: ${ramGB} GB")
$null = $sb.AppendLine("SO: $($os.Caption) ($($os.OSArchitecture))  Build: $($os.BuildNumber)")
$null = $sb.AppendLine("Instalacao do SO: " + (Fmt-Date $osInstallDate))
$null = $sb.AppendLine("Ultima inicializacao: " + (Fmt-Date $lastBoot) + "  | Uptime: $uptimeText")
$null = $sb.AppendLine("Dominio: $dominio  | Grupo de trabalho: $workgroup")
$null = $sb.AppendLine("BIOS: $($bios.SMBIOSBIOSVersion)  | Data BIOS: " + (Fmt-Date $biosDate))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Discos ==")
$disksView | ForEach-Object { $null = $sb.AppendLine("  - $($_.Unidade): Tam=$($_.TamanhoGB)GB, Livre=$($_.LivreGB)GB, Uso=$($_.UsoPct)%  FS=$($_.FS)  Rotulo=$($_.Label)") }
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
$null = $sb.AppendLine("== Video ==")
$gpuView | ForEach-Object { $null = $sb.AppendLine("  - $($_.Nome) (Driver: $($_.Driver), VRAM: $($_.VRAMMB) MB)") }
$null = $sb.AppendLine("")
if ($bitlockerInfo) {
  $null = $sb.AppendLine("== BitLocker (C:) ==")
  $null = $sb.AppendLine(($bitlockerInfo | Out-String))
}
if ($bateria) {
  $null = $sb.AppendLine("== Bateria ==")
  $null = $sb.AppendLine("  Status: $($bateria.BatteryStatus) | Carga: $($bateria.EstimatedChargeRemaining)%  | Restante: $($bateria.EstimatedRunTime) min")
}
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Softwares instalados ==")
$null = $sb.AppendLine("  Total: " + ($apps | Measure-Object).Count)
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Avaliacao (estimativa) ==")
$null = $sb.AppendLine("  Base para idade: " + ($(if ($DataBaseIdade) { (Fmt-Date $DataBaseIdade) } else { "N/D" })))
$null = $sb.AppendLine("  Meses de uso: " + ($(if ($MesesUso -ne $null) { $MesesUso } else { "N/D" })))
$null = $sb.AppendLine("  Depreciacao (meses): $DepreciacaoMeses")
$null = $sb.AppendLine("  Piso residual: $([int]($PisoResidualPercentual*100))%")
$null = $sb.AppendLine("  Valor estimado: " + ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D (informe preco e/ou data de compra)" })))

$sb.ToString() | Out-File -FilePath $txtPath -Encoding UTF8

# CSV (softwares)
$apps | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

# JSON estruturado
$payload = [pscustomobject]@{
  Timestamp    = Get-Date
  NomeInformado= $NomeComputadorInformado
  Patrimonio   = $Patrimonio
  Local        = $Local
  Responsavel  = $Responsavel
  Estado       = $Estado
  DataCompra   = $DataCompra
  PrecoCompra  = $PrecoCompra
  Observacoes  = $Notas
  Hostname     = $hostName
  Usuario      = $usuarioAtual
  Fabricante   = $cs.Manufacturer
  Modelo       = $cs.Model
  SerialBIOS   = $bios.SerialNumber
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

# ========= TXT adicional com o NOME informado (resumo curto) =========
$rotuloSb = New-Object System.Text.StringBuilder
$null = $rotuloSb.AppendLine("Resumo do equipamento")
$null = $rotuloSb.AppendLine("Gerado em: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))
$null = $rotuloSb.AppendLine("")
$null = $rotuloSb.AppendLine("Nome (operador): $NomeComputadorInformado")
$null = $rotuloSb.AppendLine("Hostname (real): $hostName")
$null = $rotuloSb.AppendLine("Patrimonio: $Patrimonio")
$null = $rotuloSb.AppendLine("Local/Setor: $Local")
$null = $rotuloSb.AppendLine("Responsavel: $Responsavel")
$null = $rotuloSb.AppendLine("Fabricante/Modelo: $($cs.Manufacturer) / $($cs.Model)")
$null = $rotuloSb.AppendLine("Serial BIOS: $($bios.SerialNumber)")
$null = $rotuloSb.AppendLine("CPU: $($cpu.Name)")
$null = $rotuloSb.AppendLine("RAM: ${ramGB} GB")
$null = $rotuloSb.AppendLine("SO: $($os.Caption) ($($os.OSArchitecture))")
$null = $rotuloSb.AppendLine("Data BIOS: " + (Fmt-Date $biosDate))
$null = $rotuloSb.AppendLine("Instalacao SO: " + (Fmt-Date $osInstallDate))
$null = $rotuloSb.AppendLine("Ultima inicializacao: " + (Fmt-Date $lastBoot))
$null = $rotuloSb.AppendLine("IP(s): " + (($redeView.IPv4 | Where-Object { $_ }) -join ", "))
$null = $rotuloSb.AppendLine("Valor estimado: " + ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D" })))
$rotuloSb.ToString() | Out-File -FilePath $rotuloPath -Encoding UTF8
# =====================================================================

# Copia para compartilhamento (opcional)
if ($CopiarParaCompartilhamento -and (Test-Path $DestinoCompartilhamento)) {
  try {
    Copy-Item -Path $saidaDir -Destination $DestinoCompartilhamento -Recurse -Force
  } catch {
    Write-Warning ("Falha ao copiar para {0}: {1}" -f $DestinoCompartilhamento, $_.Exception.Message)
  }
}

Write-Host "Inventario concluido. Arquivos em: $saidaDir"
if ($CopiarParaCompartilhamento) { Write-Host "Copia tentada para: $DestinoCompartilhamento" }
