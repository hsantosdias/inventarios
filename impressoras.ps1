<#
Cadastro de Impressoras (PowerShell) - Repositório de Informações - v1.1 PT-BR
- NÃO coleta dados do computador. Apenas perguntas ao operador.
- Coleta: Identificação geral (patrimônio, local, responsável, etc).
- Específicos de impressora: marca/modelo, tipo (menu), cor (mono/color),
  tamanhos de papel, duplex, bandejas, conectividade (menu), rede (IP/Máscara/GW/MAC/host),
  série, drivers/fila, contadores, consumíveis, garantia/manutenção/status.
- Avaliação por depreciação linear (opcional).
- Saída: TXT detalhado, JSON compatível com padrão do seu sistema, TXT curto
  e TXT “IMPRESSORA_MARCA_MODELO_DATA”.
- Opcional: cópia para compartilhamento SMB com credenciais e limpeza de sessão/cache.

Compatibilidade: Windows PowerShell 5+ (Windows 10/11)
#>

# =================== CONFIG ===================
$SaidaRaiz                  = "C:\Temp\Inventario"

# Copiar para compartilhamento
$CopiarParaCompartilhamento = $true
$DestinoCompartilhamento    = "\\192.168.1.101\Dados\temp"

# Parâmetros de avaliação
$DepreciacaoMesesPadrao     = 48    # 36/48/60
$PisoResidualPercentual     = 0.15  # 15%
# ==============================================

# =================== FUNÇÕES ===================
function Try-Get { param([scriptblock]$Block) try { & $Block } catch { $null } }

function Fmt-Date($d) { if ($d) { $d.ToString("dd/MM/yyyy HH:mm:ss") } else { "N/D" } }

function Parse-Data {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $dt = $null
  $formats = @("dd/MM/yyyy","d/M/yyyy","yyyy-MM-dd","dd-MM-yyyy")
  foreach ($f in $formats) {
    if ([datetime]::TryParseExact($s,$f,$null,[System.Globalization.DateTimeStyles]::AssumeLocal,[ref]$dt)) { return $dt }
  }
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

function Get-UncHost {
  param([Parameter(Mandatory)][string]$UncPath)
  if ($UncPath -match '^\\\\([^\\]+)') { return $matches[1] } else { return $null }
}

function Copy-WithCredentialsAndCleanup {
  param(
    [Parameter(Mandatory)][string]$SourcePath,
    [Parameter(Mandatory)][string]$SharePath,
    [Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential
  )
  if (-not (Test-Path $SourcePath)) { throw "Caminho de origem não existe: $SourcePath" }

  $driveName = "INV"
  $uncServer = Get-UncHost -UncPath $SharePath

  try {
    if (Get-PSDrive -Name $driveName -ErrorAction SilentlyContinue) {
      Remove-PSDrive -Name $driveName -Force -ErrorAction SilentlyContinue
    }
    New-PSDrive -Name $driveName -PSProvider FileSystem -Root $SharePath -Credential $Credential -ErrorAction Stop | Out-Null

    $dest = "$driveName`:"
    Copy-Item -Path $SourcePath -Destination $dest -Recurse -Force -ErrorAction Stop
  }
  finally {
    if (Get-PSDrive -Name $driveName -ErrorAction SilentlyContinue) {
      Remove-PSDrive -Name $driveName -Force -ErrorAction SilentlyContinue
    }
    try { & cmd.exe /c "net use $SharePath /delete /y" | Out-Null } catch {}
    if ($uncServer) {
      try { & cmd.exe /c "cmdkey /delete:$uncServer" | Out-Null } catch {}
    }
  }
}
# ==============================================

# =================== MENUS =====================
$TiposImpressora = [ordered]@{
  1 = "Laser (mono)"
  2 = "Laser (color)"
  3 = "Jato de Tinta"
  4 = "Tanque de Tinta"
  5 = "Térmica (recibo)"
  6 = "Matricial"
  7 = "Multifuncional Laser"
  8 = "Multifuncional Jato"
  9 = "Plotter / Grande formato"
  10 = "Etiqueta / Código de barras"
  11 = "Fiscal / SAT"
  12 = "Outra"
}

$ConectividadeOptions = [ordered]@{
  1 = "USB"
  2 = "Ethernet (LAN)"
  3 = "Wi-Fi"
  4 = "Wi-Fi Direct"
  5 = "Bluetooth"
}

$ProtocolosOptions = [ordered]@{
  1 = "IPP/IPPS"
  2 = "RAW 9100 (JetDirect)"
  3 = "LPR/LPD"
  4 = "SMB/Compartilhamento"
  5 = "AirPrint"
  6 = "Mopria"
}

$TamanhosPapelOptions = [ordered]@{
  1 = "A4"
  2 = "A3"
  3 = "Carta"
  4 = "Ofício"
  5 = "10x15"
  6 = "Etiqueta/rolo"
}

$DuplexOptions = [ordered]@{
  1 = "Automático"
  2 = "Manual"
  3 = "Sem duplex"
}

$CorOptions = [ordered]@{
  1 = "Monocromática"
  2 = "Colorida"
}
# ==============================================

# =================== PROMPTS (PT-BR + confirmação) ===================
do {
  Clear-Host
  Write-Host "== Cadastro de Impressora =="

  # Identificação geral
  $NomeEquipamento = Read-Host "Nome do equipamento (etiqueta/operador)"
  $Patrimonio      = Read-Host "Número do patrimônio (ex.: 2025-00123)"
  $Local           = Read-Host "Local/Setor (ex.: Secretaria / Sala 12)"
  $Responsavel     = Read-Host "Responsável (ex.: Maria Silva)"
  $EstadoStr       = Read-Host "Condição geral (Novo/Bom/Regular/Ruim) [opcional]"
  $DataCompraStr   = Read-Host "Data de compra (dd/mm/aaaa) [opcional]"
  $PrecoCompraStr  = Read-Host "Preço de compra (ex.: 1299,90) [opcional]"
  $Notas           = Read-Host "Observações [opcional]"

  # Específicos da impressora
  $Marca           = Read-Host "Marca (ex.: HP, Brother, Epson, Canon)"
  $Modelo          = Read-Host "Modelo (ex.: M404dn, L3250)"
  $NumeroSerie     = Read-Host "Número de série [opcional]"

  Write-Host ""
  Write-Host "Tipo de impressora:"
  $TiposImpressora.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | % { Write-Host $_ }
  $TipoSel         = Read-Host "Escolha o tipo (número)"
  $TipoImpressora  = $TiposImpressora[[int]$TipoSel]

  Write-Host ""
  Write-Host "Cor:"
  $CorOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | % { Write-Host $_ }
  $CorSel          = Read-Host "1=Monocromática, 2=Colorida"
  $TipoCor         = $CorOptions[[int]$CorSel]

  Write-Host ""
  Write-Host "Duplex:"
  $DuplexOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | % { Write-Host $_ }
  $DuplexSel       = Read-Host "1=Automático, 2=Manual, 3=Sem duplex"
  $Duplex          = $DuplexOptions[[int]$DuplexSel]

  Write-Host ""
  Write-Host "Tamanhos de papel suportados (múltiplos separados por vírgula):"
  $TamanhosPapelOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | % { Write-Host $_ }
  $PapSel          = Read-Host "Ex.: 1,3,4"
  $TamanhosPapel   = @()
  if ($PapSel) {
    $TamanhosPapel = $PapSel -split ',' | % { $_.Trim() } | ? { $_ -match '^\d+$' } | % { $TamanhosPapelOptions[[int]$_] } | ? { $_ }
  }

  Write-Host ""
  Write-Host "Conectividade (múltiplos separados por vírgula):"
  $ConectividadeOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | % { Write-Host $_ }
  $ConSel          = Read-Host "Ex.: 1,2"
  $Conectividade   = @()
  if ($ConSel) {
    $Conectividade = $ConSel -split ',' | % { $_.Trim() } | ? { $_ -match '^\d+$' } | % { $ConectividadeOptions[[int]$_] } | ? { $_ }
  }

  Write-Host ""
  Write-Host "Protocolos (múltiplos separados por vírgula):"
  $ProtocolosOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | % { Write-Host $_ }
  $ProtSel         = Read-Host "Ex.: 1,4"
  $Protocolos      = @()
  if ($ProtSel) {
    $Protocolos = $ProtSel -split ',' | % { $_.Trim() } | ? { $_ -match '^\d+$' } | % { $ProtocolosOptions[[int]$_] } | ? { $_ }
  }

  # Rede
  $IP              = Read-Host "Endereço IP [opcional]"
  $Mascara         = Read-Host "Máscara de rede [opcional]"
  $Gateway         = Read-Host "Gateway [opcional]"
  $HostRede        = Read-Host "Nome/Hostname na rede [opcional]"
  $MAC             = Read-Host "Endereço MAC [opcional]"

  # Fila/driver
  $FilaServidor    = Read-Host "Fila no servidor (ex.: \\SRV-PRINT\HP-M404) [opcional]"
  $DriverVersao    = Read-Host "Versão do driver [opcional]"

  # Contadores/consumíveis
  $PaginasTotais   = Read-Host "Contador total de páginas [opcional]"
  $PaginasPB       = Read-Host "Páginas PB [opcional]"
  $PaginasCor      = Read-Host "Páginas Cor [opcional]"
  $TonerK          = Read-Host "Toner/Preto (%) [opcional]"
  $TonerC          = Read-Host "Toner/Ciano (%) [opcional]"
  $TonerM          = Read-Host "Toner/Magenta (%) [opcional]"
  $TonerY          = Read-Host "Toner/Amarelo (%) [opcional]"
  $DrumK           = Read-Host "Unidade de imagem/Drum Preto (%) [opcional]"
  $CicloMensalMax  = Read-Host "Ciclo mensal máximo (páginas) [opcional]"
  $Bandejas        = Read-Host "Número de bandejas (ex.: 1,2,3) [opcional]"

  # Garantia / manutenção / status
  $GarantiaAteStr  = Read-Host "Garantia até (dd/mm/aaaa) [opcional]"
  $UltManutStr     = Read-Host "Última manutenção (dd/mm/aaaa) [opcional]"
  $ProxManutStr    = Read-Host "Próxima manutenção prevista (dd/mm/aaaa) [opcional]"
  $StatusOper      = Read-Host "Status operacional (OK/Com falhas/Em manutenção) [opcional]"

  # Avaliação (depreciação)
  $UsarAvaliacao   = Read-Host "Deseja calcular valor estimado por depreciação? (S/N)"

  # Parse/normalize
  $Estado       = if ($EstadoStr) { $EstadoStr.Trim() } else { $null }
  $DataCompra   = Parse-Data $DataCompraStr
  $PrecoCompra  = Parse-Preco $PrecoCompraStr
  $GarantiaAte  = Parse-Data $GarantiaAteStr
  $UltManut     = Parse-Data $UltManutStr
  $ProxManut    = Parse-Data $ProxManutStr

  # Pré-visualização
  $dcText = if ($DataCompra) { Fmt-Date $DataCompra } else { "N/D" }
  $pcText = if ($PrecoCompra) { "R$ {0:N2}" -f $PrecoCompra } else { "N/D" }
  $estadoText = if ($Estado) { $Estado } else { "N/D" }

  Write-Host ""
  Write-Host "== Confirme os dados =="
  Write-Host ("Nome informado     : {0}" -f $NomeEquipamento)
  Write-Host ("Patrimônio         : {0}" -f $Patrimonio)
  Write-Host ("Local/Setor        : {0}" -f $Local)
  Write-Host ("Responsável        : {0}" -f $Responsavel)
  Write-Host ("Condição           : {0}" -f $estadoText)
  Write-Host ("Data de compra     : {0}" -f $dcText)
  Write-Host ("Preço de compra    : {0}" -f $pcText)
  Write-Host ("Marca / Modelo     : {0} / {1}" -f $Marca, $Modelo)
  Write-Host ("Tipo               : {0}" -f $TipoImpressora)
  Write-Host ("Cor                : {0}" -f $TipoCor)
  Write-Host ("Duplex             : {0}" -f $Duplex)
  Write-Host ("Papel suportado    : {0}" -f ($(if ($TamanhosPapel.Count) { $TamanhosPapel -join ", " } else { "N/D" })))
  Write-Host ("Conectividade      : {0}" -f ($(if ($Conectividade.Count) { $Conectividade -join ", " } else { "N/D" })))
  Write-Host ("Protocolos         : {0}" -f ($(if ($Protocolos.Count) { $Protocolos -join ", " } else { "N/D" })))
  Write-Host ("IP/Máscara/GW      : {0} / {1} / {2}" -f ($(if ($IP) { $IP } else { "N/D" })), ($(if ($Mascara) { $Mascara } else { "N/D" })), ($(if ($Gateway) { $Gateway } else { "N/D" })))
  Write-Host ("Hostname / MAC     : {0} / {1}" -f ($(if ($HostRede) { $HostRede } else { "N/D" })), ($(if ($MAC) { $MAC } else { "N/D" })))
  Write-Host ("Fila / Driver      : {0} / {1}" -f ($(if ($FilaServidor) { $FilaServidor } else { "N/D" })), ($(if ($DriverVersao) { $DriverVersao } else { "N/D" })))
  Write-Host ("Contadores (Tot/PB/Cor) : {0} / {1} / {2}" -f ($(if ($PaginasTotais) { $PaginasTotais } else { "N/D" })), ($(if ($PaginasPB) { $PaginasPB } else { "N/D" })), ($(if ($PaginasCor) { $PaginasCor } else { "N/D" })))
  Write-Host ("Toner K/C/M/Y (%)  : {0}/{1}/{2}/{3}" -f ($(if ($TonerK) { $TonerK } else { "N/D" })), ($(if ($TonerC) { $TonerC } else { "N/D" })), ($(if ($TonerM) { $TonerM } else { "N/D" })), ($(if ($TonerY) { $TonerY } else { "N/D" })))
  Write-Host ("Drum K (%)         : {0}" -f ($(if ($DrumK) { $DrumK } else { "N/D" })))
  Write-Host ("Ciclo mensal máx.  : {0}" -f ($(if ($CicloMensalMax) { $CicloMensalMax } else { "N/D" })))
  Write-Host ("Bandejas           : {0}" -f ($(if ($Bandejas) { $Bandejas } else { "N/D" })))
  Write-Host ("Garantia até       : {0}" -f (Fmt-Date $GarantiaAte))
  Write-Host ("Últ. manutenção    : {0}" -f (Fmt-Date $UltManut))
  Write-Host ("Próx. manutenção   : {0}" -f (Fmt-Date $ProxManut))
  Write-Host ("Status operacional : {0}" -f ($(if ($StatusOper) { $StatusOper } else { "N/D" })))
  Write-Host ("Observações        : {0}" -f ($(if ($Notas) { $Notas } else { "-" })))

  Write-Host ""
  $confirm = Read-Host "Digite 1 para CONFIRMAR e continuar, ou 2 para REINICIAR"
} while ($confirm -ne '1')
# =====================================================================

# =================== AVALIAÇÃO (opcional) ===================
$ValorEstimado   = $null
$DepreciacaoMeses= $DepreciacaoMesesPadrao
$DataBaseIdade   = $DataCompra
$hoje            = Get-Date
if ($UsarAvaliacao -match '^[sS]') {
  if (-not $DataBaseIdade) { $DataBaseIdade = $UltManut }
  if (-not $DataBaseIdade) { $DataBaseIdade = $hoje }
  $MesesUso = Calc-Meses $DataBaseIdade $hoje
  if ($PrecoCompra -and $MesesUso -ne $null) {
    $deprMensal = [decimal]($PrecoCompra / $DepreciacaoMeses)
    $acumulada  = [decimal]($deprMensal * [Math]::Min($MesesUso,$DepreciacaoMeses))
    $valorResid = [decimal]([Math]::Max(($PrecoCompra * (1 - $PisoResidualPercentual)), $PrecoCompra - $acumulada))
    $pisoAbs    = [decimal]($PrecoCompra * $PisoResidualPercentual)
    if ($valorResid -lt $pisoAbs) { $valorResid = $pisoAbs }
    $ValorEstimado = [decimal]([Math]::Round($valorResid,2))
  }
} else { $MesesUso = $null }
# ============================================================

# =================== SAÍDA (pastas/arquivos) ===================
$stamp         = Get-Date -Format "yyyyMMdd-HHmmss"
$pastaBaseName = ("IMPRESSORA_{0}_{1}_{2}" -f ($Marca -replace '\s',''), ($Modelo -replace '\s',''), $stamp)
$saidaDir      = Join-Path $SaidaRaiz $pastaBaseName
New-Item -ItemType Directory -Path $saidaDir -Force | Out-Null

$hostName = $env:COMPUTERNAME
$txtPath   = Join-Path $saidaDir ("{0}_resumo.txt"    -f $hostName)
$jsonPath  = Join-Path $saidaDir ("{0}_sistema.json"  -f $hostName)
$csvPath   = Join-Path $saidaDir ("{0}_softwares.csv" -f $hostName) # compatibilidade (vazio)

# TXT detalhado
$sb = New-Object System.Text.StringBuilder
$null = $sb.AppendLine("==== Cadastro de Impressora ====")
$null = $sb.AppendLine("Data/Hora: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))
$null = $sb.AppendLine("Nome informado: $NomeEquipamento")
$null = $sb.AppendLine("Patrimônio: $Patrimonio")
$null = $sb.AppendLine("Local/Setor: $Local")
$null = $sb.AppendLine("Responsável: $Responsavel")
$null = $sb.AppendLine("Condição: " + ($(if ($Estado) { $Estado } else { "N/D" })))
$null = $sb.AppendLine("Preço de compra: " + ($(if ($PrecoCompra) { ("R$ {0:N2}" -f $PrecoCompra) } else { "N/D" })))
$null = $sb.AppendLine("Data de compra: " + (Fmt-Date $DataCompra))
$null = $sb.AppendLine("Observações: " + ($(if ($Notas) { $Notas } else { "-" })))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("Marca/Modelo: $Marca / $Modelo")
$null = $sb.AppendLine("Tipo / Cor / Duplex: $TipoImpressora / $TipoCor / $Duplex")
$null = $sb.AppendLine("Papel suportado: " + ($(if ($TamanhosPapel.Count) { $TamanhosPapel -join ", " } else { "N/D" })))
$null = $sb.AppendLine("Conectividade: " + ($(if ($Conectividade.Count) { $Conectividade -join ", " } else { "N/D" })))
$null = $sb.AppendLine("Protocolos: " + ($(if ($Protocolos.Count) { $Protocolos -join ", " } else { "N/D" })))
$null = $sb.AppendLine("IP/Máscara/GW: " + ($(if ($IP) { $IP } else { "N/D" })) + " / " + ($(if ($Mascara) { $Mascara } else { "N/D" })) + " / " + ($(if ($Gateway) { $Gateway } else { "N/D" })))
$null = $sb.AppendLine("Hostname / MAC: " + ($(if ($HostRede) { $HostRede } else { "N/D" })) + " / " + ($(if ($MAC) { $MAC } else { "N/D" })))
$null = $sb.AppendLine("Fila no servidor: " + ($(if ($FilaServidor) { $FilaServidor } else { "N/D" })))
$null = $sb.AppendLine("Versão do driver: " + ($(if ($DriverVersao) { $DriverVersao } else { "N/D" })))
$null = $sb.AppendLine("Contadores (Tot/PB/Cor): " + ($(if ($PaginasTotais) { $PaginasTotais } else { "N/D" })) + " / " + ($(if ($PaginasPB) { $PaginasPB } else { "N/D" })) + " / " + ($(if ($PaginasCor) { $PaginasCor } else { "N/D" })))
$null = $sb.AppendLine("Toner K/C/M/Y (%): " + ($(if ($TonerK) { $TonerK } else { "N/D" })) + "/" + ($(if ($TonerC) { $TonerC } else { "N/D" })) + "/" + ($(if ($TonerM) { $TonerM } else { "N/D" })) + "/" + ($(if ($TonerY) { $TonerY } else { "N/D" })))
$null = $sb.AppendLine("Drum K (%): " + ($(if ($DrumK) { $DrumK } else { "N/D" })))
$null = $sb.AppendLine("Ciclo mensal máximo: " + ($(if ($CicloMensalMax) { $CicloMensalMax } else { "N/D" })))
$null = $sb.AppendLine("Bandejas: " + ($(if ($Bandejas) { $Bandejas } else { "N/D" })))
$null = $sb.AppendLine("Garantia até: " + (Fmt-Date $GarantiaAte))
$null = $sb.AppendLine("Última manutenção: " + (Fmt-Date $UltManut))
$null = $sb.AppendLine("Próxima manutenção: " + (Fmt-Date $ProxManut))
$null = $sb.AppendLine("Status operacional: " + ($(if ($StatusOper) { $StatusOper } else { "N/D" })))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Avaliação (estimativa) ==")
$null = $sb.AppendLine("  Base para idade: " + ($(if ($DataBaseIdade) { (Fmt-Date $DataBaseIdade) } else { "N/D" })))
$null = $sb.AppendLine("  Meses de uso: " + ($(if ($MesesUso -ne $null) { $MesesUso } else { "N/D" })))
$null = $sb.AppendLine("  Depreciação (meses): $DepreciacaoMeses")
$null = $sb.AppendLine("  Piso residual: $([int]($PisoResidualPercentual*100))%")
$null = $sb.AppendLine("  Valor estimado: " + ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D" })))

$sb.ToString() | Out-File -FilePath $txtPath -Encoding UTF8

# CSV (compatibilidade; vazio)
@() | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

# JSON estruturado (compatível com padrão)
$payload = [pscustomobject]@{
  Timestamp        = Get-Date
  TipoEquipamento  = "Impressora"
  NomeInformado    = $NomeEquipamento
  Patrimonio       = $Patrimonio
  Local            = $Local
  Responsavel      = $Responsavel
  Estado           = $Estado
  DataCompra       = $DataCompra
  PrecoCompra      = $PrecoCompra
  Observacoes      = $Notas

  Hostname         = $env:COMPUTERNAME
  Usuario          = $null
  Fabricante       = $Marca
  Modelo           = $Modelo
  SerialBIOS       = $NumeroSerie
  CPU              = $null
  Cores            = $null
  Threads          = $null
  RAM_GB           = $null
  SO               = $null
  Arquitetura      = $null
  Build            = $null
  OSInstall        = $null
  BIOSDate         = $null
  LastBoot         = $null
  UptimeMin        = $null
  Dominio          = $null
  Workgroup        = $null
  GPUs             = @()
  Discos           = @()
  Redes            = @(
    [pscustomobject]@{
      IP       = $IP
      Mascara  = $Mascara
      Gateway  = $Gateway
      Hostname = $HostRede
      MAC      = $MAC
    }
  )

  TipoImpressora    = $TipoImpressora
  TipoCor           = $TipoCor
  Duplex            = $Duplex
  TamanhosPapel     = $TamanhosPapel
  Conectividade     = $Conectividade
  Protocolos        = $Protocolos
  FilaServidor      = $FilaServidor
  DriverVersao      = $DriverVersao

  PaginasTotais     = $PaginasTotais
  PaginasPB         = $PaginasPB
  PaginasCor        = $PaginasCor
  TonerK            = $TonerK
  TonerC            = $TonerC
  TonerM            = $TonerM
  TonerY            = $TonerY
  DrumK             = $DrumK
  CicloMensalMax    = $CicloMensalMax
  Bandejas          = $Bandejas

  GarantiaAte       = $GarantiaAte
  UltimaManutencao  = $UltManut
  ProxManutencao    = $ProxManut
  StatusOperacional = $StatusOper

  BaseIdade         = $DataBaseIdade
  MesesUso          = $MesesUso
  DepMesesCfg       = $DepreciacaoMeses
  PisoResidual      = $PisoResidualPercentual
  ValorEstimado     = $ValorEstimado
}
$payload | Add-Member -NotePropertyName DataCompraText  -NotePropertyValue (Fmt-Date $DataCompra)
$payload | Add-Member -NotePropertyName GarantiaAteText -NotePropertyValue (Fmt-Date $GarantiaAte)
$payload | Add-Member -NotePropertyName UltManutText    -NotePropertyValue (Fmt-Date $UltManut)
$payload | Add-Member -NotePropertyName ProxManutText   -NotePropertyValue (Fmt-Date $ProxManut)

$payload | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

# ========= TXT adicional (resumo curto) =========
$rotuloPath = Join-Path $saidaDir ("{0}info.txt" -f ($NomeEquipamento -replace '[\\/:*?""<>|]','_'))
$rotuloSb = New-Object System.Text.StringBuilder
$null = $rotuloSb.AppendLine("Resumo do equipamento (Impressora)")
$null = $rotuloSb.AppendLine("Gerado em: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))
$null = $rotuloSb.AppendLine("")
$null = $rotuloSb.AppendLine("Nome (operador): $NomeEquipamento")
$null = $rotuloSb.AppendLine("Marca/Modelo: $Marca / $Modelo")
$null = $rotuloSb.AppendLine("Tipo/Cor/Duplex: $TipoImpressora / $TipoCor / $Duplex")
$null = $rotuloSb.AppendLine("Papel: " + ($(if ($TamanhosPapel.Count) { $TamanhosPapel -join ", " } else { "N/D" })))
$null = $rotuloSb.AppendLine("Conexão: " + ($(if ($Conectividade.Count) { $Conectividade -join ", " } else { "N/D" })))
$null = $rotuloSb.AppendLine("IP/Host: " + ($(if ($IP) { $IP } else { "N/D" })) + " / " + ($(if ($HostRede) { $HostRede } else { "N/D" })))
$null = $rotuloSb.AppendLine("Local: $Local")
$null = $rotuloSb.AppendLine("Responsável: $Responsavel")
$null = $rotuloSb.AppendLine("Valor estimado: " + ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D" })))
$rotuloSb.ToString() | Out-File -FilePath $rotuloPath -Encoding UTF8

# ===== TXT “nome-modelo-data” (rótulo rápido) =====
$rotuloRapido = Join-Path $saidaDir ("IMPRESSORA_{0}_{1}_{2}.txt" -f ($Marca -replace '\s',''), ($Modelo -replace '\s',''), (Get-Date -Format "yyyyMMdd"))
"IMPRESSORA: $Marca $Modelo`r`nLocal: $Local`r`nPatrimônio: $Patrimonio`r`nIP: $IP" | Out-File -FilePath $rotuloRapido -Encoding UTF8
# =====================================================================

# =================== CÓPIA PARA COMPARTILHAMENTO =====================
$DesejaCopiar = if ($CopiarParaCompartilhamento) { Read-Host "Deseja copiar para $DestinoCompartilhamento ? (S/N)" } else { "N" }
if ($CopiarParaCompartilhamento -and ($DesejaCopiar -match '^[sS]')) {
  $UsuarioCompart = Read-Host "Usuário do compartilhamento (ex.: DOMINIO\usuario ou servidor\usuario)"
  $SenhaCompart   = Read-Host "Senha do compartilhamento (digitação oculta)" -AsSecureString
  try {
    if (-not $UsuarioCompart -or -not $SenhaCompart) {
      Write-Warning "Cópia habilitada, mas usuário/senha não informados. Pulando cópia."
    } else {
      $cred = New-Object System.Management.Automation.PSCredential($UsuarioCompart, $SenhaCompart)
      Write-Host "Tentando copiar a pasta para: $DestinoCompartilhamento"
      Copy-WithCredentialsAndCleanup -SourcePath $saidaDir -SharePath $DestinoCompartilhamento -Credential $cred
      Write-Host "Cópia realizada para: $DestinoCompartilhamento (sessão e cache de credenciais limpos)"
    }
  } catch {
    Write-Warning ("Falha ao copiar para {0}: {1}" -f $DestinoCompartilhamento, $_.Exception.Message)
  }
}

Write-Host "Cadastro de impressora concluído. Arquivos em: $saidaDir"
