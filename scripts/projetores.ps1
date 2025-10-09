<#
Cadastro de Projetores (PowerShell) - Repositório de Informações - v2 PT-BR
- NÃO coleta dados do computador. Apenas perguntas ao operador.
- Integração: tenta carregar "catalogo-projetores.ps1" (se existir) para pré-preencher dados.
- Coleta: Identificação, patrimônio, local, responsável, estado, compra, preço, notas.
- Específicos: marca, modelo, tipo (menu ou catálogo), resolução, brilho, contraste,
  fonte de luz, horas de lâmpada/vida, entradas (menu), conectividade (menu),
  voltagem, instalação, acessórios, n° de série, garantia, manutenção (últ/próx), status.
- Avaliação por depreciação linear (opcional).
- Saída: TXT detalhado, JSON (compatível com seus padrões), TXT curto, TXT “PROJETOR_MARCA_MODELO_DATA”.
- Opcional: cópia para compartilhamento SMB com usuário/senha e limpeza de sessão/cache.

Compatibilidade: PowerShell 5+ (Windows 10/11)
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


# ============ IMPORTA CATÁLOGO (se existir) ============
$CatalogoLoaded = $false
$CatalogoProjetores = $null

try {
  $CatalogoPath = Join-Path $PSScriptRoot 'catalogo-projetores.ps1'
  if (-not (Test-Path $CatalogoPath)) {
    $CatalogoPathAlt = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\catalogo-projetores.ps1'
    if (Test-Path $CatalogoPathAlt) { $CatalogoPath = $CatalogoPathAlt }
  }
  if (Test-Path $CatalogoPath) {
    . $CatalogoPath
    if ($CatalogoProjetores) {
      $CatalogoLoaded = $true
      Write-Host "Catálogo de projetores carregado: $CatalogoPath" -ForegroundColor DarkCyan
    }
  } else {
    Write-Host "Catálogo 'catalogo-projetores.ps1' não encontrado (usando catálogo interno)..." -ForegroundColor DarkYellow
  }
} catch {
  Write-Warning "Falha ao carregar catálogo: $($_.Exception.Message)"
}

# Catálogo interno (fallback) com os modelos solicitados
if (-not $CatalogoLoaded) {
  $CatalogoProjetores = @(
    # EPSON (LCD / 3LCD)
    [pscustomobject]@{ Marca="Epson"; Modelo="S41+"; Descricao="PROJETOR XGA X41+ BR 3600 EPSON"; TipoProjetor="LCD"; ResolucaoSug=$null; BrilhoSug=3600; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="Epson"; Modelo="S8+";  Descricao="Epson S8+";                            TipoProjetor="LCD"; ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="Epson"; Modelo="S12+"; Descricao="Epson S12+";                           TipoProjetor="LCD"; ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="Epson"; Modelo="S18+"; Descricao="Epson S18+";                           TipoProjetor="LCD"; ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="Epson"; Modelo="S31+"; Descricao="Epson S31+";                           TipoProjetor="LCD"; ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada" }

    # BENQ (DLP)
    [pscustomobject]@{ Marca="BenQ";  Modelo="MS550"; Descricao="Projetor BenQ MS550 SVGA 3600 Ansi";   TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=3600; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="BenQ";  Modelo="MX611"; Descricao="PROJETOR XGA MX611 BR 4000 BENQ";     TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=4000; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="BenQ";  Modelo="MS531"; Descricao="PROJETOR SVGA MS531 BR 3300 BENQ";    TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=3300; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
    [pscustomobject]@{ Marca="BenQ";  Modelo="MX560"; Descricao="PROJETOR XGA MX560 IM 4000";          TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=4000; ContrasteSug=$null; FonteDeLuz="Lâmpada" }
  )
}

function Get-CatalogoProjetores {
  $i = 0
  $CatalogoProjetores | ForEach-Object {
    $i++
    [pscustomobject]@{
      Id           = $i
      Marca        = $_.Marca
      Modelo       = $_.Modelo
      Descricao    = $_.Descricao
      TipoProjetor = $_.TipoProjetor
      ResolucaoSug = $_.ResolucaoSug
      BrilhoSug    = $_.BrilhoSug
      ContrasteSug = $_.ContrasteSug
      FonteDeLuz   = $_.FonteDeLuz
    }
  }
}

function Get-ModeloDoCatalogo {
  param(
    [Parameter(Mandatory)][string]$Marca,
    [Parameter(Mandatory)][string]$Modelo
  )
  $m = $CatalogoProjetores |
       Where-Object { $_.Marca -ieq $Marca -and $_.Modelo -ieq $Modelo } |
       Select-Object -First 1
  if (-not $m) { return $null }
  [pscustomobject]@{
    Marca        = $m.Marca
    Modelo       = $m.Modelo
    TipoProjetor = $m.TipoProjetor
    Resolucao    = $m.ResolucaoSug
    BrilhoLumens = $m.BrilhoSug
    Contraste    = $m.ContrasteSug
    FonteDeLuz   = $m.FonteDeLuz
  }
}

function Select-ModeloDoCatalogo {
  $lista = Get-CatalogoProjetores | Sort-Object Marca, Modelo
  if (-not $lista) { return $null }

  Write-Host ""
  Write-Host "== Catálogo de Projetores ==" -ForegroundColor Cyan
  $lista | Format-Table Id,Marca,Modelo,Descricao,TipoProjetor,BrilhoSug -AutoSize
  Write-Host ""
  $sel = Read-Host "Digite o Id do modelo (ou 0 para 'Outro modelo')"

  if ($sel -match '^\s*0\s*$') { return $null }
  if ($sel -notmatch '^\d+$') { return $null }

  $escolha = $lista | Where-Object { $_.Id -eq [int]$sel } | Select-Object -First 1
  if (-not $escolha) { return $null }

  return (Get-ModeloDoCatalogo -Marca $escolha.Marca -Modelo $escolha.Modelo)
}
# ========================================================


# =================== FUNÇÕES ===================
function Try-Get { param([scriptblock]$Block) try { & $Block } catch { $null } }
function Fmt-Date($d) { if ($d) { $d.ToString("dd/MM/yyyy HH:mm:ss") } else { "N/D" } }
function Nz([object]$v, [string]$default = "N/D") { if ($null -ne $v -and "$v" -ne "") { $v } else { $default } }

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
  elseif ($clean -match ',\d{1,2}$')                   { $clean = $clean -replace ',','.' }
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

function Get-UncHost { param([Parameter(Mandatory)][string]$UncPath) if ($UncPath -match '^\\\\([^\\]+)') { return $matches[1] } else { return $null } }

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
$TiposProjetor = [ordered]@{
  1 = "LCD"
  2 = "DLP"
  3 = "LCoS"
  4 = "Laser"
  5 = "LED"
  6 = "Híbrido (Laser/LED)"
  7 = "Ultra Curta Distância"
  8 = "Interativo"
  9 = "Outros"
}

$EntradasOptions = [ordered]@{
  1 = "HDMI"
  2 = "VGA"
  3 = "DisplayPort"
  4 = "USB-C (Alt Mode)"
  5 = "USB A (mídia)"
  6 = "AV/Composite"
  7 = "Audio In/Out"
}

$ConectividadeOptions = [ordered]@{
  1 = "Wi-Fi"
  2 = "LAN/RJ-45"
  3 = "Miracast"
  4 = "Bluetooth"
  5 = "Sem fio proprietário"
}
# ==============================================


# =================== PROMPTS (PT-BR + confirmação) ===================
do {
  Clear-Host
  Write-Host "== Cadastro de Projetor =="

  # Identificação geral
  $NomeComputadorInformado = Read-Host "Nome do equipamento (etiqueta/operador)"
  $Patrimonio     = Read-Host "Número do patrimônio (ex.: 2025-00123)"
  $Local          = Read-Host "Local/Setor (ex.: Biblioteca / Sala 3)"
  $Responsavel    = Read-Host "Responsável (ex.: Maria Silva)"
  $EstadoStr      = Read-Host "Condição geral (Novo/Bom/Regular/Ruim) [opcional]"
  $DataCompraStr  = Read-Host "Data de compra (dd/mm/aaaa) [opcional]"
  $PrecoCompraStr = Read-Host "Preço de compra (ex.: 3499,90) [opcional]"
  $Notas          = Read-Host "Observações [opcional]"

  # ======== PRÉ-PREENCHIMENTO PELO CATÁLOGO (opcional) =========
  $Marca = $null; $Modelo = $null; $TipoProjetor = $null; $Resolucao = $null; $BrilhoLumens = $null; $Contraste = $null; $FonteDeLuz = $null
  if ($CatalogoProjetores -and $CatalogoProjetores.Count -gt 0) {
    $usarCat = Read-Host "Deseja selecionar um modelo do catálogo? (S/N)"
    if ($usarCat -match '^[sS]') {
      $preset = Select-ModeloDoCatalogo
      if ($preset) {
        $Marca        = $preset.Marca
        $Modelo       = $preset.Modelo
        if ($preset.TipoProjetor) { $TipoProjetor = $preset.TipoProjetor }
        if ($preset.Resolucao)    { $Resolucao    = $preset.Resolucao }
        if ($preset.BrilhoLumens) { $BrilhoLumens = $preset.BrilhoLumens }
        if ($preset.Contraste)    { $Contraste    = $preset.Contraste }
        if ($preset.FonteDeLuz)   { $FonteDeLuz   = $preset.FonteDeLuz }
        Write-Host ("Pré-preenchido: {0} / {1} {2}" -f $Marca, $Modelo, ($(if ($TipoProjetor) { "[$TipoProjetor]" } else { "" }))) -ForegroundColor Green
      } else {
        Write-Host "Outro modelo (preenchimento manual)." -ForegroundColor Yellow
      }
    }
  }

  # Dados específicos (pergunta somente o que ainda falta)
  if (-not $Marca)        { $Marca        = Read-Host "Marca do projetor (ex.: Epson, BenQ, Optoma)" }
  if (-not $Modelo)       { $Modelo       = Read-Host "Modelo do projetor (ex.: EB-X05)" }
  $NumeroSerie    = Read-Host "Número de série [opcional]"

  if (-not $TipoProjetor) {
    Write-Host ""
    Write-Host "Tipos de projetor:"
    $TiposProjetor.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | ForEach-Object { Write-Host $_ }
    $TipoSel        = Read-Host "Escolha o tipo (número, ex.: 1=LCD, 2=DLP, 4=Laser)"
    $TipoProjetor   = $TiposProjetor[[int]$TipoSel]
  }

  if (-not $Resolucao)    { $Resolucao    = Read-Host "Resolução nativa (ex.: 1024x768, 1280x800, 1920x1080)" }
  if (-not $BrilhoLumens) { $BrilhoLumens = Read-Host "Brilho (ANSI lumens) [ex.: 3600]" }
  if (-not $Contraste)    { $Contraste    = Read-Host "Contraste (ex.: 15000:1) [opcional]" }
  if (-not $FonteDeLuz)   { $FonteDeLuz   = Read-Host "Fonte de luz (Lâmpada/Laser/LED/Outro)" }

  $HorasLampada   = Read-Host "Horas de lâmpada/fonte já utilizadas [opcional]"
  $VidaLampada    = Read-Host "Vida útil estimada da lâmpada/fonte (horas) [opcional]"

  Write-Host ""
  Write-Host "Entradas disponíveis (selecione múltiplas, separados por vírgula):"
  $EntradasOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | ForEach-Object { Write-Host $_ }
  $EntradasSel    = Read-Host "Ex.: 1,2,4"
  $Entradas       = @()
  if ($EntradasSel) {
    $Entradas = $EntradasSel -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { $EntradasOptions[[int]$_] } | Where-Object { $_ }
  }

  Write-Host ""
  Write-Host "Conectividade (selecione múltiplas, separados por vírgula):"
  $ConectividadeOptions.GetEnumerator() | Sort-Object Key | ForEach-Object { "{0,2}) {1}" -f $_.Key, $_.Value } | ForEach-Object { Write-Host $_ }
  $ConSel         = Read-Host "Ex.: 1,2"
  $Conectividade  = @()
  if ($ConSel) {
    $Conectividade = $ConSel -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { $ConectividadeOptions[[int]$_] } | Where-Object { $_ }
  }

  $Voltagem       = Read-Host "Voltagem (ex.: 110V, 220V, bivolt) [opcional]"
  $Instalacao     = Read-Host "Instalação (teto/mesa/suporte móvel) [opcional]"
  $AcessoriosStr  = Read-Host "Acessórios (separe por vírgula) [opcional]"
  $Acessorios     = @()
  if ($AcessoriosStr) { $Acessorios = $AcessoriosStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } }

  $GarantiaAteStr = Read-Host "Garantia até (dd/mm/aaaa) [opcional]"
  $UltManutStr    = Read-Host "Última manutenção (dd/mm/aaaa) [opcional]"
  $ProxManutStr   = Read-Host "Próxima manutenção prevista (dd/mm/aaaa) [opcional]"
  $StatusOper     = Read-Host "Status operacional (OK/Com falhas/Em manutenção) [opcional]"

  # Avaliação (depreciação)
  $UsarAvaliacao  = Read-Host "Deseja calcular valor estimado por depreciação? (S/N)"
  $Estado      = if ($EstadoStr) { $EstadoStr.Trim() } else { $null }
  $DataCompra  = Parse-Data $DataCompraStr
  $PrecoCompra = Parse-Preco $PrecoCompraStr
  $GarantiaAte = Parse-Data $GarantiaAteStr
  $UltManut    = Parse-Data $UltManutStr
  $ProxManut   = Parse-Data $ProxManutStr

  # Pré-visualização
  $dcText = Nz (Fmt-Date $DataCompra)
  $pcText = if ($PrecoCompra) { "R$ {0:N2}" -f $PrecoCompra } else { "N/D" }
  $estadoText = Nz $Estado

  Write-Host ""
  Write-Host "== Confirme os dados =="
  Write-Host ("Nome informado     : {0}" -f $NomeComputadorInformado)
  Write-Host ("Patrimônio         : {0}" -f $Patrimonio)
  Write-Host ("Local/Setor        : {0}" -f $Local)
  Write-Host ("Responsável        : {0}" -f $Responsavel)
  Write-Host ("Condição           : {0}" -f $estadoText)
  Write-Host ("Data de compra     : {0}" -f $dcText)
  Write-Host ("Preço de compra    : {0}" -f $pcText)
  Write-Host ("Marca / Modelo     : {0} / {1}" -f $Marca, $Modelo)
  Write-Host ("Tipo de projetor   : {0}" -f $TipoProjetor)
  Write-Host ("Resolução          : {0}" -f $Resolucao)
  Write-Host ("Brilho (lumens)    : {0}" -f $BrilhoLumens)
  Write-Host ("Contraste          : {0}" -f (Nz $Contraste))
  Write-Host ("Fonte de luz       : {0}" -f $FonteDeLuz)
  Write-Host ("Horas / Vida (h)   : {0} / {1}" -f (Nz $HorasLampada), (Nz $VidaLampada))
  Write-Host ("Entradas           : {0}" -f ($(if ($Entradas.Count) { $Entradas -join ", " } else { "N/D" })))
  Write-Host ("Conectividade      : {0}" -f ($(if ($Conectividade.Count) { $Conectividade -join ", " } else { "N/D" })))
  Write-Host ("Voltagem           : {0}" -f (Nz $Voltagem))
  Write-Host ("Instalação         : {0}" -f (Nz $Instalacao))
  Write-Host ("Acessórios         : {0}" -f ($(if ($Acessorios.Count) { $Acessorios -join ", " } else { "N/D" })))
  Write-Host ("Nº de série        : {0}" -f (Nz $NumeroSerie))
  Write-Host ("Garantia até       : {0}" -f (Fmt-Date $GarantiaAte))
  Write-Host ("Últ. manutenção    : {0}" -f (Fmt-Date $UltManut))
  Write-Host ("Próx. manutenção   : {0}" -f (Fmt-Date $ProxManut))
  Write-Host ("Status operacional : {0}" -f (Nz $StatusOper))
  Write-Host ("Observações        : {0}" -f (Nz $Notas, "-"))

  Write-Host ""
  $confirm = Read-Host "Digite 1 para CONFIRMAR e continuar, ou 2 para REINICIAR"
} while ($confirm -ne '1')
# =====================================================================


# =================== AVALIAÇÃO (opcional) ===================
$ValorEstimado = $null
$DepreciacaoMeses = $DepreciacaoMesesPadrao
$DataBaseIdade = $DataCompra
$hoje = Get-Date
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
} else {
  $MesesUso = $null
}
# ============================================================


# =================== SAÍDA (pastas/arquivos) ===================
$stamp        = Get-Date -Format "yyyyMMdd-HHmmss"
$pastaBaseName = ("PROJETOR_{0}_{1}_{2}" -f ($Marca -replace '\s',''), ($Modelo -replace '\s',''), $stamp)
$saidaDir     = Join-Path $SaidaRaiz $pastaBaseName
New-Item -ItemType Directory -Path $saidaDir -Force | Out-Null

$hostName = $env:COMPUTERNAME
$txtPath   = Join-Path $saidaDir ("{0}_resumo.txt"    -f $hostName)
$jsonPath  = Join-Path $saidaDir ("{0}_sistema.json"  -f $hostName)
$csvPath   = Join-Path $saidaDir ("{0}_softwares.csv" -f $hostName) # compatibilidade (vazio)

# TXT detalhado
$sb = New-Object System.Text.StringBuilder
$null = $sb.AppendLine("==== Cadastro de Projetor ====")
$null = $sb.AppendLine("Data/Hora: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))
$null = $sb.AppendLine(("Nome informado: {0}" -f $NomeComputadorInformado))
$null = $sb.AppendLine(("Patrimônio: {0}" -f $Patrimonio))
$null = $sb.AppendLine(("Local/Setor: {0}" -f $Local))
$null = $sb.AppendLine(("Responsável: {0}" -f $Responsavel))
$null = $sb.AppendLine(("Condição: {0}" -f (Nz $Estado)))
$null = $sb.AppendLine(("Preço de compra: {0}" -f ($(if ($PrecoCompra) { "R$ {0:N2}" -f $PrecoCompra } else { "N/D" }))))
$null = $sb.AppendLine(("Data de compra: {0}" -f (Fmt-Date $DataCompra)))
$null = $sb.AppendLine(("Observações: {0}" -f (Nz $Notas, "-")))
$null = $sb.AppendLine("")
$null = $sb.AppendLine(("Marca/Modelo: {0} / {1}" -f $Marca, $Modelo))
$null = $sb.AppendLine(("Tipo de projetor: {0}" -f $TipoProjetor))
$null = $sb.AppendLine(("Resolução nativa: {0}" -f $Resolucao))
$null = $sb.AppendLine(("Brilho (ANSI lumens): {0}" -f $BrilhoLumens))
$null = $sb.AppendLine(("Contraste: {0}" -f (Nz $Contraste)))
$null = $sb.AppendLine(("Fonte de luz: {0}" -f $FonteDeLuz))
$null = $sb.AppendLine(("Horas usadas / Vida (h): {0} / {1}" -f (Nz $HorasLampada), (Nz $VidaLampada)))
$null = $sb.AppendLine(("Entradas: {0}" -f ($(if ($Entradas.Count) { $Entradas -join ", " } else { "N/D" }))))
$null = $sb.AppendLine(("Conectividade: {0}" -f ($(if ($Conectividade.Count) { $Conectividade -join ", " } else { "N/D" }))))
$null = $sb.AppendLine(("Voltagem: {0}" -f (Nz $Voltagem)))
$null = $sb.AppendLine(("Instalação: {0}" -f (Nz $Instalacao)))
$null = $sb.AppendLine(("Acessórios: {0}" -f ($(if ($Acessorios.Count) { $Acessorios -join ", " } else { "N/D" }))))
$null = $sb.AppendLine(("Número de série: {0}" -f (Nz $NumeroSerie)))
$null = $sb.AppendLine(("Garantia até: {0}" -f (Fmt-Date $GarantiaAte)))
$null = $sb.AppendLine(("Última manutenção: {0}" -f (Fmt-Date $UltManut)))
$null = $sb.AppendLine(("Próxima manutenção: {0}" -f (Fmt-Date $ProxManut)))
$null = $sb.AppendLine(("Status operacional: {0}" -f (Nz $StatusOper)))
$null = $sb.AppendLine("")
$null = $sb.AppendLine("== Avaliação (estimativa) ==")
$null = $sb.AppendLine(("  Base para idade: {0}" -f (Nz (Fmt-Date $DataBaseIdade))))
$null = $sb.AppendLine(("  Meses de uso: {0}" -f ($(if ($MesesUso -ne $null) { $MesesUso } else { "N/D" }))))
$null = $sb.AppendLine(("  Depreciação (meses): {0}" -f $DepreciacaoMeses))
$null = $sb.AppendLine(("  Piso residual: {0}%" -f ([int]($PisoResidualPercentual*100))))
$null = $sb.AppendLine(("  Valor estimado: {0}" -f ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D" }))))

$sb.ToString() | Out-File -FilePath $txtPath -Encoding UTF8

# CSV (compatibilidade; vazio)
@() | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

# JSON estruturado
$payload = [pscustomobject]@{
  Timestamp        = Get-Date
  TipoEquipamento  = "Projetor"
  NomeInformado    = $NomeComputadorInformado
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
  Redes            = @()

  TipoProjetor     = $TipoProjetor
  Resolucao        = $Resolucao
  BrilhoLumens     = $BrilhoLumens
  Contraste        = $Contraste
  FonteDeLuz       = $FonteDeLuz
  HorasUsadas      = $HorasLampada
  VidaEstimLampada = $VidaLampada
  Entradas         = $Entradas
  Conectividade    = $Conectividade
  Voltagem         = $Voltagem
  Instalacao       = $Instalacao
  Acessorios       = $Acessorios
  NumeroSerie      = $NumeroSerie
  GarantiaAte      = $GarantiaAte
  UltimaManutencao = $UltManut
  ProxManutencao   = $ProxManut
  StatusOperacional= $StatusOper

  BaseIdade        = $DataBaseIdade
  MesesUso         = $MesesUso
  DepMesesCfg      = $DepreciacaoMeses
  PisoResidual     = $PisoResidualPercentual
  ValorEstimado    = $ValorEstimado
}
$payload | Add-Member -NotePropertyName DataCompraText  -NotePropertyValue (Fmt-Date $DataCompra)
$payload | Add-Member -NotePropertyName GarantiaAteText -NotePropertyValue (Fmt-Date $GarantiaAte)
$payload | Add-Member -NotePropertyName UltManutText    -NotePropertyValue (Fmt-Date $UltManut)
$payload | Add-Member -NotePropertyName ProxManutText   -NotePropertyValue (Fmt-Date $ProxManut)

$payload | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

# ========= TXT adicional (resumo curto) =========
$rotuloPath = Join-Path $saidaDir ("{0}info.txt" -f ($NomeComputadorInformado -replace '[\\/:*?""<>|]','_'))
$rotuloSb = New-Object System.Text.StringBuilder
$null = $rotuloSb.AppendLine("Resumo do equipamento (Projetor)")
$null = $rotuloSb.AppendLine("Gerado em: " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))
$null = $rotuloSb.AppendLine("")
$null = $rotuloSb.AppendLine(("Nome (operador): {0}" -f $NomeComputadorInformado))
$null = $rotuloSb.AppendLine(("Marca/Modelo: {0} / {1}" -f $Marca, $Modelo))
$null = $rotuloSb.AppendLine(("Tipo: {0}" -f $TipoProjetor))
$null = $rotuloSb.AppendLine(("Resolução: {0}" -f $Resolucao))
$null = $rotuloSb.AppendLine(("Brilho: {0} lumens" -f $BrilhoLumens))
$null = $rotuloSb.AppendLine(("Entradas: {0}" -f ($(if ($Entradas.Count) { $Entradas -join ", " } else { "N/D" }))))
$null = $rotuloSb.AppendLine(("Conectividade: {0}" -f ($(if ($Conectividade.Count) { $Conectividade -join ", " } else { "N/D" }))))
$null = $rotuloSb.AppendLine(("Local: {0}" -f $Local))
$null = $rotuloSb.AppendLine(("Responsável: {0}" -f $Responsavel))
$null = $rotuloSb.AppendLine(("Valor estimado: {0}" -f ($(if ($ValorEstimado -ne $null) { "R$ {0:N2}" -f $ValorEstimado } else { "N/D" }))))
$rotuloSb.ToString() | Out-File -FilePath $rotuloPath -Encoding UTF8

# ===== TXT “nome-modelo-data” (rótulo rápido) =====
$rotuloRapido = Join-Path $saidaDir ("PROJETOR_{0}_{1}_{2}.txt" -f ($Marca -replace '\s',''), ($Modelo -replace '\s',''), (Get-Date -Format "yyyyMMdd"))
("PROJETOR: {0} {1}`r`nLocal: {2}`r`nPatrimônio: {3}" -f $Marca, $Modelo, $Local, $Patrimonio) | Out-File -FilePath $rotuloRapido -Encoding UTF8
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

Write-Host "Cadastro concluído. Arquivos em: $saidaDir"
