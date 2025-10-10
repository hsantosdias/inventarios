# catalogo-projetores.ps1
# Catálogo de modelos de projetores usados para pré-preenchimento no cadastro
# Compatível com PowerShell 5+

# ========= DADOS DO CATÁLOGO =========
# Observações:
# - Qtde = 1 para todos os modelos (ajuste livre, se quiser consolidar por modelo).
# - Campos *_Sug são sugestões para pré-preencher o formulário.

$CatalogoProjetores = @(
  # ---------- EPSON (LCD / 3LCD) ----------
  [pscustomobject]@{ Marca="Epson"; Modelo="S41+"; Descricao="PowerLite X41+ | XGA 1024x768 | 3600 lm | até 15.000:1"; TipoProjetor="LCD"; ResolucaoSug="1024x768"; BrilhoSug=3600; ContrasteSug="15000:1"; FonteDeLuz="Lâmpada"; VidaLampSug="6000/10000"; DataCompraSug="26/07/2021"; PrecoCompraSug="2996,53"; Qtde=1 }
  [pscustomobject]@{ Marca="Epson"; Modelo="S8+";  Descricao="PowerLite S8+ | SVGA 800x600 | 2500 lm | 2000:1";        TipoProjetor="LCD"; ResolucaoSug="800x600";  BrilhoSug=2500; ContrasteSug="2000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="4000/5000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
  [pscustomobject]@{ Marca="Epson"; Modelo="S12+"; Descricao="PowerLite S12+ | SVGA 800x600 | 2800 lm";                 TipoProjetor="LCD"; ResolucaoSug="800x600";  BrilhoSug=2800; ContrasteSug=$null;      FonteDeLuz="Lâmpada"; VidaLampSug="4000/5000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
  [pscustomobject]@{ Marca="Epson"; Modelo="S18+"; Descricao="PowerLite S18+ | SVGA 800x600 | 3000 lm | até 10.000:1";  TipoProjetor="LCD"; ResolucaoSug="800x600";  BrilhoSug=3000; ContrasteSug="10000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="5000/6000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
  [pscustomobject]@{ Marca="Epson"; Modelo="S31+"; Descricao="PowerLite S31+ | SVGA 800x600 | 3200 lm | até 15.000:1"; TipoProjetor="LCD"; ResolucaoSug="800x600";  BrilhoSug=3200; ContrasteSug="15000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="5000/10000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }

  # ---------- BENQ (DLP) ----------
  [pscustomobject]@{ Marca="BenQ";  Modelo="MS550"; Descricao="MS550 | SVGA 800x600 | 3600 lm | 20000:1";               TipoProjetor="DLP"; ResolucaoSug="800x600";  BrilhoSug=3600; ContrasteSug="20000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="5000/10000/15000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
  [pscustomobject]@{ Marca="BenQ";  Modelo="MX611"; Descricao="MX611 | XGA 1024x768 | 4000 lm | 20000:1";               TipoProjetor="DLP"; ResolucaoSug="1024x768"; BrilhoSug=4000; ContrasteSug="20000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="4000/8000/10000/15000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
  [pscustomobject]@{ Marca="BenQ";  Modelo="MS531"; Descricao="MS531 | SVGA 800x600 | 3300 lm | 15000:1";               TipoProjetor="DLP"; ResolucaoSug="800x600";  BrilhoSug=3300; ContrasteSug="15000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="4500/6000/10000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
  [pscustomobject]@{ Marca="BenQ";  Modelo="MX560"; Descricao="MX560 | XGA 1024x768 | 4000 lm | 20000:1";               TipoProjetor="DLP"; ResolucaoSug="1024x768"; BrilhoSug=4000; ContrasteSug="20000:1";  FonteDeLuz="Lâmpada"; VidaLampSug="até 15000"; DataCompraSug=$null; PrecoCompraSug=$null; Qtde=1 }
)

# ========= FUNÇÕES =========

function Get-CatalogoProjetores {
  <#
    .SYNOPSIS
      Retorna o catálogo com coluna Id para facilitar seleção.
  #>
  $i = 0
  $CatalogoProjetores | ForEach-Object {
    $i++
    [pscustomobject]@{
      Id           = $i
      Marca        = $_.Marca
      Modelo       = $_.Modelo
      Qtde         = $_.Qtde
      Descricao    = $_.Descricao
      TipoProjetor = $_.TipoProjetor
      ResolucaoSug = $_.ResolucaoSug
      BrilhoSug    = $_.BrilhoSug
      ContrasteSug = $_.ContrasteSug
      FonteDeLuz   = $_.FonteDeLuz
      VidaLampSug  = $_.VidaLampSug
      DataCompraSug= $_.DataCompraSug
      PrecoCompraSug=$_.PrecoCompraSug
    }
  }
}

function Get-ModeloDoCatalogo {
  <#
    .SYNOPSIS
      Busca um item do catálogo por marca/modelo (case-insensitive) e retorna
      os campos prontos para preencher o formulário.
  #>
  param(
    [Parameter(Mandatory)][string]$Marca,
    [Parameter(Mandatory)][string]$Modelo
  )
  $m = $CatalogoProjetores |
       Where-Object { $_.Marca -ieq $Marca -and $_.Modelo -ieq $Modelo } |
       Select-Object -First 1
  if (-not $m) { return $null }

  [pscustomobject]@{
    Marca          = $m.Marca
    Modelo         = $m.Modelo
    TipoProjetor   = $m.TipoProjetor
    Resolucao      = $m.ResolucaoSug
    BrilhoLumens   = $m.BrilhoSug
    Contraste      = $m.ContrasteSug
    FonteDeLuz     = $m.FonteDeLuz
    VidaLampada    = $m.VidaLampSug
    DataCompraSug  = $m.DataCompraSug
    PrecoCompraSug = $m.PrecoCompraSug
  }
}

function Show-CatalogoRapido {
  <#
    .SYNOPSIS
      Mostra a listagem compacta: Id, Marca, Modelo, Qtde
  #>
  $lista = Get-CatalogoProjetores | Sort-Object Marca, Modelo
  if (-not $lista) { return }
  Write-Host ""
  Write-Host "== Catálogo rápido (opcional) ==" -ForegroundColor Cyan
  $lista | Select-Object Id,Marca,Modelo,Qtde | Format-Table -AutoSize
  Write-Host ""
}

function Select-ModeloDoCatalogoRapido {
  <#
    .SYNOPSIS
      Exibe o catálogo rápido e pergunta o Id. Enter pula.
  #>
  $lista = Get-CatalogoProjetores | Sort-Object Marca, Modelo
  if (-not $lista) { return $null }

  Show-CatalogoRapido
  $sel = Read-Host "Informe o Id (ou Enter para pular)"
  if ([string]::IsNullOrWhiteSpace($sel)) { return $null }
  if ($sel -notmatch '^\d+$') { return $null }

  $escolha = $lista | Where-Object { $_.Id -eq [int]$sel } | Select-Object -First 1
  if (-not $escolha) { return $null }

  return (Get-ModeloDoCatalogo -Marca $escolha.Marca -Modelo $escolha.Modelo)
}
