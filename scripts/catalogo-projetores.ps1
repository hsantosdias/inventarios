# catalogo-projetores.ps1
# Catálogo de modelos de projetores usados para pré-preenchimento no cadastro
# Compatível com PowerShell 5+
# Dica: Epson = LCD (3LCD); BenQ = DLP. Especificações não preenchidas ficam $null para o operador informar.

# ========= DADOS DO CATÁLOGO =========
$CatalogoProjetores = @(
  # ---------- EPSON (LCD / 3LCD) ----------
  [pscustomobject]@{
    Marca        = "Epson"
    Modelo       = "S41+"
    Descricao    = "PROJETOR XGA X41+ BR 3600 EPSON"
    TipoProjetor = "LCD"           # segura para Epson
    ResolucaoSug = $null           # ex.: "1024x768" se quiser preencher depois
    BrilhoSug    = 3600
    ContrasteSug = $null
    FonteDeLuz   = "Lâmpada"
  }
  [pscustomobject]@{
    Marca="Epson"; Modelo="S8+";  Descricao="Epson S8+";  TipoProjetor="LCD";  ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
  [pscustomobject]@{
    Marca="Epson"; Modelo="S12+"; Descricao="Epson S12+"; TipoProjetor="LCD";  ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
  [pscustomobject]@{
    Marca="Epson"; Modelo="S18+"; Descricao="Epson S18+"; TipoProjetor="LCD";  ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
  [pscustomobject]@{
    Marca="Epson"; Modelo="S31+"; Descricao="Epson S31+"; TipoProjetor="LCD";  ResolucaoSug=$null; BrilhoSug=$null; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }

  # ---------- BENQ (DLP) ----------
  [pscustomobject]@{
    Marca="BenQ"; Modelo="MS550"; Descricao="Projetor BenQ MS550 SVGA 3600 Ansi"; TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=3600; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
  [pscustomobject]@{
    Marca="BenQ"; Modelo="MX611"; Descricao="PROJETOR XGA MX611 BR 4000 BENQ";   TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=4000; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
  [pscustomobject]@{
    Marca="BenQ"; Modelo="MS531"; Descricao="PROJETOR SVGA MS531 BR 3300 BENQ";  TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=3300; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
  [pscustomobject]@{
    Marca="BenQ"; Modelo="MX560"; Descricao="PROJETOR XGA MX560 IM 4000";        TipoProjetor="DLP"; ResolucaoSug=$null; BrilhoSug=4000; ContrasteSug=$null; FonteDeLuz="Lâmpada"
  }
)

# ========= FUNÇÕES DE APOIO =========

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
  <#
    .SYNOPSIS
      Busca um item do catálogo por marca/modelo (case-insensitive).
    .OUTPUTS
      PSCustomObject com campos: Marca, Modelo, TipoProjetor, Resolucao, BrilhoLumens, Contraste, FonteDeLuz
      (apenas os que existirem; demais ficam $null para o operador completar no formulário)
  #>
  param(
    [Parameter(Mandatory)][string]$Marca,
    [Parameter(Mandatory)][string]$Modelo
  )
  $m = $CatalogoProjetores |
       Where-Object { $_.Marca -ieq $Marca -and $_.Modelo -ieq $Modelo } |
       Select-Object -First 1
  if (-not $m) { return $null }

  # Normaliza nomes de campos para encaixar no script de cadastro
  [pscustomobject]@{
    Marca          = $m.Marca
    Modelo         = $m.Modelo
    TipoProjetor   = $m.TipoProjetor      # ex.: "LCD" ou "DLP" (combina com seu menu)
    Resolucao      = $m.ResolucaoSug      # pode vir $null para o operador preencher
    BrilhoLumens   = $m.BrilhoSug
    Contraste      = $m.ContrasteSug
    FonteDeLuz     = $m.FonteDeLuz
  }
}

function Select-ModeloDoCatalogo {
  <#
    .SYNOPSIS
      Mostra um menu de seleção do catálogo e retorna o preset escolhido.
    .DESCRIPTION
      Inclui a opção 0 = Outro modelo (preenchimento manual).
  #>
  $lista = Get-CatalogoProjetores | Sort-Object Marca, Modelo
  if (-not $lista) { return $null }

  Write-Host ""
  Write-Host "== Catálogo de Projetores ==" -ForegroundColor Cyan
  $lista | Format-Table Id,Marca,Modelo,Descricao,TipoProjetor,BrilhoSug -AutoSize
  Write-Host ""
  $sel = Read-Host "Digite o Id do modelo (ou 0 para 'Outro modelo')"

  if ($sel -match '^\s*0\s*$') {
    return $null  # deixa para o operador digitar tudo manualmente
  }

  if ($sel -notmatch '^\d+$') { return $null }
  $escolha = $lista | Where-Object { $_.Id -eq [int]$sel } | Select-Object -First 1
  if (-not $escolha) { return $null }

  return (Get-ModeloDoCatalogo -Marca $escolha.Marca -Modelo $escolha.Modelo)
}
