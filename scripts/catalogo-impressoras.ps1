# scripts/catalogo-impressoras.ps1
# Catálogo padronizado para pré-preenchimento de cadastro de impressoras

function New-PrinterPreset {
  param(
    [Parameter(Mandatory)][string]$Marca,
    [Parameter(Mandatory)][string]$Modelo,
    [Parameter(Mandatory)][string]$TipoImpressora,  # ex.: "Laser (mono)", "Tanque de Tinta", "Multifuncional Laser"
    [Parameter(Mandatory)][string]$TipoCor,         # "Monocromática" | "Colorida"
    [Parameter(Mandatory)][string]$Duplex,          # "Automático" | "Manual" | "Sem duplex"
    [string]$Obs = $null
  )
  [pscustomobject]@{
    Marca           = $Marca.Trim()
    Modelo          = ($Modelo -replace '\s+',' ').Trim()
    TipoImpressora  = $TipoImpressora
    TipoCor         = $TipoCor
    Duplex          = $Duplex
    Observacao      = $Obs
  }
}

# === Catálogo (normalizado do seu levantamento) ===
$CatalogoImpressoras = @(
  New-PrinterPreset -Marca 'Brother' -Modelo 'HL-L1232W'              -TipoImpressora 'Laser (mono)'         -TipoCor 'Monocromática' -Duplex 'Sem duplex' -Obs 'Wi-Fi'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-120'                  -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-3110'                 -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-3150'                 -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex' -Obs 'Wi-Fi'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-3250'                 -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex' -Obs 'Wi-Fi'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-355'                  -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-375'                  -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-396'                  -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Sem duplex' -Obs 'Wi-Fi'
  New-PrinterPreset -Marca 'Epson'   -Modelo 'L-4260'                 -TipoImpressora 'Tanque de Tinta'      -TipoCor 'Colorida'      -Duplex 'Automático'
  New-PrinterPreset -Marca 'HP'      -Modelo '1102w'                  -TipoImpressora 'Laser (mono)'         -TipoCor 'Monocromática' -Duplex 'Sem duplex' -Obs 'LaserJet Pro P1102w'
  New-PrinterPreset -Marca 'Ricoh'   -Modelo 'MPC 300 Multifuncional' -TipoImpressora 'Multifuncional Laser' -TipoCor 'Colorida'      -Duplex 'Automático' -Obs 'Confirmar duplex conforme SKU'
)

# === Frequências (útil para planejamento/compras) ===
$CatalogoContagem = @(
  @{ Marca='Brother'; Modelo='HL-L1232W';          Qtde=1 }
  @{ Marca='Epson';   Modelo='L-120';              Qtde=1 }
  @{ Marca='Epson';   Modelo='L-3110';             Qtde=5 }
  @{ Marca='Epson';   Modelo='L-3150';             Qtde=1 }
  @{ Marca='Epson';   Modelo='L-3250';             Qtde=3 }
  @{ Marca='Epson';   Modelo='L-355';              Qtde=5 }
  @{ Marca='Epson';   Modelo='L-375';              Qtde=3 }
  @{ Marca='Epson';   Modelo='L-396';              Qtde=1 }
  @{ Marca='Epson';   Modelo='L-4260';             Qtde=2 }
  @{ Marca='HP';      Modelo='1102w';              Qtde=2 }
  @{ Marca='Ricoh';   Modelo='MPC 300 Multifuncional'; Qtde=1 }
) | ForEach-Object { [pscustomobject]$_ }

function Get-ModeloDoCatalogo {
  param(
    [Parameter(Mandatory)][string]$Marca,
    [Parameter(Mandatory)][string]$Modelo
  )
  $CatalogoImpressoras | Where-Object {
    $_.Marca  -ieq $Marca.Trim() -and
    $_.Modelo -ieq ($Modelo -replace '\s+',' ').Trim()
  } | Select-Object -First 1
}

function Export-CatalogoInventario {
  param(
    [string]$Destino = (Join-Path $PWD 'artifacts')
  )
  if (-not (Test-Path $Destino)) { New-Item -ItemType Directory -Path $Destino | Out-Null }
  $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
  $json  = Join-Path $Destino ("catalogo-impressoras_{0}.json" -f $stamp)
  $csv   = Join-Path $Destino ("catalogo-contagem_{0}.csv"     -f $stamp)

  $CatalogoImpressoras | ConvertTo-Json -Depth 5 | Out-File -FilePath $json -Encoding UTF8
  $CatalogoContagem    | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csv

  Write-Host "Exportado:" -ForegroundColor Cyan
  Write-Host "  $json"
  Write-Host "  $csv"
}
