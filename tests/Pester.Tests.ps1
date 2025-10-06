# tests/Pester.Tests.ps1  (compatível Pester v3/v4 - ASCII-safe)

$RepoRoot   = Split-Path -Parent $PSScriptRoot
$ScriptsDir = Join-Path $RepoRoot 'scripts'

$ExpectedScripts = @(
  'maquinas.ps1',
  'projetores.ps1',
  'impressoras.ps1'
)

Describe 'Inventarios - Estrutura de Scripts' {

  Context 'Diretorio base' {
    It 'Pasta "scripts" deve existir' {
      Test-Path $ScriptsDir | Should Be $true
    }
  }

  Context 'Arquivos obrigatorios' {
    $cases = $ExpectedScripts | ForEach-Object { @{ Name = $_ } }
    It 'Script <Name> deve existir' -TestCases $cases {
      param($Name)
      $full = Join-Path $ScriptsDir $Name
      Test-Path $full | Should Be $true
    }
  }

  Context 'Validacao de conteudo' {
    $cases = $ExpectedScripts | ForEach-Object { @{ Name = $_ } }
    It 'Script <Name> deve iniciar com comentario (# ou <#)' -TestCases $cases {
      param($Name)
      $full = Join-Path $ScriptsDir $Name

      $firstNonEmpty = $null
      if (Test-Path $full) {
        try {
          # Lê as primeiras linhas; ignora vazias e remove BOM se houver
          $lines = Get-Content -Path $full -TotalCount 10 -ErrorAction Stop
          $firstNonEmpty = $lines | Where-Object { $_ -match '\S' } | Select-Object -First 1
          if ($firstNonEmpty) { $firstNonEmpty = ($firstNonEmpty -replace "^\uFEFF","") }
        } catch {}
      }

      $firstNonEmpty | Should Not BeNullOrEmpty
      # Aceita linha que comece com "#" (comentario de linha) OU "<#" (comentario de bloco)
      $firstNonEmpty | Should Match '^\s*(#|<#)'
    }
  }
}
