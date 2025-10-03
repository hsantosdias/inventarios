Describe 'Scripts básicos' {
  It 'scripts/ existe' {
    Test-Path (Join-Path $PWD 'scripts') | Should -BeTrue
  }

  $files = @(
    'scripts/maquinas.ps1',
    'scripts/projetores.ps1',
    'scripts/impressoras.ps1'
  )

  It 'todos os scripts existem' -ForEach $files {
    param($Path)
    Test-Path (Join-Path $PWD $Path) | Should -BeTrue
  }
}
