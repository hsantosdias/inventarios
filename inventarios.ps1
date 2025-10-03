<#
Menu Principal - Inventários
Executa scripts de cadastro/inventário de acordo com a escolha.
Compatibilidade: PowerShell 5+ (Windows 10/11)
#>

# Caminho base dos scripts
$ScriptsPath = "scripts"

function Show-Menu {
    Clear-Host
    Write-Host "========================================"
    Write-Host "   Inventários - Menu Principal"
    Write-Host "========================================"
    Write-Host "1) Cadastro de Máquinas"
    Write-Host "2) Cadastro de Projetores"
    Write-Host "3) Cadastro de Impressoras"
    Write-Host "4) Sair"
    Write-Host "========================================"
}

do {
    Show-Menu
    $choice = Read-Host "Digite a opção desejada (1-4)"

    switch ($choice) {
        "1" {
            Write-Host "Executando cadastro de máquinas..." -ForegroundColor Cyan
            & "$ScriptsPath\maquinas.ps1"
            Pause
        }
        "2" {
            Write-Host "Executando cadastro de projetores..." -ForegroundColor Cyan
            & "$ScriptsPath\projetores.ps1"
            Pause
        }
        "3" {
            Write-Host "Executando cadastro de impressoras..." -ForegroundColor Cyan
            & "$ScriptsPath\impressoras.ps1"
            Pause
        }
        "4" {
            Write-Host "Saindo do sistema." -ForegroundColor Yellow
        }
        Default {
            Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red
            Pause
        }
    }
} while ($choice -ne "4")
