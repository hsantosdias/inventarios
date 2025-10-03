# Inventários (PowerShell)

Scripts de inventário e cadastro em PT-BR:

- `scripts/inventario-ti.ps1` – inventário de ativos de TI (coleta do Windows).
- `scripts/projetores.ps1` – cadastro de projetores (sem coleta do PC).
- `scripts/impressoras.ps1` – cadastro de impressoras (sem coleta do PC).

Cada script gera:
- **TXT** (resumo humano),
- **JSON** (estrutura padronizada),
- **CSV** (compatibilidade; vazio nos cadastros),
- e rótulos TXT rápidos.

Saída padrão: `C:\Temp\Inventario\<PASTA_DO_DIA>`.

> Opcional: cópia para compartilhamento SMB com **usuário/senha** informados e limpeza de cache/sessões (via `New-PSDrive`, `net use`, `cmdkey`).

---

## Requisitos

- Windows PowerShell 5.1 (ou superior) no Windows 10/11.
- Execução de scripts permitida (veja abaixo).
- Acesso ao compartilhamento SMB, se a cópia estiver habilitada.

### Execução de scripts (desbloqueio)

Se baixar os `.ps1` da internet, o Windows marca como “bloqueado”. Antes de executar:

```powershell
Unblock-File -Path .\scripts\*.ps1
# e opcionalmente:
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
