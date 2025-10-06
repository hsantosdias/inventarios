# Inventários (PowerShell)

Sistema em **PowerShell** para inventário e cadastro de equipamentos de TI, com foco em **máquinas, projetores e impressoras**.
Gera relatórios em **TXT, JSON e CSV**, além de rótulos rápidos para identificação de cada equipamento.

---

## Estrutura do Projeto

```
inventarios/
├── scripts/
│   ├── maquinas.ps1      # Inventário automático de máquinas Windows
│   ├── projetores.ps1    # Cadastro manual de projetores
│   └── impressoras.ps1   # Cadastro manual de impressoras
├── README.md
└── CONTRIBUTING.md
```

---

## Funcionalidades

- Coleta automática de informações de **máquinas Windows** (CPU, RAM, disco, rede, SO, BitLocker, etc).
- Cadastro manual de **projetores** e **impressoras** com menus interativos em PT-BR.
- Geração de relatórios:
  - **TXT** (resumo humano),
  - **JSON** (estrutura padronizada),
  - **CSV** (compatibilidade; vazio nos cadastros manuais).
- Rótulos TXT rápidos (`nome-modelo-data.txt`).
- **Cópia opcional** para compartilhamento SMB com usuário/senha.
- **Depreciação linear** de equipamentos (cálculo opcional).

---

## Como usar

1. **Clone o repositório**:

   ```powershell
   git clone https://github.com/hsantosdias/inventarios.git
   cd inventarios\scripts
   ```
2. **Habilite execução de scripts** (se necessário):

   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   Unblock-File -Path .\*.ps1
   ```
3. **Execute o script desejado**:

   ```powershell
   .\maquinas.ps1
   .\projetores.ps1
   .\impressoras.ps1
   ```
4. **Saída padrão**:

   ```
   C:\Temp\Inventario\<TIPO>_<NOME>_<DATA>
   ```

---

## Exemplo de saída (máquinas)

```
==== Inventário de Máquina ====
Data/Hora: 2025-10-03 12:00:00
Hostname: LAB-PC01
Usuário: dominio\usuario
Patrimônio: 2025-00123
Local: Sala de Robótica
Responsável: Maria Silva
Condição: Bom
...
```

##### Scripts de inventário e cadastro em PT-BR:

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

- **Windows PowerShell 5.1+** (Windows 10/11).
- Acesso ao compartilhamento SMB (se ativado).
- Permissão para execução de scripts.
- Acesso ao compartilhamento SMB, se a cópia estiver habilitada.

---

### Execução de scripts (desbloqueio)

Se baixar os `.ps1` da internet, o Windows marca como “bloqueado”. Antes de executar:

```powershell
Unblock-File -Path .\scripts\*.ps1
# e opcionalmente:
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

```

**Se ainda não funcionar**

Alguns Windows 11 vêm com modo restrito (“AllSigned”) aplicado por política.
Nesse caso, apenas para rodar nessa sessão atual, use:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\inventario-ti-v5.ps1
```

## Contribuindo

Veja o arquivo [CONTRIBUTING.md](CONTRIBUTING.md) para mais detalhes.

---

## Licença

Distribuído sob a licença **MIT**.
Você pode usar, modificar e distribuir livremente, mantendo os créditos.
