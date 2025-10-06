#  Contribuindo para Inventários

Obrigado por considerar contribuir! 

---

##  Como contribuir

1. **Faça um fork** do repositório.
2. Crie uma branch para sua modificação:
   ```bash
   git checkout -b minha-feature
   ```
3. Edite ou adicione novos scripts em `scripts/`.
4. Teste localmente no **Windows PowerShell 5.1+**.
5. Certifique-se de manter menus e mensagens em **português (PT-BR)**.
6. Valide os scripts com:
   ```powershell
   Invoke-ScriptAnalyzer -Path .\scripts -Recurse
   ```
7. Abra um **Pull Request** descrevendo suas alterações.

---

##  Padrões do projeto

- Scripts devem sempre gerar **TXT, JSON e CSV** (mesmo que CSV vazio em cadastros).
- Utilizar variáveis padrão já existentes:
  - `$Patrimonio`, `$Local`, `$Responsavel`, `$Estado`, etc.
- Seguir padrão de menus em PT-BR.
- Sempre incluir cabeçalho de comentário no início do script com:
  - Nome, versão, objetivo, compatibilidade.

---

##  Relatando problemas

Se encontrar bugs:
1. Abra uma **Issue**.
2. Inclua a mensagem de erro, versão do Windows e como reproduzir.
3. Sugira melhorias, se possível.

---

##  Licença

Contribuições são aceitas sob a licença **MIT**.  
Isso significa que qualquer colaboração será incorporada e distribuída livremente.
