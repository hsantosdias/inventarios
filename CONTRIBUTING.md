# Contribuindo

1. Crie uma branch a partir de `main`.
2. Faça commits pequenos e descritivos.
3. Rode `Invoke-ScriptAnalyzer` e os testes Pester localmente, se possível.
4. Abra um PR:
   - Explique o que mudou.
   - Anexe prints dos prompts/saídas, quando aplicável.
   - **Nunca** inclua senhas, caminhos internos sensíveis, IPs reais (use exemplos).
5. Aguarde o CI passar e revisão.

## Estilo
- PowerShell 5+, PT-BR nos prompts.
- Evite `?:` e outras construções não suportadas.
- Mantenha compatibilidade do JSON (chaves existentes).
