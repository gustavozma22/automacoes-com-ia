# Roadmap

O que ainda não foi provado em uso real e o que eu melhoraria. Cada item vira uma *issue* no GitHub.

## Limites conhecidos (honestidade sobre o que foi testado)

- [ ] **Recibo sem documento** nunca rodou em produção. Os outros três tipos de NFS-e/recibo, sim.
- [ ] **NFS-e com vários pedidos** nunca rodou, nem a rolagem de grade que ela usa. O código existe e passa no linter, mas linter não prova sequência de telas.
- [ ] **NFS-e com retenção de imposto na fonte** é retida por decisão: os campos dessa aba ainda não foram mapeados.
- [ ] **Local de negócio** derrubou a primeira tentativa e não voltou a aparecer. O campo está mapeado e vazio de propósito.

## Melhorias de engenharia

- [ ] Portar a validação de código de barras (módulo 10/11 e fator de vencimento) para Python, com testes, e usar o mesmo algoritmo nos dois lados.
- [ ] Teste de contrato entre a planilha de lote e as constantes de coluna dos scripts: falhar no CI se uma coluna mudar de lugar.
- [ ] Resolver os avisos do PSScriptAnalyzer (verbos aprovados, `Write-Host`, `catch` vazio).
- [ ] Modo "simulação" no executor de NFCom: percorre as telas e simula, sem gravar.
- [ ] README em inglês.
