# Mural de Renovação de Receitas

Painel (Claude Artifact, tipo Data) que mostra quais pacientes precisam de renovação de receita
antes da próxima consulta, considerando o teto de 60 dias da SNCR.

- `painel.html` — código do painel (fonte da verdade do cálculo; o `resumo.mjs` reutiliza o primeiro `<script>`).
- `sync-vsaude.mjs` — lê a agenda na API pública do vSaúde (somente leitura) e gera o JSON gravado em `data/rx/sync/vsaude`.
- `resumo.mjs` — gera o texto da notificação (quem entrou hoje no mural).

Rotina: tarefa agendada 5x ao dia roda `sync-vsaude.mjs`, grava o resultado no mural e notifica.
Receitas detalhadas (medicamento, quantidade, posologia) não existem na API: são lidas da aba
Prescrições do vSaúde pelo navegador e gravadas em `data/rx/receitas/<id-do-paciente>` com
`{consulta, ate, itens: "[{medicamento, quantidade, posologia, dias, controlado}]", revisar}`.

LGPD: nenhum dado de paciente é versionado aqui. O repositório é público.
