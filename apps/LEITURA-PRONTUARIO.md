# Leitura do prontuário no vSaúde (sessão do Claude no computador do médico)

A API pública do vSaúde não tem prontuário nem prescrições. Esta leitura é feita pelo
Chrome do computador do Dr. Luis Felipe, com o vSaúde logado. **Somente leitura**: nunca
clicar em salvar, editar, nova prescrição, imprimir ou assinar. Nada de dados clínicos em
git, logs ou mensagens fora dos destinos abaixo (LGPD).

## 1. Medicação para o Monitor de Peso

Painel: https://claude.ai/artifact/Q52YEmjLCTgEvgeamHoB6J

1. `ArtifactData list` em `data/users/me/peso/sugestoes` (url do painel). Pendentes são os
   documentos com `pedido` preenchido e `lidoEm` vazio. O `doc_id` é o id do paciente no vSaúde.
2. Para cada pendente, abrir `https://app.vsaude.com.br/patient/<doc_id>`, aba **Prescrições**,
   e pegar a prescrição mais recente.
3. Montar `med` só com os medicamentos ligados ao peso, a compulsão alimentar ou ao
   metabolismo (GLP-1/GIP, orlistate, sibutramina, naltrexona+bupropiona, topiramato,
   lisdexanfetamina, metformina…), no formato `Nome (marca) dose frequência`, separados por
   " + ". Ex.: `Tirzepatida (Mounjaro) 5 mg/semana, SC`.
4. `ArtifactData get` no documento e depois `set` (com `if_version`) mantendo os campos e
   acrescentando: `med`, `data` (AAAA-MM-DD da prescrição), `obs` (titulação prevista ou
   dúvida, curto) e `lidoEm` (milissegundos de agora). Sem medicação desse tipo: `med` = "" e
   `obs` = "Sem medicação para peso na última prescrição".

O médico aprova ou recusa a sugestão no painel; nada é aplicado sem ele.

## 2. Receitas para o Mural de Renovação

Mural: https://claude.ai/artifact/PHeQx6dXHnNFdQ47XSXFqW · coleção `data/rx/receitas`,
`doc_id` = id do paciente. Formato e regras no pedido enviado pela sessão
"Painel de renovação de receitas" e em `mural-receitas/README.md`.
