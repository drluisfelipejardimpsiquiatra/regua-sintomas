# Integração V Saúde pela nuvem (sem instalar nada na planilha)

Uma tarefa agendada do Claude roda **de hora em hora**:
1. lê a planilha "Pré-consulta por voz · recebimento (respostas)" pelo conector do Google Sheets;
2. roda `processar.js`, que usa o mesmo programa da planilha (`../vsaude-preconsulta.gs`):
   localiza o paciente no V Saúde (CPF; ou nome + nascimento), gera o PDF e anexa nos Arquivos;
3. escreve o resultado na coluna D, "Status V Saúde";
4. manda por e-mail os avisos (sinal de risco; paciente não encontrado).

Só processa respostas a partir de 25/11/2026 que ainda não tenham ✅ ou ⛔ no status.
Para não tentar mais uma linha, escreva ⛔ no começo do status.
Se um dia o programa for instalado na própria planilha (COMO-INSTALAR.md), desligue esta tarefa para não anexar em dobro.

Teste local sem anexar nada: `node processar.js linhas.json saida.json --simular`
