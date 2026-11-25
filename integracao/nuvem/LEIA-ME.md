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

## Link pessoal (o consultório escolhe o paciente no V Saúde)
1. O médico pede no chat: "manda o Meu humor para o CPF ...".
2. `node link.js <CPF ou nome> [humor|diva]` busca no V Saúde e gera um código aleatório e os links
   (`humor.html?c=CÓDIGO&n=PrimeiroNome`, `diva.html?c=...`).
3. A linha do código vai para a aba **Links** da planilha (código | ID do paciente no V Saúde | nome | app | data).
4. O paciente abre já identificado, sem digitar CPF. A resposta chega com "código ..." e o processador
   anexa direto no cadastro daquele ID ("localizado pelo link pessoal").
Os links sem código continuam funcionando como antes (o paciente informa o CPF).
