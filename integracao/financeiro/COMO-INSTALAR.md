# Financeiro do consultório: como instalar

Um link para cada secretária (Praia Grande, Sombrio/Cardiovitta, Torres/Vida Multiclínica) e um link seu, com o painel.
Tudo fica numa **planilha privada sua** no Google Drive. Nenhum valor passa pelo site público do GitHub.

| Link | O que abre | O que vê |
|---|---|---|
| Secretária · Praia Grande | Caixa do dia | Só a agenda e os valores de Praia Grande |
| Secretária · Sombrio | Caixa do dia | Só a Cardiovitta |
| Secretária · Torres | Caixa do dia | Só a Vida Multiclínica |
| Dr. Luis Felipe | Painel do mês | As três clínicas somadas, custos, resultado e pendências; abre o caixa de qualquer uma |

A clínica é decidida **no servidor**, pelo código do link. Mesmo alguém mexendo na página não consegue ler outra clínica.

Faça isto uma vez, no computador. Leva uns 15 minutos.

## 1. Criar a planilha e colar o programa
1. Abra a planilha **Financeiro do consultório** no seu Google Drive (se ainda não existir: **Novo → Planilhas Google**
   com esse nome). Não compartilhe com ninguém.
2. Menu **Extensões → Apps Script**.
3. Abra o arquivo **`Financeiro-completo.gs`** (é um arquivo só, com tudo dentro), selecione tudo e copie.
   Pelo navegador: [Financeiro-completo.gs (texto puro)](https://raw.githubusercontent.com/drluisfelipejardimpsiquiatra/regua-sintomas/claude/affectionate-ramanujan-7ucttm/integracao/financeiro/Financeiro-completo.gs)
   → Ctrl+A → Ctrl+C.
4. No editor do Apps Script, apague o que houver em `Código.gs`, cole (Ctrl+V) e clique em **Salvar** (disquete).

(Para quem for mexer no código: `Codigo.gs`, `Caixa.html`, `Painel.html` e `Estilo.html` são as fontes;
`node montar-arquivo-unico.js` gera de novo o `Financeiro-completo.gs`.)

## 2. Chave do V Saúde e fuso
Engrenagem **Configurações do projeto**:
- **Fuso horário**: *(GMT-03:00) Horário de Brasília*;
- **Propriedades do script → Adicionar propriedade**: nome `VSAUDE_API_KEY`, valor = a mesma chave usada na pré-consulta.

## 3. Instalar
No editor, escolha a função **instalar** → **Executar** → autorize (sua conta → Avançado → Acessar → Permitir).
Deve aparecer: *"Instalado. A chave do V Saúde funciona."*. A planilha ganha as abas:

- **Lançamentos**: um paciente por linha (Dinheiro, Pix, Débito e Crédito em colunas separadas; pagamento misto na mesma linha).
- **Saídas**: o que saiu do caixa (taxa de sala da infusão, faxina, material…).
- **Fechamentos**: o fechamento de cada dia, com o dinheiro contado na gaveta e a diferença.
- **Custos fixos**: aluguel de Praia Grande (1 salário mínimo), aluguel da Cardiovitta, aluguel e condomínio da Vida Multiclínica.
  **Preencha a coluna "Valor mensal"** (os valores não ficam no código, que é público). O condomínio pode variar: atualize quando o boleto mudar.
  Pode acrescentar linhas (ex.: contador, internet); o painel soma tudo da clínica.
  As colunas **Início** e **Fim** (mês, ex. `2026-08`) são opcionais: um custo com Início só conta a partir daquele mês
  (e, com Fim, até aquele mês). Sem elas, o custo vale para todos os meses.
- **Preços**: Pix/dinheiro R$ 600, débito R$ 609, crédito R$ 618, infusão R$ 850 e **taxa de sala por infusão R$ 250**.
  São os valores que a tela sugere; a secretária pode mudar caso a caso. Mudou o preço, mude aqui.

## 4. Publicar
1. **Implantar → Nova implantação** → engrenagem → **App da Web**.
2. Descrição: `Financeiro`. **Executar como: Eu**. **Quem pode acessar: Qualquer pessoa**.
   ("Qualquer pessoa" só abre a página; sem o código secreto do link ela mostra "Link inválido".)
3. **Implantar**.
4. Volte ao editor, escolha **gerarLinks** → **Executar**. A aba **Links** da planilha mostra os 4 links.
5. Mande a cada secretária **só o link dela**. O seu, guarde no seu painel de aplicativos.

## 5. Histórico
Na primeira execução de **instalar** o histórico é importado sozinho (uma vez só; não duplica se rodar de novo).
As funções abaixo continuam disponíveis para rodar à mão.
A função **importarHistorico** copia para a aba Lançamentos o que já está nas planilhas
"Pagamentos - Praia Grande", "Pagamentos - Sombrio" e "Pagamentos - Torres" (padronizadas em 31/10/2026).
Linhas sem data ficam de fora e aparecem no registro de execução. Rode uma vez só.
Para Torres, rode também **importarVidaMulticlinica**: lê direto a planilha que a própria Vida mantém
("Pagamentos Dr. Luis Felipe"), aba NOVEMBRO, e entende pagamentos mistos como "400,00 crédito 2x/218,00 no pix".
Pacientes sem valor na planilha da clínica entram como "Pagar depois" e aparecem nas pendências.
Praia Grande e Sombrio de novembro precisam ser lançados pelo caixa ou à mão.

## Como a secretária usa
1. Abre o link: aparecem os pacientes do dia **na agenda do V Saúde daquela clínica**.
2. Em cada um, **Lançar** → situação (Atendido, Pagar depois, Pago antes, Cortesia, Faltou, Desmarcou) →
   forma de pagamento (pode marcar mais de uma, ex.: R$ 200 em dinheiro + R$ 418 no crédito 3x) → **Salvar**.
   O valor já vem sugerido pela tabela de preços.
3. Paciente que não está na agenda: **+ Paciente fora da agenda**.
4. **Saídas do dia**: taxa de sala da infusão paga no dia, faxina, material, com "pago com" (dinheiro do caixa, Pix…).
5. **Fechar o dia**: confere o total por forma e digita o dinheiro contado na gaveta; a tela mostra a diferença.
6. Aba **Mês**: os dias do mês da clínica dela, com o que está fechado ou aberto.

## Como o painel calcula
- **Entradas** = tudo o que foi lançado (por clínica e por forma de pagamento).
- **Taxa de sala das infusões** = nº de infusões realizadas × R$ 250, calculada sozinha.
  Se a secretária pagou a taxa com o dinheiro do caixa e lançou como saída "Taxa de sala · infusão",
  o painel mostra "já paga pelo caixa" e **não conta duas vezes**.
- **Custos** = custos fixos do mês + taxa das infusões + outras saídas do caixa.
- **Resultado** = entradas − custos. O mês corrente parece pior do que é, porque o custo fixo entra inteiro desde o dia 1.
- **Pendências** (a partir do dia da instalação): paciente que está na agenda do V Saúde, com horário já passado, sem lançamento
  (diferenciando "atendido no V Saúde" de "V Saúde marcou falta"), pacientes "pagar depois" e dias sem fechamento.
  Os valores de cartão são brutos (sem a taxa da maquininha).
- **Unidade errada no V Saúde**: algumas consultas de 1ª vez são agendadas na unidade de Praia Grande mesmo quando o
  atendimento é em Torres. O painel considera o paciente lançado se houver lançamento com o mesmo nome no mesmo dia em
  qualquer clínica, então isso não vira pendência falsa. Na tela de Torres esse paciente não aparece na agenda:
  a secretária lança por "+ Paciente fora da agenda". O ideal é agendar sempre na unidade certa.

## Atualizar o programa depois
Cole o código novo e vá em **Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar**.
Assim os links continuam os mesmos.

## Se um link vazar ou uma secretária sair
Rode **gerarLinks** de novo: todos os links antigos param de funcionar e a aba Links mostra os novos.

## Faltas e busca ativa
- **Painel → "Faltas no mês"**: pela agenda do V Saúde (o que foi marcado como "não compareceu"), por clínica e por tipo
  (acompanhamento, 1ª vez, infusão), com a receita potencial perdida (faltas × preço de tabela).
- **Busca ativa** (no painel, todas as clínicas; na aba "Busca ativa" de cada secretária, só a clínica dela):
  pacientes com **2 ou mais faltas nos últimos 6 meses**. Vem primeiro quem faltou nas 2 últimas e quem está
  **sem retorno marcado**. O botão **WhatsApp** abre a conversa com uma mensagem pronta e acolhedora.
  Quem já tem a próxima consulta marcada aparece com a data: confirmar a presença na véspera
  (e, se for a política do consultório, pedir o Pix antecipado).
- As 9 consultas ao V Saúde do painel são feitas em paralelo; a página leva alguns segundos para abrir.

## Foto do fechamento e conferência das 18:30
- No **Fechamento do dia**, a secretária toca em **"📷 Enviar foto do fechamento"**. A foto é reduzida no celular
  e salva na pasta da clínica dentro de **"Fechamentos do caixa"**, no **Drive do médico** (o programa roda com a conta dele:
  a secretária não precisa de conta Google nem de login). O nome do arquivo leva a data: `Fechamento AAAA-MM-DD HH-mm-ss.jpg`.
- **instalar** cria as pastas (se faltarem) e o gatilho **conferenciaDiaria**, que roda todo dia às 18:30 e, em dias úteis,
  manda ao médico um e-mail por clínica: total recebido por forma de pagamento, atendidos no V Saúde sem lançamento,
  faltas sem lançamento, caixa fechado ou não, foto recebida ou não.
- Ao instalar, o Google pede duas autorizações novas: **Drive** (guardar as fotos) e **e-mail** (mandar a conferência).

## Dias sem caixa: estimativa pela agenda
- A aba **Estimado pela agenda** guarda, para os dias em que nenhuma secretária mandou o caixa, as consultas
  **finalizadas no V Saúde** a preço de tabela (consulta R$ 600, infusão R$ 850; ECT, retorno e cortesia entram com R$ 0).
- O painel soma essa estimativa às entradas (mês, clínica e acumulado do ano) e a mostra **hachurada**, separada do que é real.
  Infusões estimadas entram também na taxa de sala.
- Quando o caixa real daquele dia for lançado (qualquer lançamento da clínica na data), a estimativa do dia **sai sozinha**.
- As secretárias não veem a estimativa.
- Pacientes agendados na unidade errada são atribuídos à clínica onde o médico estava naquele dia
  (terça e quinta, Sombrio; dia com caixa de outra clínica, essa clínica).

## O que muda em cada link de secretária
- **Todas**: lançam os pacientes do dia (entradas), as **saídas** do caixa e fecham o dia, com foto.
- **Praia Grande**: tem a aba **Busca ativa** com os pacientes faltosos das **três clínicas** (é ela quem faz a busca).
  Os links de Sombrio e Torres não têm essa aba.
- **Sombrio**: tem o quadro **Desconto do aluguel**. A Cardiovitta desconta o aluguel aos poucos: X% das consultas do dia
  (Preços → "Sombrio · desconto do aluguel (% das consultas)") até completar o aluguel do mês (Custos fixos).
  O desconto vira uma saída "Aluguel pago com o caixa", pago com "Descontado do repasse": não mexe no dinheiro da gaveta
  e o painel não conta o aluguel duas vezes.
- **Torres**: aluguel e condomínio já estão nos custos fixos; não há desconto no caixa.

## Paciente agendado na unidade errada: a grade de atendimento
Consultas de 1ª vez costumam ser marcadas na unidade de Praia Grande mesmo quando o atendimento é em Sombrio ou Torres.
O programa usa a grade do médico (constante `GRADE` no código) para pôr cada agendamento da unidade de Praia Grande
no caixa da clínica certa: segunda e sexta de manhã, Praia Grande; segunda e sexta à tarde e quarta, Torres;
terça e quinta, Sombrio. Infusão, ECT e consulta online ficam na unidade agendada.
Agendamentos feitos nas unidades de Sombrio e Torres ficam sempre na própria clínica.
O paciente aparece no caixa certo com a etiqueta "agendado em Praia Grande", e o caixa de Praia Grande mostra
um aviso com os que foram para outra clínica. Se a grade mudar num dia (feriado, troca de agenda), a secretária
lança pelo "+ Paciente fora da agenda". Se a grade mudar de vez, ajuste `GRADE` no código.
