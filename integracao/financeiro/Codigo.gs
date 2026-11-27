/**
 * Financeiro do consultório
 * -------------------------
 * Roda dentro de uma planilha PRIVADA do Dr. Luis Felipe ("Financeiro do consultório"),
 * publicada como App da Web. Cada secretária recebe um link com um código secreto
 * e só enxerga a própria clínica; o link do médico abre o painel com tudo somado.
 *
 *   Caixa da secretária: agenda do dia puxada do V Saúde → lança forma de pagamento
 *                        e valor de cada paciente, as saídas do caixa e fecha o dia.
 *   Painel do médico:    entradas por clínica e por forma de pagamento, taxa de sala
 *                        das infusões, custos fixos, resultado do mês e pendências.
 *
 * A separação entre clínicas é feita AQUI, no servidor: a clínica vem do código do link,
 * nunca do que a tela envia.
 *
 * Instalação: veja integracao/financeiro/COMO-INSTALAR.md
 */

const FUSO = "America/Sao_Paulo";
const VS_BASE = "https://public-api.vsaude.com.br/api/services/app";

const CLINICAS = {
  // buscaAtiva: a secretária de Praia Grande faz a busca ativa das três clínicas
  PG: { nome: "Praia Grande", local: "Consultório Dr. Luis Felipe Jardim", unidade: 4176, pasta: "Praia Grande", buscaAtiva: true },
  // descontoAluguel: a Cardiovitta desconta o aluguel aos poucos, em % do que entra das consultas
  SO: { nome: "Sombrio", local: "Clínica Cardiovitta", unidade: 4177, pasta: "Sombrio (Cardiovitta)", descontoAluguel: true },
  TO: { nome: "Torres", local: "Vida Multiclínica", unidade: 5398, pasta: "Torres (Vida Multiclínica)" },
};

const SITUACOES = ["Atendido", "Pagar depois", "Pago antes", "Cortesia", "Faltou", "Desmarcou"];
const REALIZADO = { "Atendido": 1, "Pagar depois": 1, "Pago antes": 1, "Cortesia": 1 };
// "Convênio": sessões pagas pelo convênio (ECT pela Unimed); o dinheiro vem depois, repassado pelo hospital
const FORMAS = ["Dinheiro", "Pix", "Débito", "Crédito", "Convênio"];
const CATEGORIAS_SAIDA = ["Taxa de sala · infusão", "Aluguel pago com o caixa", "Retirada pelo Dr. Luis Felipe", "Faxina", "Material de limpeza", "Material de escritório",
  "Insumos e medicação", "Estacionamento", "Lanche e copa", "Outros"];
const PAGO_COM = ["Dinheiro do caixa", "Pix", "Cartão", "Descontado do repasse", "Outro"];
const PRECO_DESC_ALUGUEL = "Sombrio · desconto do aluguel (% das consultas)";
const CAT_INFUSAO = "Taxa de sala · infusão";
// saem da gaveta mas não são custo novo: o aluguel já está em "Custos fixos"; a retirada é dinheiro do próprio médico
const CAT_ALUGUEL = "Aluguel pago com o caixa", CAT_RETIRADA = "Retirada pelo Dr. Luis Felipe";

const ABAS = {
  "Lançamentos": ["ID", "Data", "Clínica", "Agendamento V Saúde", "Hora", "Paciente", "Procedimento", "Situação",
    "Dinheiro", "Pix", "Débito", "Crédito", "Convênio", "Parcelas", "Total", "Observação", "Atualizado em"],
  "Saídas": ["ID", "Data", "Clínica", "Categoria", "Descrição", "Valor", "Pago com", "Atualizado em"],
  "Fechamentos": ["Data", "Clínica", "Entradas", "Saídas", "Dinheiro esperado na gaveta", "Dinheiro contado",
    "Diferença", "Observação", "Fechado em"],
  // "Início" e "Fim" (mês, ex. 2026-08) são opcionais: custo que só vale a partir de / até um mês
  "Custos fixos": ["Clínica", "Descrição", "Valor mensal", "Observação", "Início", "Fim"],
  "Preços": ["Item", "Valor"],
  // dias em que a secretária não mandou o caixa: consultas finalizadas no V Saúde × preço de tabela
  "Estimado pela agenda": ["Data", "Clínica", "Agendamento V Saúde", "Hora", "Paciente", "Procedimento", "Valor estimado", "Motivo"],
};
const PASTA_FOTOS = "Fechamentos do caixa"; // pasta no Drive do médico; uma subpasta por clínica
const MOEDA = '"R$" #,##0.00';

// ===================== INSTALAÇÃO =====================

/** Rodar UMA vez pelo menu Executar (cria as abas e testa a chave do V Saúde). */
function instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(FUSO);
  Object.keys(ABAS).forEach(nome => {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    if (sh.getLastRow() === 0) {
      const cab = ABAS[nome];
      sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold").setBackground("#F0ECDD");
      sh.setFrozenRows(1);
    }
  });
  // versão nova: coluna "Convênio" logo depois de "Crédito" na aba Lançamentos
  const lsh = ss.getSheetByName("Lançamentos"), cabL = lsh.getRange(1, 1, 1, lsh.getLastColumn()).getValues()[0];
  if (cabL.indexOf("Convênio") < 0 && cabL.indexOf("Crédito") >= 0) {
    const pos = cabL.indexOf("Crédito") + 1;
    lsh.insertColumnAfter(pos); lsh.getRange(1, pos + 1).setValue("Convênio").setFontWeight("bold").setBackground("#F0ECDD");
  }
  // datas como texto (aaaa-mm-dd) para o Google não converter nem trocar o fuso
  ["Lançamentos", "Saídas", "Fechamentos"].forEach(n => ss.getSheetByName(n).getRange("A:B").setNumberFormat("@"));
  ss.getSheetByName("Lançamentos").getRange("I:M").setNumberFormat(MOEDA);
  ss.getSheetByName("Lançamentos").getRange("O:O").setNumberFormat(MOEDA);
  ss.getSheetByName("Saídas").getRange("F:F").setNumberFormat(MOEDA);
  ss.getSheetByName("Fechamentos").getRange("C:G").setNumberFormat(MOEDA);
  ss.getSheetByName("Custos fixos").getRange("C:C").setNumberFormat(MOEDA);
  ss.getSheetByName("Preços").getRange("B:B").setNumberFormat(MOEDA);

  const precos = ss.getSheetByName("Preços");
  if (precos.getLastRow() === 1) precos.getRange(2, 1, 6, 2).setValues([
    ["Consulta · Pix ou dinheiro", 600],
    ["Consulta · Débito", 609],
    ["Consulta · Crédito", 618],
    ["Consulta com desconto", 525],
    ["Infusão de escetamina", 850],
    ["Taxa de sala por infusão (custo)", 250],
  ]);
  if (!linhas_("Preços").some(r => r.Item === PRECO_DESC_ALUGUEL)) precos.appendRow([PRECO_DESC_ALUGUEL, 15]);
  if (!linhas_("Preços").some(r => r.Item === "ECT · repasse do hospital (Unimed)"))
    precos.appendRow(["ECT · repasse do hospital (Unimed)", ""]); // preencha: valor que o hospital repassa por sessão
  const fixos = ss.getSheetByName("Custos fixos");
  if (fixos.getRange(1, 5).getValue() === "") fixos.getRange(1, 5, 1, 2).setValues([["Início", "Fim"]]).setFontWeight("bold").setBackground("#F0ECDD");
  // valores ficam só na planilha privada (o código é público): preencha a coluna "Valor mensal"
  if (fixos.getLastRow() === 1) fixos.getRange(2, 1, 4, 4).setValues([
    ["Praia Grande", "Aluguel · Praia Grande (1 salário mínimo)", "", "Preencha com o salário mínimo vigente"],
    ["Sombrio", "Aluguel · Clínica Cardiovitta", "", "Preencha o valor"],
    ["Torres", "Aluguel · Vida Multiclínica", "", "Preencha o valor"],
    ["Torres", "Condomínio · Vida Multiclínica", "", "Preencha o valor; pode variar a cada mês"],
  ]);

  const padrao = ss.getSheetByName("Sheet1") || ss.getSheetByName("Página1");
  if (padrao && padrao.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(padrao);

  const P = PropertiesService.getScriptProperties();
  // conferência automática todo dia às 18:30 (a função ignora sábado e domingo)
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "conferenciaDiaria").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("conferenciaDiaria").timeBased().atHour(18).nearMinute(30).everyDays(1).inTimezone(FUSO).create();
  Object.keys(CLINICAS).forEach(c => pastaClinica_(c)); // cria as pastas das fotos, se faltarem
  if (!P.getProperty("INICIO")) P.setProperty("INICIO", hoje_()); // pendências só a partir do dia da instalação
  if (!chave_()) throw new Error("Falta guardar a chave: Configurações do projeto → Propriedades do script → VSAUDE_API_KEY");
  vs_("POST", "/HealthCareUnitService/GetAll", {});
  // na primeira instalação, traz o histórico das planilhas antigas (uma vez só)
  if (!P.getProperty("HISTORICO_IMPORTADO")) {
    try { importarHistorico(); importarVidaMulticlinica(["NOVEMBRO"]); P.setProperty("HISTORICO_IMPORTADO", hoje_()); }
    catch (err) { Logger.log("Histórico não importado agora (" + err.message + "). Rode importarHistorico depois."); }
  }
  Logger.log("Instalado. A chave do V Saúde funciona. Agora: Implantar → Nova implantação → App da Web, e depois rode gerarLinks.");
}

/**
 * Cria (ou recria) os códigos secretos e escreve os 4 links na aba "Links".
 * Rodar depois da primeira implantação. Rodar de novo invalida TODOS os links antigos
 * (use se um link vazar ou uma secretária sair).
 */
function gerarLinks() {
  const url = urlApp_();
  if (!url) throw new Error("Implante primeiro: Implantar → Nova implantação → App da Web.");
  const tokens = {}, linhas = [["Quem", "Link (não compartilhe com outras pessoas)"]];
  const novo = () => Utilities.getUuid().replace(/-/g, "").slice(0, 20);
  const med = novo(); tokens[med] = "MEDICO";
  linhas.push(["Dr. Luis Felipe · painel (todas as clínicas)", url + "?k=" + med]);
  Object.keys(CLINICAS).forEach(c => {
    const t = novo(); tokens[t] = c;
    linhas.push(["Secretária · " + CLINICAS[c].nome + " (" + CLINICAS[c].local + ")", url + "?k=" + t]);
  });
  PropertiesService.getScriptProperties().setProperty("TOKENS", JSON.stringify(tokens));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Links") || ss.insertSheet("Links");
  sh.clear(); sh.getRange(1, 1, linhas.length, 2).setValues(linhas); sh.getRange("A1:B1").setFontWeight("bold");
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 640);
  Logger.log(linhas.map(l => l.join(": ")).join("\n"));
}

// ===================== APP DA WEB =====================

function doGet(e) {
  const k = String((e && e.parameter && e.parameter.k) || "");
  let papel = null;
  try { papel = papel_(k); } catch (err) {}
  if (!papel) return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1"><div style="font:16px system-ui;padding:40px 20px;text-align:center;color:#2B2A22">' +
    '<h2 style="font-weight:500">Link inválido</h2><p>Peça o link do caixa ao Dr. Luis Felipe.</p></div>').setTitle("Financeiro");
  const c = String((e.parameter && e.parameter.c) || "");
  const t = HtmlService.createTemplate(arquivo_(papel === "MEDICO" && !CLINICAS[c] ? "Painel" : "Caixa"));
  t.k = k;
  t.clin = papel === "MEDICO" ? c : papel;
  t.medico = papel === "MEDICO";
  t.url = urlApp_();
  t.embed = false;
  return t.evaluate()
    .setTitle(papel === "MEDICO" && !CLINICAS[c] ? "Financeiro · painel" : "Caixa · " + CLINICAS[t.clin].nome)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Caixa de uma clínica para abrir DENTRO do painel do médico (sem trocar de página:
 * navegar de dentro do app para outro endereço do app é barrado pelo Google em alguns navegadores).
 */
function caixaHtml(k, c) {
  if (papel_(k) !== "MEDICO") throw new Error("Só o link do médico abre o caixa de outra clínica.");
  if (!CLINICAS[c]) throw new Error("Clínica inválida.");
  const t = HtmlService.createTemplate(arquivo_("Caixa"));
  t.k = k; t.clin = c; t.medico = true; t.url = urlApp_(); t.embed = true;
  return t.evaluate().getContent();
}

/** URL publicada do app (/exec). Rodando pelo editor o Google devolve a /dev, que só abre para o dono:
 *  nesse caso use a propriedade URL_APP (Configurações do projeto → Propriedades do script). */
function urlApp_() {
  const p = PropertiesService.getScriptProperties().getProperty("URL_APP");
  if (p) return p;
  const u = ScriptApp.getService().getUrl();
  if (!/\/dev$/.test(u)) return u;
  // o Google às vezes devolve a /dev mesmo no app publicado: usa a /exec gravada na aba Links
  try {
    const l = linhas_("Links").map(r => String(r["Link (não compartilhe com outras pessoas)"] || "")).find(x => /\/exec\?/.test(x));
    if (l) return l.split("?")[0];
  } catch (e) {}
  return u;
}

function include(nome) { return arquivo_(nome); }

/** Telas: embutidas (versão de arquivo único, ARQUIVOS_) ou nos arquivos HTML do projeto. */
function arquivo_(nome) {
  if (typeof ARQUIVOS_ !== "undefined" && ARQUIVOS_[nome]) return ARQUIVOS_[nome];
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

// ===================== CAIXA DA SECRETÁRIA =====================

/** Tudo de um dia de uma clínica: agenda do V Saúde, lançamentos, saídas e fechamento. */
function dia(k, c, data) {
  const cod = clinicaDe_(k, c); dataOk_(data);
  let agenda = [], aviso = "", outras = [];
  try {
    const cods = Object.keys(CLINICAS), ags = agendas_(cods.map(c => [CLINICAS[c].unidade, data, data])), por = {};
    cods.forEach((c, i) => { por[c] = ags[i]; });
    const R = redistribuir_(por);
    agenda = R[cod];
    // agendados na unidade desta clínica que, pela grade, foram atendidos em outra (só aviso)
    cods.filter(c => c !== cod).forEach(c => R[c].filter(a => a.unidade === cod)
      .forEach(a => outras.push({ hora: a.hora, paciente: a.paciente, procedimento: a.procedimento, clinica: CLINICAS[c].nome })));
  }
  catch (err) { aviso = "Não consegui ler a agenda do V Saúde agora (" + err.message + "). Lance pelo botão “Paciente fora da agenda”."; }
  return {
    clinica: Object.assign({ cod: cod }, CLINICAS[cod]), data: data, hoje: hoje_(), aviso: aviso, agenda: agenda, outras: outras,
    lancamentos: linhas_("Lançamentos").filter(r => r.Data === data && r["Clínica"] === CLINICAS[cod].nome).map(lancOut_),
    saidas: linhas_("Saídas").filter(r => r.Data === data && r["Clínica"] === CLINICAS[cod].nome).map(saidaOut_),
    fechamento: fechamentoOut_(linhas_("Fechamentos").find(r => r.Data === data && r["Clínica"] === CLINICAS[cod].nome)),
    precos: precos_(), situacoes: SITUACOES, formas: FORMAS, categorias: CATEGORIAS_SAIDA, pagoCom: PAGO_COM,
    fotos: fotosDoDia_(cod, data).map(f => f.getName()),
    aluguel: CLINICAS[cod].descontoAluguel ? descontoAluguel_(CLINICAS[cod].nome, data) : null,
  };
}

/**
 * Sombrio: o aluguel do mês é descontado aos poucos, X% do que entra das consultas no dia (Preços),
 * até completar o valor do aluguel (Custos fixos). Lançado como saída "Aluguel pago com o caixa".
 */
function descontoAluguel_(nome, data) {
  const mes = data.slice(0, 7), pct = num_(precos_()[PRECO_DESC_ALUGUEL]) || 0;
  const fx = linhas_("Custos fixos").find(f => f["Clínica"] === nome && /aluguel/i.test(f["Descrição"]) && fixoNoMes_(f, mes));
  const valorMes = fx ? num_(fx["Valor mensal"]) : 0;
  const sai = linhas_("Saídas").filter(r => r["Clínica"] === nome && String(r.Data).slice(0, 7) === mes && r.Categoria === CAT_ALUGUEL);
  const hoje = sai.filter(r => r.Data === data).reduce((s, r) => s + num_(r.Valor), 0);
  const antes = sai.filter(r => r.Data !== data).reduce((s, r) => s + num_(r.Valor), 0);
  const consultas = linhas_("Lançamentos").filter(r => r["Clínica"] === nome && r.Data === data)
    .reduce((s, r) => { const l = lancOut_(r); return s + l.total - l.valores["Convênio"]; }, 0);
  const falta = Math.max(valorMes - antes, 0);
  return { pct: pct, valorMes: valorMes, consultas: consultas, jaAntes: antes, hoje: hoje, falta: falta,
    sugestao: Math.round(Math.min(consultas * pct / 100, falta) * 100) / 100 };
}

function salvarLancamento(k, c, o) {
  const cod = clinicaDe_(k, c), nome = CLINICAS[cod].nome; dataOk_(o.data);
  const paciente = String(o.paciente || "").trim().slice(0, 120);
  if (!paciente) throw new Error("Informe o nome do paciente.");
  if (SITUACOES.indexOf(o.situacao) < 0) throw new Error("Escolha a situação.");
  const v = {}; FORMAS.forEach(f => { v[f] = REALIZADO[o.situacao] ? num_(o[f]) : 0; });
  const total = FORMAS.reduce((s, f) => s + v[f], 0);
  const parcelas = v["Crédito"] > 0 ? Math.min(Math.max(parseInt(o.parcelas, 10) || 1, 1), 18) : "";
  return comTrava_(() => {
    const sh = aba_("Lançamentos"), todas = linhas_("Lançamentos");
    let alvo = o.id ? todas.find(r => r.ID === o.id) : null;
    if (!alvo && o.agendamento) alvo = todas.find(r => r["Agendamento V Saúde"] === String(o.agendamento));
    if (alvo && alvo["Clínica"] !== nome) throw new Error("Este lançamento é de outra clínica.");
    const id = alvo ? alvo.ID : novoId_();
    const linha = [id, o.data, nome, String(o.agendamento || ""), String(o.hora || ""), paciente,
      String(o.procedimento || "").slice(0, 80), o.situacao, v["Dinheiro"], v["Pix"], v["Débito"], v["Crédito"], v["Convênio"],
      parcelas, total, String(o.obs || "").slice(0, 300), agora_()];
    if (alvo) sh.getRange(alvo._linha, 1, 1, linha.length).setValues([linha]);
    else sh.appendRow(linha);
    return id;
  });
}

function excluirLancamento(k, c, id) { return excluir_("Lançamentos", clinicaDe_(k, c), id); }

function salvarSaida(k, c, o) {
  const cod = clinicaDe_(k, c), nome = CLINICAS[cod].nome; dataOk_(o.data);
  const valor = num_(o.valor);
  if (!(valor > 0)) throw new Error("Informe o valor da saída.");
  const cat = CATEGORIAS_SAIDA.indexOf(o.categoria) >= 0 ? o.categoria : "Outros";
  const pago = PAGO_COM.indexOf(o.pagoCom) >= 0 ? o.pagoCom : "Dinheiro do caixa";
  return comTrava_(() => {
    const sh = aba_("Saídas"), alvo = o.id ? linhas_("Saídas").find(r => r.ID === o.id) : null;
    if (alvo && alvo["Clínica"] !== nome) throw new Error("Esta saída é de outra clínica.");
    const id = alvo ? alvo.ID : novoId_();
    const linha = [id, o.data, nome, cat, String(o.descricao || "").slice(0, 200), valor, pago, agora_()];
    if (alvo) sh.getRange(alvo._linha, 1, 1, linha.length).setValues([linha]); else sh.appendRow(linha);
    return id;
  });
}

function excluirSaida(k, c, id) { return excluir_("Saídas", clinicaDe_(k, c), id); }

/** Fecha (ou refecha) o dia: grava os totais calculados aqui e o dinheiro contado na gaveta. */
function fecharDia(k, c, data, contado, obs) {
  const cod = clinicaDe_(k, c), nome = CLINICAS[cod].nome; dataOk_(data);
  return comTrava_(() => {
    const t = totaisDia_(nome, data);
    const cont = contado === "" || contado == null ? "" : num_(contado);
    const linha = [data, nome, t.entradas, t.saidas, t.esperado, cont, cont === "" ? "" : Math.round((cont - t.esperado) * 100) / 100,
      String(obs || "").slice(0, 300), agora_()];
    const alvo = linhas_("Fechamentos").find(r => r.Data === data && r["Clínica"] === nome), sh = aba_("Fechamentos");
    if (alvo) sh.getRange(alvo._linha, 1, 1, linha.length).setValues([linha]); else sh.appendRow(linha);
    return fechamentoOut_(objDe_(ABAS["Fechamentos"], linha));
  });
}

/** Resumo do mês de UMA clínica (o que a secretária pode ver). mes = "aaaa-mm". */
function mesClinica(k, c, mes) {
  const cod = clinicaDe_(k, c), nome = CLINICAS[cod].nome; mesOk_(mes);
  return Object.assign({ clinica: Object.assign({ cod: cod }, CLINICAS[cod]), mes: mes }, resumoClinica_(nome, mes));
}

/** Recebe a foto do fechamento (já reduzida no celular) e guarda na pasta da clínica no Drive do médico. */
function enviarFoto(k, c, data, dataUrl) {
  const cod = clinicaDe_(k, c); dataOk_(data);
  const m = String(dataUrl || "").match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!m) throw new Error("Envie uma foto (JPG ou PNG).");
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 8 * 1024 * 1024) throw new Error("Foto grande demais.");
  const nome = "Fechamento " + data + " " + Utilities.formatDate(new Date(), FUSO, "HH-mm-ss") + (m[1] === "image/png" ? ".png" : ".jpg");
  pastaClinica_(cod).createFile(Utilities.newBlob(bytes, m[1], nome));
  return nome;
}

function pastaClinica_(cod) {
  const it = DriveApp.getFoldersByName(PASTA_FOTOS), raiz = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
  const sub = raiz.getFoldersByName(CLINICAS[cod].pasta);
  return sub.hasNext() ? sub.next() : raiz.createFolder(CLINICAS[cod].pasta);
}

function fotosDoDia_(cod, data) {
  try {
    const out = [], it = pastaClinica_(cod).getFiles();
    while (it.hasNext()) {
      const f = it.next(), nome = f.getName();
      const dia = (nome.match(/(\d{4}-\d{2}-\d{2})/) || [])[1] || Utilities.formatDate(f.getDateCreated(), FUSO, "yyyy-MM-dd");
      if (dia === data) out.push(f);
    }
    return out;
  } catch (e) { return []; }
}

/**
 * Roda sozinha todo dia às 18:30 (gatilho criado por instalar). Em dias úteis, confere cada clínica
 * (agenda do V Saúde × lançamentos, caixa fechado, foto enviada) e manda o resumo por e-mail ao médico.
 */
function conferenciaDiaria() {
  const hoje = hoje_(), dow = Utilities.formatDate(new Date(), FUSO, "u");
  if (dow === "6" || dow === "7") return;
  const hora = Utilities.formatDate(new Date(), FUSO, "HH:mm");
  const lanc = linhas_("Lançamentos"), fech = linhas_("Fechamentos"), lancados = {};
  lanc.forEach(l => { if (l["Agendamento V Saúde"]) lancados[l["Agendamento V Saúde"]] = 1; lancados[l.Data + "|" + nomeChave_(l.Paciente)] = 1; });
  const cods = Object.keys(CLINICAS);
  let ags = null; try {
    const a = agendas_(cods.map(c => [CLINICAS[c].unidade, hoje, hoje])), por = {};
    cods.forEach((c, i) => { por[c] = a[i]; }); const R = redistribuir_(por); ags = cods.map(c => R[c]);
  } catch (e) {}
  const blocos = [];
  let algum = false;
  cods.forEach((cod, i) => {
    const C = CLINICAS[cod], ag = ags ? ags[i].filter(a => a.hora <= hora) : [];
    const doDia = lanc.filter(l => l.Data === hoje && l["Clínica"] === C.nome).map(lancOut_);
    if (!ag.length && !doDia.length) { blocos.push("<b>" + C.nome + "</b>: sem atendimento hoje."); return; }
    algum = true;
    const f = {}; FORMAS.forEach(x => { f[x] = 0; }); doDia.forEach(l => FORMAS.forEach(x => { f[x] += l.valores[x]; }));
    const total = FORMAS.reduce((s, x) => s + f[x], 0);
    const sem = ag.filter(a => a.finalizado && !lancados[a.id] && !lancados[a.data + "|" + nomeChave_(a.paciente)]);
    const faltasSem = ag.filter(a => a.status === 6 && !lancados[a.id] && !lancados[a.data + "|" + nomeChave_(a.paciente)]);
    const fechado = fech.some(x => x.Data === hoje && x["Clínica"] === C.nome), fotos = fotosDoDia_(cod, hoje).length;
    const li = [];
    li.push("Recebido: <b>" + brl_(total) + "</b> (" + FORMAS.filter(x => f[x]).map(x => x + " " + brl_(f[x])).join(" · ") + ") · " + doDia.length + " lançamentos");
    if (sem.length) li.push("⚠ Atendidos no V Saúde sem lançamento: " + sem.map(a => a.hora + " " + a.paciente.split(" ")[0]).join(", "));
    if (faltasSem.length) li.push("Faltas marcadas no V Saúde sem lançamento: " + faltasSem.map(a => a.paciente.split(" ")[0]).join(", "));
    li.push(fechado ? "Caixa fechado ✓" : "⚠ Caixa ainda não fechado");
    li.push(fotos ? "Foto do fechamento recebida ✓ (" + fotos + ")" : "Sem foto do fechamento");
    blocos.push("<b>" + C.nome + "</b><ul><li>" + li.join("</li><li>") + "</li></ul>");
  });
  if (!ags) blocos.push("<i>Não consegui ler a agenda do V Saúde agora; conferi só a planilha.</i>");
  const quem = Session.getEffectiveUser().getEmail();
  const html = (algum ? blocos.join("") : "Nenhum atendimento hoje nas três clínicas.") +
    (function () { try { const A = anual_(hoje.slice(0, 4)); return "<p>Acumulado de " + A.ano + ": <b>" + brl_(A.total) + "</b> (média " + brl_(A.media) + "/mês" +
      (A.estimado ? "; inclui " + brl_(A.estimado) + " estimado pela agenda" : "") + ")</p>"; } catch (e) { return ""; } })() +
    '<p><a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">Abrir a planilha</a></p>';
  MailApp.sendEmail({ to: quem, subject: "Caixa do dia " + hoje.slice(8, 10) + "/" + hoje.slice(5, 7) + " · conferência das 18:30", htmlBody: html });
}

function brl_(v) { return "R$ " + (Math.round(v * 100) / 100).toFixed(2).replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, "."); }

// ===================== PAINEL DO MÉDICO =====================

function painel(k, mes) {
  if (papel_(k) !== "MEDICO") throw new Error("Só o link do médico abre o painel.");
  mesOk_(mes);
  const fixos = linhas_("Custos fixos"), precos = precos_(), taxa = precos["Taxa de sala por infusão (custo)"] || 0;
  const hoje = hoje_(), pend = [], clinicas = [];
  const horaAgora = Utilities.formatDate(new Date(), FUSO, "HH:mm");
  const inicio = PropertiesService.getScriptProperties().getProperty("INICIO") || "2000-01-01";
  // um paciente conta como lançado pelo agendamento OU pelo nome no mesmo dia, em qualquer clínica
  // (consultas de 1ª vez às vezes são agendadas no V Saúde na unidade errada)
  const lancados = {};
  linhas_("Lançamentos").forEach(l => {
    if (l["Agendamento V Saúde"]) lancados[l["Agendamento V Saúde"]] = 1;
    lancados[l.Data + "|" + nomeChave_(l.Paciente)] = 1;
  });
  // 3 agendas por clínica, todas em paralelo: o mês escolhido, os últimos 180 dias e de hoje em diante
  const cods = Object.keys(CLINICAS), pedidos = [];
  cods.forEach(cod => { const u = CLINICAS[cod].unidade;
    pedidos.push([u, mes + "-01", fimMes_(mes)], [u, somaDiasStr_(hoje, -180), hoje], [u, hoje, "9999-12-31"]); });
  let ags = null, erroAg = "";
  try { ags = agendas_(pedidos); } catch (err) { erroAg = err.message; }
  let doMes = null;
  if (ags) { const por = {}; cods.forEach((c, i) => { por[c] = ags[3 * i]; }); doMes = redistribuir_(por); }
  const reinc = [];
  cods.forEach((cod, i) => {
    const C = CLINICAS[cod], r = resumoClinica_(C.nome, mes, true);
    const fx = fixos.filter(f => f["Clínica"] === C.nome && fixoNoMes_(f, mes)).map(f => ({ descricao: f["Descrição"], valor: f["Valor mensal"] === "" ? null : num_(f["Valor mensal"]) }));
    const totFixos = fx.reduce((s, f) => s + (f.valor || 0), 0);
    const taxaDevida = r.infusoes * taxa;
    const saidasOutras = r.saidas - r.taxaPagaCaixa - r.aluguelPagoCaixa - r.retiradas;
    const custos = saidasOutras + taxaDevida + totFixos;
    clinicas.push(Object.assign({ cod: cod, local: C.local, nome: C.nome }, r, {
      fixos: fx, totFixos: totFixos, taxaDevida: taxaDevida, taxaPorInfusao: taxa,
      saidasOutras: saidasOutras, custos: custos, resultado: r.entradas - custos,
    }));
    // pendências: agenda do V Saúde até hoje sem lançamento
    const agenda = doMes ? doMes[cod] : [];
    if (!ags) pend.push({ clinica: C.nome, data: "", paciente: "", motivo: "Agenda do V Saúde indisponível: " + erroAg });
    clinicas[clinicas.length - 1].faltasVS = faltas_(agenda, hoje, horaAgora, precos);
    if (ags) reinc.push.apply(reinc, reincidentes_(ags[3 * i + 1], ags[3 * i + 2], hoje, horaAgora, C.nome));
    agenda.filter(a => (a.data < hoje || (a.data === hoje && a.hora <= horaAgora)) && a.data >= inicio
        && !lancados[a.id] && !lancados[a.data + "|" + nomeChave_(a.paciente)]).forEach(a => pend.push({
      clinica: C.nome, data: a.data, hora: a.hora, paciente: a.paciente, procedimento: a.procedimento,
      motivo: a.status === 6 ? "V Saúde marcou falta · sem lançamento" : a.finalizado ? "Atendido no V Saúde · sem lançamento" : "Na agenda · sem lançamento",
    }));
    r.lista.filter(l => l.situacao === "Pagar depois").forEach(l => pend.push({
      clinica: C.nome, data: l.data, hora: l.hora, paciente: l.paciente, procedimento: l.procedimento, motivo: "A receber (pagar depois)" }));
    r.dias.filter(d => !d.fechado && d.data < hoje && d.data >= inicio).forEach(d => pend.push({
      clinica: C.nome, data: d.data, paciente: "", motivo: "Dia sem fechamento de caixa" }));
    delete clinicas[clinicas.length - 1].lista;
  });
  const soma = campo => clinicas.reduce((s, x) => s + x[campo], 0);
  const formas = {}; FORMAS.concat([FORMA_EST]).forEach(f => { formas[f] = clinicas.reduce((s, x) => s + (x.formas[f] || 0), 0); });
  pend.sort((a, b) => (a.data + (a.hora || "")).localeCompare(b.data + (b.hora || "")));
  return {
    mes: mes, hoje: hoje, inicio: inicio, planilha: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    total: { entradas: soma("entradas"), custos: soma("custos"), resultado: soma("resultado"), atendimentos: soma("atendimentos"),
      infusoes: soma("infusoes"), formas: formas, estimado: soma("estimado"), nEstimado: soma("nEstimado"),
      faltasVS: clinicas.reduce((s, x) => s + x.faltasVS.faltas, 0), atendidosVS: clinicas.reduce((s, x) => s + x.faltasVS.atendidos, 0),
      perdida: clinicas.reduce((s, x) => s + x.faltasVS.perdida, 0) },
    clinicas: clinicas, pendencias: pend, buscaAtiva: reinc, ano: anual_(mes.slice(0, 4)),
  };
}

// ===================== CÁLCULOS =====================

/**
 * Entradas do ano mês a mês: lançamentos diários + aba "Histórico mensal" (totais de meses
 * anteriores ao sistema, só para dias que não têm lançamento diário).
 */
function anual_(ano) {
  const meses = [];
  for (let m = 1; m <= 12; m++) meses.push({ mes: ano + "-" + ("0" + m).slice(-2), total: 0, estimado: 0, porClinica: {} });
  const soma = (mes, clinica, v) => {
    const M = meses[+mes.slice(5, 7) - 1]; if (!M || !v) return;
    M.total += v; M.porClinica[clinica] = (M.porClinica[clinica] || 0) + v;
  };
  linhas_("Lançamentos").forEach(r => { if (String(r.Data).slice(0, 4) === ano) soma(String(r.Data).slice(0, 7), r["Clínica"], lancOut_(r).total); });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Histórico mensal");
  if (sh) linhas_("Histórico mensal").forEach(r => { const m = String(r["Mês"]); if (m.slice(0, 4) === ano) soma(m.slice(0, 7), r["Clínica"], num_(r.Recebido)); });
  estimados_().forEach(e => { if (e.data.slice(0, 4) === ano) { soma(e.data.slice(0, 7), e.clinica, e.valor); meses[+e.data.slice(5, 7) - 1].estimado += e.valor; } });
  meses.forEach(M => { M.total = Math.round(M.total * 100) / 100; });
  const comValor = meses.filter(M => M.total > 0);
  const total = comValor.reduce((s, M) => s + M.total, 0);
  return { ano: ano, meses: meses, total: total, estimado: meses.reduce((s, M) => s + M.estimado, 0), media: comValor.length ? total / comValor.length : 0, nMeses: comValor.length };
}

function resumoClinica_(nome, mes, comEstimado) {
  const lanc = linhas_("Lançamentos").filter(r => r["Clínica"] === nome && String(r.Data).slice(0, 7) === mes).map(lancOut_);
  const sai = linhas_("Saídas").filter(r => r["Clínica"] === nome && String(r.Data).slice(0, 7) === mes).map(saidaOut_);
  const fech = {}; linhas_("Fechamentos").forEach(f => { if (f["Clínica"] === nome) fech[f.Data] = fechamentoOut_(f); });
  const formas = {}; FORMAS.forEach(f => { formas[f] = 0; });
  const dias = {};
  const d = x => dias[x] || (dias[x] = { data: x, entradas: 0, saidas: 0, atendimentos: 0, faltas: 0 });
  let atend = 0, infus = 0, faltas = 0, cortesias = 0, aReceber = 0;
  lanc.forEach(l => {
    FORMAS.forEach(f => { formas[f] += l.valores[f]; });
    const D = d(l.data); D.entradas += l.total;
    if (REALIZADO[l.situacao]) { atend++; D.atendimentos++; if (ehInfusao_(l.procedimento)) infus++; }
    if (l.situacao === "Faltou") { faltas++; D.faltas++; }
    if (l.situacao === "Cortesia") cortesias++;
    if (l.situacao === "Pagar depois") aReceber++;
  });
  // estimativa pela agenda (só no painel do médico): entra como forma "Estimado pela agenda"
  let estimado = 0, nEst = 0;
  if (comEstimado) estimados_().filter(e => e.clinica === nome && e.data.slice(0, 7) === mes).forEach(e => {
    estimado += e.valor; nEst++; atend++; if (ehInfusao_(e.procedimento)) infus++;
    const D = d(e.data); D.entradas += e.valor; D.estimado = (D.estimado || 0) + e.valor; D.atendimentos++;
  });
  if (estimado) formas[FORMA_EST] = estimado;
  let saidas = 0, taxaPaga = 0, aluguelCaixa = 0, retiradas = 0; const porCat = {};
  sai.forEach(s => {
    saidas += s.valor; d(s.data).saidas += s.valor; porCat[s.categoria] = (porCat[s.categoria] || 0) + s.valor;
    if (s.categoria === CAT_INFUSAO) taxaPaga += s.valor;
    if (s.categoria === CAT_ALUGUEL) aluguelCaixa += s.valor;
    if (s.categoria === CAT_RETIRADA) retiradas += s.valor;
  });
  const lista = Object.keys(dias).sort().map(x => Object.assign(dias[x], { fechado: !!fech[x], diferenca: fech[x] ? fech[x].diferenca : null }));
  return { entradas: FORMAS.reduce((s, f) => s + formas[f], 0) + estimado, formas: formas, estimado: estimado, nEstimado: nEst, atendimentos: atend, infusoes: infus, faltas: faltas,
    cortesias: cortesias, aReceber: aReceber, saidas: saidas, taxaPagaCaixa: taxaPaga, aluguelPagoCaixa: aluguelCaixa, retiradas: retiradas, saidasPorCategoria: porCat, dias: lista, lista: lanc };
}

const FORMA_EST = "Estimado pela agenda";

/** Mês "AAAA-MM" de uma célula de Início/Fim (aceita 2026-08, 08/2026 ou uma data). */
function mesCelula_(x) {
  const t = String(x || "").trim(); let m;
  if ((m = t.match(/^(\d{4})-(\d{2})/))) return m[1] + "-" + m[2];
  if ((m = t.match(/^(?:\d{1,2}\/)?(\d{1,2})\/(\d{4})/))) return m[2] + "-" + ("0" + m[1]).slice(-2);
  return "";
}
function fixoNoMes_(f, mes) {
  const ini = mesCelula_(f["Início"]), fim = mesCelula_(f.Fim);
  return (!ini || ini <= mes) && (!fim || fim >= mes);
}

/**
 * Aba "Estimado pela agenda": consultas finalizadas no V Saúde em dias sem caixa da secretária,
 * a preço de tabela. Se depois chegar o caixa real daquele dia (qualquer lançamento da clínica na data),
 * a estimativa do dia deixa de contar sozinha.
 */
function estimados_() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estimado pela agenda")) return [];
  // repasse de convênio (ECT) não é o caixa da secretária: não substitui a estimativa do dia
  const real = {}; linhas_("Lançamentos").forEach(r => { const l = lancOut_(r);
    if (l.total - l.valores["Convênio"] > 0 || l.valores["Convênio"] === 0) real[r["Clínica"] + "|" + r.Data] = 1; });
  return linhas_("Estimado pela agenda").map(r => ({ data: String(r.Data), clinica: r["Clínica"], hora: String(r.Hora || ""), paciente: r.Paciente,
    procedimento: r.Procedimento, valor: num_(r["Valor estimado"]), motivo: r.Motivo }))
    .filter(e => /^\d{4}-\d\d-\d\d$/.test(e.data) && !real[e.clinica + "|" + e.data]);
}

function totaisDia_(nome, data) {
  const lanc = linhas_("Lançamentos").filter(r => r.Data === data && r["Clínica"] === nome).map(lancOut_);
  const sai = linhas_("Saídas").filter(r => r.Data === data && r["Clínica"] === nome).map(saidaOut_);
  const entradas = lanc.reduce((s, l) => s + l.total, 0), dinheiro = lanc.reduce((s, l) => s + l.valores["Dinheiro"], 0);
  const saidas = sai.reduce((s, x) => s + x.valor, 0), saidasDinheiro = sai.filter(x => x.pagoCom === "Dinheiro do caixa").reduce((s, x) => s + x.valor, 0);
  return { entradas: entradas, saidas: saidas, esperado: Math.round((dinheiro - saidasDinheiro) * 100) / 100 };
}

function nomeChave_(n) {
  return String(n || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
    .replace(/\b(de|da|do|das|dos|e)\b/g, " ").replace(/[^a-z ]/g, " ").replace(/\s+/g, " ").trim();
}

function ehInfusao_(p) { return /infus|escetamin|cetamin/i.test(String(p || "")); }

// ===================== V SAÚDE =====================

function chave_() { return PropertiesService.getScriptProperties().getProperty("VSAUDE_API_KEY"); }

function vs_(metodo, caminho, corpo) {
  const op = { method: metodo, muteHttpExceptions: true, headers: { "VSAUDE-API-KEY": chave_() } };
  if (corpo !== undefined) { op.contentType = "application/json"; op.payload = JSON.stringify(corpo); }
  const r = UrlFetchApp.fetch(VS_BASE + caminho, op);
  let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
  if (r.getResponseCode() >= 300 || !j || j.success === false)
    throw new Error("V Saúde " + r.getResponseCode() + ": " + ((j && j.error && j.error.message) || r.getContentText().slice(0, 150)));
  return j.result;
}

const STATUS_FINALIZADO = { 8: 1, 81: 1, 82: 1, 90: 1 };
const STATUS_CANCELADO = { 4: 1, 40: 1, 41: 1, 5: 1, 50: 1, 51: 1, 7: 1 };
function statusTxt_(s) {
  if (STATUS_FINALIZADO[s]) return "Finalizado";
  if (s === 6) return "Faltou";
  if (s === 3 || s === 30 || s === 31) return "Confirmado";
  if (s === 2 || s === 20 || s === 21) return "Remarcado";
  return "Agendado";
}

/**
 * Agendamentos com paciente de uma unidade entre duas datas (aaaa-mm-dd, inclusive), no horário de Brasília.
 * A API devolve tudo a partir de "from" (o "range" não corta), por isso o filtro de datas é feito aqui.
 */
/**
 * Grade de atendimento do médico: em que clínica ele está em cada dia/horário.
 * Usada só para corrigir agendamentos feitos na unidade de Praia Grande (onde caem as consultas de 1ª vez
 * marcadas na unidade errada). Agendamentos nas unidades de Sombrio e Torres ficam sempre na própria clínica.
 * Dia da semana (1 = segunda … 5 = sexta): uma clínica o dia todo, ou [[a partir de, clínica], ...].
 */
const GRADE = {
  1: [["00:00", "PG"], ["13:00", "TO"]],   // segunda: manhã Praia Grande, tarde Torres
  2: "SO",                                 // terça: Sombrio
  3: "TO",                                 // quarta: Torres
  4: "SO",                                 // quinta: Sombrio
  5: [["00:00", "PG"], ["13:00", "TO"]],   // sexta: manhã Praia Grande, tarde Torres
};
function clinicaDaGrade_(data, hora) {
  const g = GRADE[new Date(data + "T12:00:00").getDay()];
  if (!g) return "";
  if (typeof g === "string") return g;
  let c = ""; g.forEach(x => { if (hora >= x[0]) c = x[1]; });
  return c;
}
/** Em que clínica o atendimento acontece (infusão, ECT e consulta online ficam na unidade agendada). */
function clinicaReal_(a, codUnidade) {
  if (codUnidade !== "PG" || a.online || ehInfusao_(a.procedimento) || /\bECT\b/.test(a.procedimento)) return codUnidade;
  return clinicaDaGrade_(a.data, a.hora) || codUnidade;
}
/** {PG:[...],SO:[...],TO:[...]} por unidade → por clínica real; cada item ganha "unidade" (cod da unidade agendada). */
function redistribuir_(porUnidade) {
  const out = {}; Object.keys(CLINICAS).forEach(c => { out[c] = []; });
  Object.keys(porUnidade).forEach(u => (porUnidade[u] || []).forEach(a => {
    const c = clinicaReal_(a, u); (out[c] || out[u]).push(Object.assign({}, a, { unidade: u, unidadeNome: CLINICAS[u].nome }));
  }));
  Object.keys(out).forEach(c => out[c].sort((x, y) => (x.data + x.hora).localeCompare(y.data + y.hora)));
  return out;
}

function agenda_(unidade, de, ate) { return agendas_([[unidade, de, ate]])[0]; }

/** Várias consultas de agenda em paralelo: pedidos = [[unidade, de, ate], ...]. */
function agendas_(pedidos) {
  const reqs = pedidos.map(p => ({ url: VS_BASE + "/ScheduleService/GetAll", method: "post", muteHttpExceptions: true,
    headers: { "VSAUDE-API-KEY": chave_() }, contentType: "application/json",
    payload: JSON.stringify({ from: p[1] + "T03:00:00Z", range: 3, careUnits: [p[0]], showCanceledAppointments: false }) }));
  return UrlFetchApp.fetchAll(reqs).map((r, i) => {
    let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
    if (r.getResponseCode() >= 300 || !j || j.success === false)
      throw new Error("V Saúde " + r.getResponseCode() + ": " + ((j && j.error && j.error.message) || r.getContentText().slice(0, 150)));
    return mapAgenda_(j.result || [], pedidos[i][0], pedidos[i][1], pedidos[i][2]);
  });
}

/** A API devolve ~6 meses a partir de "from" (o "range" não corta), por isso o filtro de datas é feito aqui. */
function mapAgenda_(res, unidade, de, ate) {
  return res.filter(a => a.patient && a.careUnit && Number(a.careUnit.id) === unidade && !STATUS_CANCELADO[a.status])
    .map(a => {
      const dt = new Date(a.date);
      return { id: String(a.id), data: Utilities.formatDate(dt, FUSO, "yyyy-MM-dd"), hora: Utilities.formatDate(dt, FUSO, "HH:mm"),
        pacienteId: String(a.patient.id || ""), tel: String(a.patient.phoneNumber || ""),
        paciente: String(a.patient.name || "").replace(/\s+/g, " ").trim(), procedimento: String((a.procedure && a.procedure.name) || "").replace(/\s+/g, " ").trim(),
        status: a.status, statusTxt: statusTxt_(a.status), finalizado: !!STATUS_FINALIZADO[a.status],
        // o V Saúde nem sempre marca "remotely": a consulta online vem no nome do procedimento
        online: !!a.remotely || /online/i.test((a.procedure && a.procedure.name) || "") };
    })
    .filter(a => a.data >= de && a.data <= ate && !/bloqueio|reuni|palestra/i.test(a.procedimento) && !/programa r[áa]dio|r[áa]dio \d/i.test(a.paciente))
    .sort((x, y) => (x.data + x.hora).localeCompare(y.data + y.hora));
}

function tipoAtend_(proc) { return ehInfusao_(proc) ? "Infusão" : /1ª|1a vez|primeira/i.test(proc) ? "1ª vez" : "Acompanhamento"; }

/** Taxa de faltas (pela agenda do V Saúde) dos agendamentos que já passaram. */
function faltas_(agenda, hoje, horaAgora, precos) {
  const passou = a => a.data < hoje || (a.data === hoje && a.hora <= horaAgora);
  const tipos = {}, out = { atendidos: 0, faltas: 0, perdida: 0, porTipo: [] };
  agenda.filter(passou).forEach(a => {
    const t = tipoAtend_(a.procedimento), T = tipos[t] || (tipos[t] = { tipo: t, atendidos: 0, faltas: 0 });
    if (a.finalizado) { T.atendidos++; out.atendidos++; }
    if (a.status === 6) {
      T.faltas++; out.faltas++;
      out.perdida += t === "Infusão" ? (precos["Infusão de escetamina"] || 0) : (precos["Consulta · Pix ou dinheiro"] || 0);
    }
  });
  out.porTipo = ["Acompanhamento", "1ª vez", "Infusão"].filter(t => tipos[t]).map(t => tipos[t]);
  out.taxa = out.atendidos + out.faltas ? out.faltas / (out.atendidos + out.faltas) : 0;
  return out;
}

/**
 * Busca ativa: pacientes com 2+ faltas nos últimos 180 dias.
 * hist = agenda dos últimos 180 dias; fut = agenda de hoje em diante.
 */
function reincidentes_(hist, fut, hoje, horaAgora, clinica) {
  const passou = a => a.data < hoje || (a.data === hoje && a.hora <= horaAgora);
  const P = {};
  hist.filter(passou).forEach(a => {
    const k = a.pacienteId || nomeChave_(a.paciente);
    const p = P[k] || (P[k] = { paciente: a.paciente, tel: a.tel, faltas: 0, atendidos: 0, ultimaFalta: "", seq: [] });
    if (a.status === 6) { p.faltas++; p.ultimaFalta = a.data; p.seq.push("F"); }
    else if (a.finalizado) { p.atendidos++; p.seq.push("A"); }
  });
  const prox = {};
  fut.filter(a => !passou(a)).forEach(a => { const k = a.pacienteId || nomeChave_(a.paciente); if (!prox[k]) prox[k] = a.data + " " + a.hora; });
  return Object.keys(P).filter(k => P[k].faltas >= 2).map(k => {
    const p = P[k], s = p.seq.join("");
    return { clinica: clinica, paciente: p.paciente, tel: p.tel.replace(/\D/g, ""), faltas: p.faltas, atendidos: p.atendidos,
      ultimaFalta: p.ultimaFalta, faltouUltima: /F$/.test(s), faltouDuasUltimas: /FF$/.test(s), proxima: prox[k] || "" };
  }).sort((x, y) => (+y.faltouDuasUltimas - +x.faltouDuasUltimas) || (+!y.proxima - +!x.proxima) || (y.faltas - x.faltas) || y.ultimaFalta.localeCompare(x.ultimaFalta));
}

/** Lista de busca ativa de UMA clínica (tela da secretária). */
function buscaAtiva(k, c) {
  const cod = clinicaDe_(k, c), C = CLINICAS[cod], hoje = hoje_(), hora = Utilities.formatDate(new Date(), FUSO, "HH:mm");
  if (!C.buscaAtiva && papel_(k) !== "MEDICO") throw new Error("A busca ativa é feita pela secretária de Praia Grande.");
  // Praia Grande (Gabriela) vê as três clínicas; o médico, abrindo o caixa de outra clínica, vê só aquela
  const cods = C.buscaAtiva ? Object.keys(CLINICAS) : [cod], pedidos = [];
  cods.forEach(x => { const u = CLINICAS[x].unidade; pedidos.push([u, somaDiasStr_(hoje, -180), hoje], [u, hoje, "9999-12-31"]); });
  const ag = agendas_(pedidos), lista = [];
  cods.forEach((x, i) => lista.push.apply(lista, reincidentes_(ag[2 * i], ag[2 * i + 1], hoje, hora, CLINICAS[x].nome)));
  lista.sort((x, y) => (+y.faltouDuasUltimas - +x.faltouDuasUltimas) || (+!y.proxima - +!x.proxima) || (y.faltas - x.faltas) || y.ultimaFalta.localeCompare(x.ultimaFalta));
  return { clinica: Object.assign({ cod: cod }, C), todas: cods.length > 1, lista: lista };
}

function somaDiasStr_(d, n) {
  const x = new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10) + n));
  return x.toISOString().slice(0, 10);
}

// ===================== PLANILHA =====================

function aba_(nome) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh) throw new Error("Falta a aba " + nome + ". Rode a função instalar.");
  return sh;
}

function linhas_(nome) {
  const sh = aba_(nome), v = sh.getDataRange().getValues(), cab = v[0] || [];
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(x => x === "")) continue;
    const o = objDe_(cab, v[i]); o._linha = i + 1; out.push(o);
  }
  return out;
}

function objDe_(cab, linha) {
  const o = {};
  cab.forEach((h, j) => {
    let x = linha[j];
    if (x instanceof Date) x = h === "Data" ? Utilities.formatDate(x, FUSO, "yyyy-MM-dd") : Utilities.formatDate(x, FUSO, "dd/MM/yyyy HH:mm");
    o[h] = x === undefined ? "" : typeof x === "number" ? x : String(x);
  });
  return o;
}

function lancOut_(r) {
  const valores = {}; FORMAS.forEach(f => { valores[f] = num_(r[f]); });
  return { id: r.ID, data: r.Data, agendamento: String(r["Agendamento V Saúde"] || ""), hora: String(r.Hora || ""), paciente: r.Paciente,
    procedimento: r.Procedimento, situacao: r["Situação"], valores: valores, parcelas: r.Parcelas === "" ? "" : Number(r.Parcelas),
    total: FORMAS.reduce((s, f) => s + valores[f], 0), obs: r["Observação"] };
}

function saidaOut_(r) {
  return { id: r.ID, data: r.Data, categoria: r.Categoria, descricao: r["Descrição"], valor: num_(r.Valor), pagoCom: r["Pago com"] };
}

function fechamentoOut_(r) {
  if (!r) return null;
  return { data: r.Data, entradas: num_(r.Entradas), saidas: num_(r["Saídas"]), esperado: num_(r["Dinheiro esperado na gaveta"]),
    contado: r["Dinheiro contado"] === "" ? null : num_(r["Dinheiro contado"]), diferenca: r["Diferença"] === "" ? null : num_(r["Diferença"]),
    obs: r["Observação"], fechadoEm: r["Fechado em"] };
}

function precos_() {
  const p = {}; linhas_("Preços").forEach(r => { if (r.Item) p[r.Item] = num_(r.Valor); }); return p;
}

function excluir_(abaNome, cod, id) {
  return comTrava_(() => {
    const alvo = linhas_(abaNome).find(r => r.ID === id);
    if (!alvo) return false;
    if (alvo["Clínica"] !== CLINICAS[cod].nome) throw new Error("Este registro é de outra clínica.");
    aba_(abaNome).deleteRow(alvo._linha);
    return true;
  });
}

// ===================== AUXILIARES =====================

function papel_(k) {
  const t = JSON.parse(PropertiesService.getScriptProperties().getProperty("TOKENS") || "{}");
  const p = /^[a-f0-9]{20}$/.test(String(k || "")) ? t[k] : null;
  if (!p) throw new Error("Link inválido. Peça um novo link ao Dr. Luis Felipe.");
  return p;
}

/** A clínica SEMPRE sai do código do link; só o médico pode escolher outra. */
function clinicaDe_(k, c) {
  const p = papel_(k);
  if (p !== "MEDICO") return p;
  if (!CLINICAS[c]) throw new Error("Clínica desconhecida.");
  return c;
}

function comTrava_(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function num_(x) {
  if (typeof x === "number") return isFinite(x) ? Math.round(x * 100) / 100 : 0;
  let s = String(x || "").replace(/[^\d,.-]/g, "");
  if (!s) return 0;
  if (s.indexOf(",") >= 0) s = s.replace(/\./g, "").replace(",", ".");          // 1.234,56
  else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, "");             // 1.234
  const n = parseFloat(s);
  return isFinite(n) && n > 0 ? Math.round(n * 100) / 100 : 0;
}

function dataOk_(d) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d || ""))) throw new Error("Data inválida."); }
function mesOk_(m) { if (!/^\d{4}-\d{2}$/.test(String(m || ""))) throw new Error("Mês inválido."); }
function hoje_() { return Utilities.formatDate(new Date(), FUSO, "yyyy-MM-dd"); }
function agora_() { return Utilities.formatDate(new Date(), FUSO, "dd/MM/yyyy HH:mm"); }
function novoId_() { return Utilities.getUuid().replace(/-/g, "").slice(0, 10); }
function fimMes_(mes) {
  const a = +mes.slice(0, 4), m = +mes.slice(5, 7), ult = new Date(Date.UTC(a, m, 0)).getUTCDate();
  return mes + "-" + ("0" + ult).slice(-2);
}

// ===================== HISTÓRICO =====================

/**
 * Opcional, rodar UMA vez: copia para a aba Lançamentos o histórico das planilhas
 * "Pagamentos - Praia Grande", "Pagamentos - Sombrio" e "Pagamentos - Torres" (formato padronizado de 31/10/2026).
 * Linhas sem data ficam de fora e aparecem no registro de execução.
 */
function importarHistorico() {
  const FONTES = { "Praia Grande": "1XAuELsXRIlfCc4DjqJW1qwOBezJw8h_s2q_XAneSzSo",
    "Sombrio": "1Ke8wxv7VX-F-rM2wHPbDxlqTNT6oGkbNvxzUzZRQhlA", "Torres": "1iEV5rTYcQBXX6qs0kkhEKt5GDg8kgm_pDBVvafS89Zk" };
  const SIT = { "Atendido": "Atendido", "Falta": "Faltou", "Remarcado": "Desmarcou", "Cortesia": "Cortesia",
    "Pago antecipado": "Pago antes", "Sem registro": "Pagar depois" };
  const ja = {}; linhas_("Lançamentos").forEach(r => { ja[[r.Data, r["Clínica"], r.Paciente].join("|")] = 1; });
  const novas = [], fora = [];
  Object.keys(FONTES).forEach(nome => {
    const v = SpreadsheetApp.openById(FONTES[nome]).getSheetByName("Lançamentos").getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const [d, turno, pac, modal, sit, din, pix, cart, tipo, parc, obs] = v[i];
      if (!pac) continue;
      let data = d instanceof Date ? Utilities.formatDate(d, FUSO, "yyyy-MM-dd") : String(d || "").replace(/^(\d{2})\/(\d{2})\/(\d{4})$/, "$3-$2-$1");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) { fora.push(nome + ": " + pac); continue; }
      if (ja[[data, nome, String(pac).trim()].join("|")]) continue;
      const c = num_(cart), deb = /d[ée]bito/i.test(tipo) ? c : 0, cred = c - deb;
      const vals = [num_(din), num_(pix), deb, cred];
      const proc = vals.some(x => x === 850) ? "Infusão de escetamina" : (modal === "Online" ? "Consulta online" : "Consulta");
      novas.push([novoId_(), data, nome, "", String(turno || ""), String(pac).trim(), proc, SIT[sit] || "Atendido",
        vals[0], vals[1], vals[2], vals[3], 0, cred > 0 ? (parseInt(parc, 10) || 1) : "", vals.reduce((s, x) => s + x, 0),
        ("Importado da planilha antiga. " + (obs || "")).trim(), agora_()]);
    }
  });
  if (novas.length) aba_("Lançamentos").getRange(aba_("Lançamentos").getLastRow() + 1, 1, novas.length, novas[0].length).setValues(novas);
  Logger.log(novas.length + " lançamentos importados." + (fora.length ? "\nSem data (não importados):\n" + fora.join("\n") : ""));
}

/**
 * Opcional: importa a planilha que a própria Vida Multiclínica mantém ("Pagamentos Dr. Luis Felipe"),
 * um bloco por dia. Por padrão só a aba NOVEMBRO (abril a outubro já vêm por importarHistorico).
 * Ex.: importarVidaMulticlinica(["NOVEMBRO", "DEZEMBRO"]). Não duplica quem já foi lançado no mesmo dia.
 */
function importarVidaMulticlinica(abas) {
  const ID = "1P4LINpbjM3vw6_H8r8ioKOkRKqsEZObPiEkoviQWI4Y", nome = "Torres";
  abas = abas && abas.length ? abas : ["NOVEMBRO"];
  const ja = {}; linhas_("Lançamentos").forEach(r => { if (r["Clínica"] === nome) ja[r.Data + "|" + nomeChave_(r.Paciente)] = 1; });
  const novas = [], avisos = [], ss = SpreadsheetApp.openById(ID);
  abas.forEach(aba => {
    const sh = ss.getSheetByName(aba); if (!sh) { avisos.push("Aba não encontrada: " + aba); return; }
    novas.push.apply(novas, linhasVida_(sh.getDataRange().getDisplayValues(), ja, avisos));
  });
  if (novas.length) aba_("Lançamentos").getRange(aba_("Lançamentos").getLastRow() + 1, 1, novas.length, novas[0].length).setValues(novas);
  Logger.log(novas.length + " lançamentos importados da Vida Multiclínica." + (avisos.length ? "\nConferir:\n" + avisos.join("\n") : ""));
}

function linhasVida_(v, ja, avisos) {
  const out = []; let data = "", turno = "";
  const PULAR = /^(paciente|dinheiro|pix|cart[aã]o|total:?)$/i;
  v.forEach(r => {
    const a = String(r[0] || "").trim(), b = String(r[1] || "").trim(), c = String(r[2] || "").trim(), d = String(r[3] || "").trim();
    const m = c.match(/DATA:\s*(\d{2})\/(\d{2})\/(\d{4})/i);
    if (m) { data = m[3] + "-" + m[2] + "-" + m[1]; turno = ""; return; }
    if (d) turno = d;
    if (!a || !data || PULAR.test(a)) return;
    if (ja[data + "|" + nomeChave_(a)]) return;
    const p = pagamentoVida_(b, c);
    if (p.aviso) avisos.push(data + " · " + a + ": " + p.aviso);
    const vals = FORMAS.map(f => p.vals[f] || 0), total = vals.reduce((s, x) => s + x, 0);
    out.push([novoId_(), data, "Torres", "", turno, a, total === 850 ? "Infusão de escetamina" : "Consulta", p.sit,
      vals[0], vals[1], vals[2], vals[3], vals[4], p.vals["Crédito"] ? p.parcelas : "", total,
      ("Importado da planilha da Vida Multiclínica: " + [b, c].filter(String).join(" · ")).slice(0, 300), agora_()]);
    ja[data + "|" + nomeChave_(a)] = 1;
  });
  return out;
}

/** Lê "R$ 618,00" + "400,00 crédito 2x/218,00 no pix" → valores por forma de pagamento. */
function pagamentoVida_(valorTxt, formaTxt) {
  const txt = (valorTxt + " " + formaTxt).toLowerCase();
  if (/n[aã]o veio|faltou|internad/.test(txt)) return { sit: "Faltou", vals: {}, parcelas: "" };
  const total = num_(valorTxt);
  if (!total) return { sit: "Pagar depois", vals: {}, parcelas: "", aviso: "sem valor na planilha da clínica" };
  const forma = s => /pix|dep[oó]sito|transfer/.test(s) ? "Pix" : /dinheiro|esp[eé]cie/.test(s) ? "Dinheiro" : /d[eé]bito/.test(s) ? "Débito" : /cr[eé]dito|cart/.test(s) ? "Crédito" : "";
  const vals = {}; let parcelas = 1, aviso = "";
  const parc = formaTxt.toLowerCase().match(/(\d+)\s*x/); if (parc) parcelas = +parc[1];
  const partes = String(formaTxt || "").toLowerCase().split(/\/|\+| e /).map(x => x.trim()).filter(String);
  if (partes.length > 1) {
    partes.forEach(pt => { const f = forma(pt), n = num_((pt.match(/[\d.,]+/) || [""])[0].replace(/^\d+x$/, "")); if (f && n) vals[f] = (vals[f] || 0) + n; });
    const soma = Object.keys(vals).reduce((s, f) => s + vals[f], 0);
    if (Math.abs(soma - total) > 0.01) aviso = "soma das formas (" + soma + ") diferente do valor (" + total + ")";
  } else {
    const f = forma(String(formaTxt).toLowerCase());
    if (f) vals[f] = total; else { vals["Pix"] = total; aviso = "forma de pagamento não reconhecida, lançado como Pix"; }
  }
  return { sit: "Atendido", vals: vals, parcelas: parcelas, aviso: aviso };
}
